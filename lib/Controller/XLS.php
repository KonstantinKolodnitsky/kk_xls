<?php
/**
 *Created by Konstantin Kolodnitsky
 * Date: 02.10.13
 * Time: 10:16
 */
namespace KonstantinKolodnitsky\kk_xls;
require_once __DIR__.'/../../vendor/phpexcel/PHPExcel.php';
class Controller_XLS extends \AbstractController{
    public $properties = array(
        'creator' => 'ATK4 addon kk_xls',
        'lastModifiedBy' => 'kk_xls',
        'title' => 'ATK4 addon kk_xls',
        'subject' => 'xls data',
        'description' => 'This file has been generated via kk_xls addon for ATK4',
        'keywords' => 'atk4, agiletoolkit, addons, kk_xls',
        'category' => 'data'
    );
    function init(){
        parent::init();

        // add add-on locations to pathfinder
        $l = $this->api->locate('addons',__NAMESPACE__,'location');
        $addon_location = $this->api->locate('addons',__NAMESPACE__);
        $this->api->pathfinder->addLocation($addon_location,array(
            'php'=>array('lib','vendor')
        ))->setParent($l);
    }
    function getXLS($m, $properties, $fields = null, $fields_width = null, $count_totals = null){
        /*Checking data*/
        if(is_subclass_of($m,'SQL_Model')){
            $data=$m->getRows();
        }elseif(is_array($m)){
            if(array_diff_key($m[0],array_keys(array_keys($m[0])))){
                $data = $m;
            }else{
                throw $this->exception('An array should be associative');
            }
        }else{
            throw $this->exception('Data should be associative array or Model');
        }
        $this->generateXLS($data, $properties, $fields, $fields_width, $count_totals);
    }
    function generateXLS($data, $properties, $fields = null, $fields_width = null, $count_totals = null){
        $objPHPExcel = new \PHPExcel();
        $objPHPExcel
            ->getProperties()
            ->setCreator($properties['creator'])
            ->setLastModifiedBy($properties['lastModifiedBy'])
            ->setTitle($properties['title'])
            ->setSubject($properties['subject'])
            ->setDescription($properties['description'])
            ->setKeywords($properties['keywords'])
            ->setCategory($properties['category']);

        /*Set columns width*/
        if($fields_width){
            if(is_array($fields_width)){
                $count = 0;
                foreach ($fields_width as $width){
                    if(is_numeric($width)){
                        $objPHPExcel->getActiveSheet()->getColumnDimension($this->getColumnIndex($count))->setWidth($width);
                        $count++;
                    }else{
                        throw $this->exception('Values in $field_width should be numeric');
                    }
                }
            }else{
                throw $this->exception('$field_width should be an array');
            }
        }

        /*Get fields to show*/
        $exportFields = $this->getExportFields($data,$fields);

        /*Set first row with titles*/
        for($i=0; $i<count($exportFields); $i++){
            $objRichText = new \PHPExcel_RichText();
            $objPayable = $objRichText->createTextRun($exportFields[$i]);
            $objPayable->getFont()->setBold(true);
            $objPHPExcel->getActiveSheet()->setCellValue($this->getColumnIndex($i).'1', $objRichText);
        }
        $objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(20);

        /*Add content*/
        for($i=0; $i<count($data); $i++){
            $objPHPExcel->getActiveSheet()->getRowDimension($i+2)->setRowHeight(20);
            for($j=0; $j<count($exportFields); $j++){
                $objPHPExcel->setActiveSheetIndex(0)->setCellValue($this->getColumnIndex($j).($i+2),$data[$i][$exportFields[$j]]);
            }
        }
        /*Add totals*/
        if($count_totals){
            if(is_array($count_totals)){
                $objRichText = new \PHPExcel_RichText();
                $objPayable = $objRichText->createTextRun('TOTAL');
                $objPayable->getFont()->setBold(true);
                $objPHPExcel->getActiveSheet()->setCellValue($this->getColumnIndex(0).($i+2), $objRichText);

                /*Get fields to count*/
                $getFieldsList = $this->getExportFields($data,$count_totals);
                foreach($getFieldsList as $field){
                    $total = 0;
                    for($i=0; $i<count($data); $i++){
                        $total = $total+(float)$data[$i][$field];
                    }
                    $objRichText = new \PHPExcel_RichText();
                    $objPayable = $objRichText->createTextRun($total);
                    $objPayable->getFont()->setBold(true);
                    $objPHPExcel->getActiveSheet()->setCellValue($this->getColumnIndexByName($exportFields, $field).($i+2), $objRichText);

                    $objPHPExcel->getActiveSheet()->getRowDimension($i+2)->setRowHeight(20);
                }
            }else{
                throw $this->exception('$count_totals should be an array');
            }
        }

        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$properties['title'].'-'.date('Y-m-d-H-i-s').'.xls"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0

        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');

        exit;
    }

    /**
     * Checks if specified fields are in the source data or extracts fields titles from data
     * @param $arr
     * @param null $myFields
     * @return array
     * @throws \BaseException
     */
    function getExportFields($arr,$myFields = null){
        $fields = array();
        if($myFields && is_array($myFields)){
            foreach ($myFields as $myval){
                if(array_key_exists($myval,$arr[0])){
                    $fields[] = $myval;
                }else{
                    throw $this->exception('You have specified wrong field name');
                }
            }
        }else{
            foreach ($arr[0] as $key => $val){
                $fields[] = $key;
            }
        }
        return $fields;
    }

    /**
     * @param $i
     * @return mixed
     */
    function getColumnIndex($i){
        $columns=array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
        return $columns[$i];
    }

    /**
     * @param array $fields
     * @param string $field_name
     * @return int
     */
    function getColumnIndexByName($fields, $field_name){
        $index = 0;
        foreach($fields as $val){
            if($val == $field_name){
                return $this->getColumnIndex($index);
            }
            $index++;
        }
    }
}