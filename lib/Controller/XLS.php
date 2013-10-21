<?php
/**
 *Created by Konstantin Kolodnitsky
 * Date: 02.10.13
 * Time: 10:16
 */
namespace kk_xls;
require_once __DIR__.'/../../vendor/phpexcel/PHPExcel.php';
class Controller_XLS extends \AbstractController{
    function init(){
        parent::init();

        // add add-on locations to pathfinder
        $l = $this->api->locate('addons',__NAMESPACE__,'location');
        $addon_location = $this->api->locate('addons',__NAMESPACE__);
        $this->api->pathfinder->addLocation($addon_location,array(
            'php'=>array('lib','vendor')
        ))->setParent($l);
    }

    /**
     * @param $properties
     * @param $m
     * @param null $fields
     * @throws \BaseException
     */
    function generateXLS($properties, $m, $fields = null){
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

        /*Getting fields*/
        $exportFields = $this->getExportFields($data,$fields);

        $objPHPExcel = new \PHPExcel();
        $objPHPExcel->getProperties()->setCreator($properties['creator'])
            ->setLastModifiedBy($properties['lastModifiedBy'])
            ->setTitle($properties['title'])
            ->setSubject($properties['subject'])
            ->setDescription($properties['description'])
            ->setKeywords($properties['keywords'])
            ->setCategory($properties['category']);
        $objPHPExcel->getActiveSheet()->getColumnDimension($this->getColumnIndex(0))->setWidth(15);
        $objPHPExcel->getActiveSheet()->getColumnDimension($this->getColumnIndex(1))->setWidth(30);
        $objPHPExcel->getActiveSheet()->getColumnDimension($this->getColumnIndex(2))->setWidth(10);
        $objPHPExcel->getActiveSheet()->getColumnDimension($this->getColumnIndex(3))->setWidth(14);
        $objPHPExcel->getActiveSheet()->getColumnDimension($this->getColumnIndex(4))->setWidth(12);
        $objPHPExcel->getActiveSheet()->getColumnDimension($this->getColumnIndex(5))->setWidth(15);
        $objPHPExcel->getActiveSheet()->getColumnDimension($this->getColumnIndex(6))->setWidth(15);

        for($i=0; $i<count($exportFields); $i++){
            $objRichText = new \PHPExcel_RichText();
            $objPayable = $objRichText->createTextRun($exportFields[$i]);
            $objPayable->getFont()->setBold(true);
            $objPHPExcel->getActiveSheet()->setCellValue($this->getColumnIndex($i).'1', $objRichText);
        }
        $objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(20);

        $total_spent=0;
        for($i=0; $i<count($data); $i++){
            $objPHPExcel->getActiveSheet()->getRowDimension($i+2)->setRowHeight(20);
            for($j=0; $j<count($exportFields); $j++){
                $objPHPExcel->setActiveSheetIndex(0)->setCellValue($this->getColumnIndex($j).($i+2),$data[$i][$exportFields[$j]]);
            }
            $total_spent=$total_spent+(float)$data[$i]['spent_time'];
        }

        $objRichText = new \PHPExcel_RichText();
        $objPayable = $objRichText->createTextRun('TOTAL');
        $objPayable->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->setCellValue($this->getColumnIndex(0).($i+2), $objRichText);

        $objRichText = new \PHPExcel_RichText();
        $objPayable = $objRichText->createTextRun($total_spent);
        $objPayable->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->setCellValue($this->getColumnIndex(4).($i+2), $objRichText);

        $objPHPExcel->getActiveSheet()->getRowDimension($i+2)->setRowHeight(20);

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
}