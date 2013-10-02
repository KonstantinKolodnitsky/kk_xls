<?php
/**
 *Created by Konstantin Kolodnitsky
 * Date: 02.10.13
 * Time: 10:16
 */
namespace kk_xls;
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
    function generateXLS($properties, $m){
        require_once '../vendor/phpexcel/PHPExcel.php';
        exit('qwe');
        $objPHPExcel = new PHPExcel();
        $objPHPExcel->getProperties()->setCreator($properties['creator'])
            ->setLastModifiedBy($properties['creator'])
            ->setTitle($properties['title'])
            ->setSubject($properties['subject'])
            ->setDescription($properties['description'])
            ->setKeywords($properties['keywords'])
            ->setCategory($properties['category']);
//        $objPHPExcel->getActiveSheet()->getColumnDimension($this->getColumnIndex(0))->setWidth(15);
//        $objPHPExcel->getActiveSheet()->getColumnDimension($this->getColumnIndex(1))->setWidth(30);
//        $objPHPExcel->getActiveSheet()->getColumnDimension($this->getColumnIndex(2))->setWidth(10);
//        $objPHPExcel->getActiveSheet()->getColumnDimension($this->getColumnIndex(3))->setWidth(14);
//        $objPHPExcel->getActiveSheet()->getColumnDimension($this->getColumnIndex(4))->setWidth(12);
//        $objPHPExcel->getActiveSheet()->getColumnDimension($this->getColumnIndex(5))->setWidth(12);
//        $objPHPExcel->getActiveSheet()->getColumnDimension($this->getColumnIndex(6))->setWidth(15);

        for($i=0; $i<count($this->export_fields); $i++){
            $objRichText = new PHPExcel_RichText();
            $objPayable = $objRichText->createTextRun($this->export_fields[$i]);
            $objPayable->getFont()->setBold(true);
            $objPHPExcel->getActiveSheet()->setCellValue($this->getColumnIndex($i).'1', $objRichText);
        }
//        $objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(20);

        if(is_object($m)){
            $data=$m->getRows();
        }else{
            $data = $m;
        }
        $total_spent=0;
        for($i=0; $i<count($data); $i++){
            $objPHPExcel->getActiveSheet()->getRowDimension($i+2)->setRowHeight(20);
            for($j=0; $j<count($this->export_fields); $j++){
                $objPHPExcel->setActiveSheetIndex(0)->setCellValue($this->getColumnIndex($j).($i+2),$data[$i][$this->export_fields[$j]]);
            }
            $total_spent=$total_spent+(float)$data[$i]['spent_time'];
        }

        $objRichText = new PHPExcel_RichText();
        $objPayable = $objRichText->createTextRun('TOTAL');
        $objPayable->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->setCellValue($this->getColumnIndex(0).($i+2), $objRichText);

        $objRichText = new PHPExcel_RichText();
        $objPayable = $objRichText->createTextRun($total_spent);
        $objPayable->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->setCellValue($this->getColumnIndex(5).($i+2), $objRichText);

        $objPHPExcel->getActiveSheet()->getRowDimension($i+2)->setRowHeight(20);

        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="report-colubris-'.date('Y-m-i-H-i-s').'.xls"');
        header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0

        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');

        exit;
    }
}