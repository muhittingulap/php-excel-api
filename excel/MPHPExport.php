<?php
require_once(ROOT_FOLDER .'integrasyon/export/PHPExcel.php');

class MPHPExport
{
    public function ExcelExport($D = array(), $filename ="system/data/autorapor/1.xlsx", $Export = 1, $WorkSheetTitle = "DATA")
    {
         $objPHPExcel = new PHPExcel();
         $objPHPExcel->getActiveSheet()->fromArray($D, null, 'A1'); // datayı başlangıç column undan itibaren basar
         $objPHPExcel->getActiveSheet()->setTitle($WorkSheetTitle); // work sheet başlık

         /** @var PHPExcel_Worksheet $sheet */  // bu foreach autosize yapar
        foreach ($objPHPExcel->getAllSheets() as $sheet) {
            for ($col = 0; $col <= PHPExcel_Cell::columnIndexFromString($sheet->getHighestDataColumn()); $col++) {
                $sheet->getColumnDimensionByColumn($col)->setAutoSize(true);
            }
        }

        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save($filename);
    }
}

?>