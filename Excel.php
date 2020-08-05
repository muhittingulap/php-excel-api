<?php
/**
 * Created by PhpStorm.
 * User: muhittin
 * Date: 28.03.2020
 * Time: 15:17
 */

if (!class_exists("PHPExcel")) require_once(ROOT_FOLDER . 'inc/cls/excel/PHPExcel.php');

class Excel
{
    public $filename = "excel_";
    public $baslangic = "A1";
    public $worksheetname = "DATA";
    public $dir = ""; /* excelin kaydedileği tam yol */
    public $type = 1; /* 1: Sadece Kaydet 2: Direk İndir */
    public $keys = 1; /* 1: Sadece Kaydet 2: Direk İndir */

    public function __construct()
    {
        $this->filename .= date('d_m_Y--H_i_s');
        $this->excel = new PHPExcel();
    }

    public function setFileName($data)
    {
        $this->filename = $data;
        return $this;
    }

    public function setKeys($data)
    {
        $this->keys = $data;
        return $this;
    }

    public function setWorkSheetName($data)
    {
        $this->worksheetname = $data;
        return $this;
    }

    public function setBaslangic($data)
    {
        $this->baslangic = $data;
        return $this;
    }

    public function setDir($data)
    {
        $this->dir = $data;
        return $this;
    }

    public function setType($data)
    {
        $this->type = $data;
        return $this;
    }

    public function setData($data)
    {
        $this->data = $data;
        return $this;
    }

    public function export()
    {
        if ($this->keys) array_unshift($this->data, array_keys($this->data[0])); /* Keysleri Data arayinin başına ekliyoruz */

        $this->excel->getActiveSheet()->fromArray($this->data, null, $this->baslangic); // datayı başlangıç column undan itibaren basar
        $this->excel->getActiveSheet()->setTitle($this->worksheetname); // work sheet başlık

        // @var PHPExcel_Worksheet $sheet *  // bu foreach autosize yapar
        foreach ($this->excel->getAllSheets() as $sheet) {
            for ($col = 0; $col <= PHPExcel_Cell::columnIndexFromString($sheet->getHighestDataColumn()); $col++) {
                $sheet->getColumnDimensionByColumn($col)->setAutoSize(true);
            }
        }


        $name = $this->dir . $this->filename . ".xlsx";

        $objWriter = PHPExcel_IOFactory::createWriter($this->excel, 'Excel2007');
        $objWriter->save($name);

        return array("file" => $name, "name" => $this->filename . ".xlsx");
    }

    public function import()
    {
        $excelReader = PHPExcel_IOFactory::createReaderForFile($this->filename);
        $excelObj = $excelReader->load($this->filename);
        $worksheet = $excelObj->getActiveSheet();
        $data=$worksheet->toArray(null,false,true,false);



        return $data;

    }


}