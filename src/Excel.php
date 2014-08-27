<?php

namespace duncan3dc\Phpexcel;

class Excel extends \PHPExcel
{
    const BOLD    =   1;
    const ITALIC  =   2;
    const LEFT    =   4;
    const RIGHT   =   8;
    const CENTER  =  16;


    public static function read($filename, $key = -1)
    {
        $type = \PHPExcel_IOFactory::identify($filename);
        $reader = \PHPExcel_IOFactory::createReader($type);
        $reader->setReadDataOnly(true);
        $excel = $reader->load($filename);

        if (is_string($key)) {
            $sheet = $excel->getSheetByName($key);
        } elseif ($key > -1) {
            $sheet = $excel->getSheet($key);
        } else {
            $sheets = [];
            foreach ($excel->getWorksheetIterator() as $sheet) {
                $sheets[] = $sheet->toArray(null, false, false, false);
            }
            return $sheets;
        }

        return $sheet->toArray(null, false, false, false);
    }


    public function save($filename)
    {
        return (new \PHPExcel_Writer_Excel2007($this))->save($filename);
    }


    public function output($filename)
    {
        $tmp = tempnam("/tmp", "excel_");

        $this->save($tmp);

        $data = file_get_contents($tmp);

        if (substr($filename, -5) != ".xlsx") {
            $filename .= ".xlsx";
        }

        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        header("Content-Length: " . strlen($data));

        echo $data;

        unlink($tmp);
    }


    public static function getCellName($col, $row)
    {
        $cell = "";

        $col += 65;

        $prefix = 64;
        while ($col > 90) {
            $prefix++;
            $col -= 26;
        }

        if ($prefix > 64) {
            $cell .= chr($prefix);
        }

        $cell .= chr($col);

        $cell .= $row;

        return $cell;
    }


    public function setCell($cell, $value, $style = null)
    {
        preg_match("/^([a-z]+)([0-9]+)$/i", $cell, $matches);
        $col = $matches[1];
        $row = $matches[2];

        $this->getActiveSheet()->SetCellValue($cell, $value);

        $this->getActiveSheet()->getColumnDimension($col)->setAutoSize(true);

        if (!$style) {
            return true;
        }

        if ($style & static::BOLD) {
            $this->getActiveSheet()->getStyle($cell)->getFont()->setBold(true);
        }

        if ($style & static::ITALIC) {
            $this->getActiveSheet()->getStyle($cell)->getFont()->setItalic(true);
        }

        if ($style & static::LEFT) {
            $this->getActiveSheet()->getStyle($cell)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        }

        if ($style & static::RIGHT) {
            $this->getActiveSheet()->getStyle($cell)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        }

        if ($style & static::CENTER) {
            $this->getActiveSheet()->getStyle($cell)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        }
    }


    public function addImage($cell, $path)
    {
        $image = new \PHPExcel_Worksheet_Drawing();

        $name = pathinfo($path, PATHINFO_BASENAME);
        $image->setName($name);

        $image->setPath($path);

        $row = preg_replace("/^[a-z]+([0-9]+)$/i", "$1", $cell);
        $image->setHeight(50);
        $this->getActiveSheet()->getRowDimension($row)->setRowHeight(50);

        $image->setCoordinates($cell);

        $image->setWorksheet($this->getActiveSheet());
    }
}
