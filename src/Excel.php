<?php

namespace duncan3dc\Phpexcel;

class Excel extends \PHPExcel
{
    const BOLD    =   1;
    const ITALIC  =   2;
    const LEFT    =   4;
    const RIGHT   =   8;
    const CENTER  =  16;
    const STRING  =  32;
    const NUMERIC =  64;



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

        $sheet = $this->GetActiveSheet();

        if ($style & static::STRING) {
            $sheet->SetCellValueExplicit($cell, $value, \PHPExcel_Cell_DataType::TYPE_STRING);
        } elseif ($style & static::NUMERIC) {
            $sheet->SetCellValueExplicit($cell, $value, \PHPExcel_Cell_DataType::TYPE_NUMERIC);
        } else {
            $sheet->SetCellValue($cell, $value);
        }

        $sheet->GetColumnDimension($col)->setAutoSize(true);

        $cellStyle = $sheet->GetStyle($cell);

        $font = $cellStyle->getFont();
        if ($style & static::BOLD) {
            $font->setBold(true);
        }
        if ($style & static::ITALIC) {
            $font->setItalic(true);
        }

        $alignment = $cellStyle->getAlignment();
        if ($style & static::LEFT) {
            $alignment->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        }
        if ($style & static::RIGHT) {
            $alignment->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        }
        if ($style & static::CENTER) {
            $alignment->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        }

        return $this;
    }


    public function addImage($cell, $path)
    {
        $image = new \PHPExcel_Worksheet_Drawing;

        $name = pathinfo($path, PATHINFO_BASENAME);
        $image->setName($name);

        $image->setPath($path);

        $row = preg_replace("/^[a-z]+([0-9]+)$/i", "$1", $cell);
        $image->setHeight(50);
        $this->getActiveSheet()->getRowDimension($row)->setRowHeight(50);

        $image->setCoordinates($cell);

        $image->setWorksheet($this->getActiveSheet());
    }


    /**
     * Writes a row of data to the excel document based on a multidimensional array with styling.
     *
     * Example - This will write the first column with Test and align it right and format it as a numeric.
     * [
     *   "field1"   =>  [
     *      "value"     =>  "Test",
     *      "align"     =>  true,
     *      "numeric"   =>  true,
     *   ],
     * ]
     *
     * @param array $data The data row to write with styling.
     * @param int   $row  The row number to write to.
     *
     * @return void
     */
    public function writeRow(array $data, $row = 1)
    {
        $sheet = $this->getActiveSheet();

        $values = [];

        $styles = [
            "bold"      =>  [],
            "align"     =>  [],
            "numeric"   =>  [],
        ];

        $column = 0;
        foreach ($data as $field => $detail) {
            if (array_key_exists("bold", $detail) && $detail["bold"]) {
                $styles["bold"][$field] = $column;
            }

            if (array_key_exists("numeric", $detail) && $detail["numeric"]) {
                $styles["numeric"][$field] = $column;
            }

            if (array_key_exists("align", $detail) && $detail["align"]) {
                $styles["align"][$field] = $column;
            }

            $values[$field] = $detail["value"];

            $column++;
        }

        $rowStart = $this->getCellName(0, $row);

        $rowEnd = $this->getCellName($column, $row);

        $sheet->getStyle("{$rowStart}:{$rowEnd}")->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

        foreach ($styles as $type => $columns) {
            if (empty($columns)) {
                continue;
            }

            $consecutive = $this->getConsecutiveFields($columns);

            foreach ($consecutive as $fields) {
                $startName = $this->getCellName(current($fields), $row);

                $endName = $this->getCellName(end($fields), $row);

                $cellName = "{$startName}:{$endName}";
                if ($startName === $endName) {
                    $cellName = $startName;
                }

                $cellStyles = $sheet->getStyle($cellName);
                if ($type === "bold") {
                    $cellStyles->getFont()->setBold(true);
                }

                if ($type === "numeric") {
                    $cellStyles->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
                }

                if ($type === "align") {
                    $cellStyles->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                }
            }
        }

        $sheet->fromArray($values, null, $this->getCellName(0, $row), true);
    }


    /**
     * Get an array containing consecutive fields based upon their values.
     *
     * @param array $data An array of fields with their column numbers as values.
     *
     * @return array
     */
    private function getConsecutiveFields(array $data)
    {
        $consecutive = [];

        $lastColumn = null;

        $count = 0;
        foreach ($data as $key => $number) {
            if (is_null($lastColumn)) {
                $count++;

                $lastColumn = $number;

                $consecutive[$count][$key] = $number;

                continue;
            }

            if ($number == ($lastColumn + 1)) {
                $consecutive[$count][$key] = $number;

                $lastColumn = $number;

                continue;
            }

            $count++;

            $consecutive[$count][$key] = $number;

            $lastColumn = null;
        }

        return $consecutive;
    }
}
