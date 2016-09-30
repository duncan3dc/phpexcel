<?php

namespace duncan3dc\PhpexcelTests;

use duncan3dc\Phpexcel\Excel;

class ExcelTest extends \PHPUnit_Framework_TestCase
{

    public function testReadAll()
    {
        $result = Excel::read(__DIR__ . "/spreadsheets/multiple_sheets.xlsx");

        $this->assertSame([
            [
                ["Header 1", "Header 2", "Header 3"],
                ["Row 1A", "Row 1B", "Row 1C"],
                ["Row 2A", "Row 2B", "Row 2C"],
            ],
            [
                ["Sheet 2", "Data"],
            ],
        ], $result);
    }


    public function testReadSingleSheet1()
    {
        $result = Excel::read(__DIR__ . "/spreadsheets/multiple_sheets.xlsx", 0);

        $this->assertSame([
            ["Header 1", "Header 2", "Header 3"],
            ["Row 1A", "Row 1B", "Row 1C"],
            ["Row 2A", "Row 2B", "Row 2C"],
        ], $result);
    }
    public function testReadSingleSheet2()
    {
        $result = Excel::read(__DIR__ . "/spreadsheets/multiple_sheets.xlsx", 1);

        $this->assertSame([
            ["Sheet 2", "Data"],
        ], $result);
    }


    public function testReadByName()
    {
        $result = Excel::read(__DIR__ . "/spreadsheets/multiple_sheets.xlsx", "Sheet2");

        $this->assertSame([
            ["Sheet 2", "Data"],
        ], $result);
    }


    public function assertWriteAndRead(array $expected, callable $callback)
    {
        $tmp = tempnam("/tmp", "phpexcel_");

        $excel = new Excel;

        $callback($excel);

        $excel->save($tmp);

        $result = Excel::read($tmp, 0);

        $this->assertSame($expected, $result);

        unlink($tmp);
    }


    public function testSetCell1()
    {
        $this->assertWriteAndRead([
            ["Test"],
        ], function ($excel) {
            $excel->setCell("A1", "Test");
        });
    }
    public function testSetCell2()
    {
        $this->assertWriteAndRead([
            ["7E4"],
        ], function ($excel) {
            $excel->setCell("A1", "7E4", Excel::STRING);
        });
    }
    public function testSetCell3()
    {
        $this->assertWriteAndRead([
            [800.0],
        ], function ($excel) {
            $excel->setCell("A1", "800A", Excel::NUMERIC);
        });
    }
}
