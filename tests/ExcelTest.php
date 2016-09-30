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
}
