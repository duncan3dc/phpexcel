phpexcel
========

A simple wrapper around the PHPExcel library

[![Build Status](https://travis-ci.org/duncan3dc/phpexcel.svg?branch=master)](https://travis-ci.org/duncan3dc/phpexcel)
[![Latest Stable Version](https://poser.pugx.org/duncan3dc/phpexcel/version.svg)](https://packagist.org/packages/duncan3dc/phpexcel)


Static Methods
--------------
* read(string $filename[, mixed $key]): array - Reads a spreadsheet and converts it's contents to an array (using the PHPExcel toString() method).
By default this will return an enumerated array with each element representing one sheet.
A specific sheet can be requested by passing the $key argument, either as an integer representing the (zero based) sheet number, or as a string of the sheet name.
* getCellName(int $col, int $row): string - Convert a numeric column number (zero based) and row number into a cell name (eg B3)


Public Methods
--------------
* save(string $filename): null - Calls the save() method on the PHPExcel_Writer_Excel2007 class.
* output(string $filename): null - Outputs the spreadsheet to the browser to prompt a download.
* addImage(string $cell, string $path): null - Add an image to the specified cell.
* setCell(string $cell, mixed $value[, int $style]): null - Set a cell to a value.
The $style parameter can be used to set several styles on the cell, using the following class constants:
BOLD
ITALIC
LEFT
RIGHT
CENTER


Examples
--------

The Excel class uses a namespace of duncan3dc
```php
use duncan3dc\Excel;
```

-------------------

```php
$excel = new Excel();

$excel->setCell("A1","Title", Excel::BOLD | EXCEL::CENTER);

for($i = 1; $i < 10; $i++) {
    $cell = $excel->getCellName($i,1);
    $excel->setCell($cell,"Test " . $i);
}

for($i = 1; $i < 10; $i++) {
    $cell = Excel::getCellName($i,2);
    $excel->setCell($cell,"Value " . $i);
}

$excel->save("/tmp/text.xls");
```
