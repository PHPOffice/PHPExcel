# PHPExcel - OpenXML - Read, Write and Create spreadsheet documents in PHP - Spreadsheet engine
PHPExcel is a library written in pure PHP and providing a set of classes that allow you to write to and read from different spreadsheet file formats, like Excel (BIFF) .xls, Excel 2007 (OfficeOpenXML) .xlsx, CSV, Libre/OpenOffice Calc .ods, Gnumeric, PDF, HTML, ... This project is built around Microsoft's OpenXML standard and PHP.

Master: [![Build Status](https://travis-ci.org/PHPOffice/PHPExcel.png?branch=master)](http://travis-ci.org/PHPOffice/PHPExcel)

Develop: [![Build Status](https://travis-ci.org/PHPOffice/PHPExcel.png?branch=develop)](http://travis-ci.org/PHPOffice/PHPExcel)

[![Join the chat at https://gitter.im/PHPOffice/PHPExcel](https://img.shields.io/badge/GITTER-join%20chat-green.svg)](https://gitter.im/PHPOffice/PHPExcel)

## File Formats supported

### Reading
 * BIFF 5-8 (.xls) Excel 95 and above
 * Office Open XML (.xlsx) Excel 2007 and above
 * SpreadsheetML (.xml) Excel 2003
 * Open Document Format/OASIS (.ods)
 * Gnumeric
 * HTML
 * SYLK
 * CSV

### Writing
 * BIFF 8 (.xls) Excel 95 and above
 * Office Open XML (.xlsx) Excel 2007 and above
 * HTML
 * CSV
 * PDF (using either the tcPDF, DomPDF or mPDF libraries, which need to be installed separately)


## Requirements
 * PHP version 5.2.0 or higher
 * PHP extension php_zip enabled (required if you need PHPExcel to handle .xlsx .ods or .gnumeric files)
 * PHP extension php_xml enabled
 * PHP extension php_gd2 enabled (optional, but required for exact column width autocalculation)

*Note:* PHP 5.6.29 has [a bug](https://bugs.php.net/bug.php?id=73530) that
prevents SQLite3 caching to work correctly. Use a newer (or older) versions of
PHP if you need SQLite3 caching.

## Want to contribute?

PHPExcel developement for next version has moved under its new name PhpSpreadsheet. So please head over to [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet#want-to-contribute) to contribute patches and features.


## License
PHPExcel is licensed under [LGPL (GNU LESSER GENERAL PUBLIC LICENSE)](https://github.com/PHPOffice/PHPExcel/blob/master/license.md)
