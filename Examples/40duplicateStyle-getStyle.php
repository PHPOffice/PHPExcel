<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

date_default_timezone_set('Europe/London');

/** Include PHPExcel */
require_once '../Classes/PHPExcel.php';

$ncols = 50;
$nrows = 100;

echo date('H:i:s') , " Create new PHPExcel object" , EOL;
$objPHPExcel = new PHPExcel();
$worksheet = $objPHPExcel->getActiveSheet();

echo date('H:i:s') , " Create styles array" , EOL;
$styles = array();
for ($i = 0; $i < 10; $i++) {
	$style = new PHPExcel_Style();
	$style->getFont()->setSize($i + 4);
	$styles[] = $style;
}

echo date('H:i:s') , " Add data (begin)" , EOL;
$t = microtime(true);
echo date('H:i:s') , " _cellXfCollection count before: " , count($objPHPExcel->getCellXfCollection()), EOL;
for ($col = 0; $col < $ncols; $col++) {
	for ($row = 0; $row < $nrows; $row++) {
		$str = ($row + $col);
		$coord = PHPExcel_Cell::stringFromColumnIndex($col) . ($row + 1);
		$worksheet->setCellValue($coord, $str);
		if ($col === 0) {
			$style = $styles[$row % count($styles)];
			$worksheet->duplicateStyle($style, $coord);
		}
	}
}
echo date('H:i:s') , " _cellXfCollection count after first column styles: " , count($objPHPExcel->getCellXfCollection()), EOL;
for ($row = 0; $row < $nrows; $row++) {
	$style = $worksheet->getStyleByColumnAndRow(0, $row + 1);
	$cellRange =
		PHPExcel_Cell::stringFromColumnIndex(1) . ($row + 1) .
		':' .
		PHPExcel_Cell::stringFromColumnIndex($ncols - 1) . ($row + 1);
	$worksheet->duplicateStyle($style, $cellRange);
}
echo date('H:i:s') , " _cellXfCollection count after rows styles: " , count($objPHPExcel->getCellXfCollection()), EOL;
$d = microtime(true) - $t;
echo date('H:i:s') , " Add data (end), time: " . round($d, 2) . " s", EOL;


echo date('H:i:s') , " Write to Excel2007 format" , EOL;
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME)) , EOL;


echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;

echo date('H:i:s') , " Done writing file" , EOL;
echo 'File has been created in ' , getcwd() , EOL;
