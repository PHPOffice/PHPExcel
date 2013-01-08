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

$filename = str_replace('.php', '.xlsx', __FILE__);

$peSettings = array(
	'CachedObjectStorage' => array(
		'type' => 'Memory',
		'args' => array(),
		),
	'Writer' => array(
		'type' => 'Excel2007'
		),
	);

$workbook = new PHPExcel_Adapter_Spreadsheet_Excel_Writer($filename, $peSettings);

$ncols = 50;
$nrows = 100;

$formatArray = array();
for ($i = 0; $i < 10; $i++) {
	$format = $workbook->addFormat();
	$format->setSize($i + 4);
	$formatArray[] = $format;
}

$worksheet = $workbook->addWorksheet();

echo date('H:i:s') , " Add data (begin)" , EOL;
$t = microtime(true);
for ($col = 0; $col < $ncols; $col++) {
	for ($row = 0; $row < $nrows; $row++) {
		$str = ($row + $col);
		$format = $formatArray[$row % count($formatArray)];
		$worksheet->write($row, $col, $str, $format);
	}
}
$d = microtime(true) - $t;
echo date('H:i:s') , " Add data (end), time: " . round($d, 2) . " s", EOL;

echo date('H:i:s') , " Write file" , EOL;
$t = microtime(true);
$workbook->close();
$d = microtime(true) - $t;
echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME)) ,
	", time: " . round($d, 2) . " s" , EOL;

echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;

echo date('H:i:s') , " Done writing file" , EOL;
echo 'File has been created in ' , getcwd() , EOL;
?>
