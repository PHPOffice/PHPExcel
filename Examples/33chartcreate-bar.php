<?php

/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

date_default_timezone_set('Europe/London');

/**
 * PHPExcel
 *
 * Copyright (c) 2006 - 2016 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2016 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */

/** PHPExcel */
require_once dirname(__FILE__) . '/../src/Bootstrap.php';


$objPHPExcel = new \PhpOffice\PhpExcel\Spreadsheet();
$objWorksheet = $objPHPExcel->getActiveSheet();
$objWorksheet->fromArray(
	array(
		array('',	2010,	2011,	2012),
		array('Q1',   12,   15,		21),
		array('Q2',   56,   73,		86),
		array('Q3',   52,   61,		69),
		array('Q4',   30,   32,		0),
	)
);

//	Set the Labels for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesLabels = array(
	new \PhpOffice\PhpExcel\Chart\DataSeriesValues('String', 'Worksheet!$B$1', NULL, 1),	//	2010
	new \PhpOffice\PhpExcel\Chart\DataSeriesValues('String', 'Worksheet!$C$1', NULL, 1),	//	2011
	new \PhpOffice\PhpExcel\Chart\DataSeriesValues('String', 'Worksheet!$D$1', NULL, 1),	//	2012
);
//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues = array(
	new \PhpOffice\PhpExcel\Chart\DataSeriesValues('String', 'Worksheet!$A$2:$A$5', NULL, 4),	//	Q1 to Q4
);
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues = array(
	new \PhpOffice\PhpExcel\Chart\DataSeriesValues('Number', 'Worksheet!$B$2:$B$5', NULL, 4),
	new \PhpOffice\PhpExcel\Chart\DataSeriesValues('Number', 'Worksheet!$C$2:$C$5', NULL, 4),
	new \PhpOffice\PhpExcel\Chart\DataSeriesValues('Number', 'Worksheet!$D$2:$D$5', NULL, 4),
);

//	Build the dataseries
$series = new \PhpOffice\PhpExcel\Chart\DataSeries(
	\PhpOffice\PhpExcel\Chart\DataSeries::TYPE_BARCHART,		// plotType
	\PhpOffice\PhpExcel\Chart\DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues)-1),			// plotOrder
	$dataSeriesLabels,								// plotLabel
	$xAxisTickValues,								// plotCategory
	$dataSeriesValues								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series->setPlotDirection(\PhpOffice\PhpExcel\Chart\DataSeries::DIRECTION_BAR);

//	Set the series in the plot area
$plotArea = new \PhpOffice\PhpExcel\Chart\PlotArea(NULL, array($series));
//	Set the chart legend
$legend = new \PhpOffice\PhpExcel\Chart\Legend(\PhpOffice\PhpExcel\Chart\Legend::POSITION_RIGHT, NULL, false);

$title = new \PhpOffice\PhpExcel\Chart\Title('Test Bar Chart');
$yAxisLabel = new \PhpOffice\PhpExcel\Chart\Title('Value ($k)');


//	Create the chart
$chart = new \PhpOffice\PhpExcel\Chart(
	'chart1',		// name
	$title,			// title
	$legend,		// legend
	$plotArea,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel		// yAxisLabel
);

//	Set the position where the chart should appear in the worksheet
$chart->setTopLeftPosition('A7');
$chart->setBottomRightPosition('H20');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart);


// Save Excel 2007 file
echo date('H:i:s') , " Write to Excel2007 format" , EOL;
$objWriter = \PhpOffice\PhpExcel\IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->setIncludeCharts(TRUE);
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME)) , EOL;


// Echo memory peak usage
echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;

// Echo done
echo date('H:i:s') , " Done writing file" , EOL;
echo 'File has been created in ' , getcwd() , EOL;
