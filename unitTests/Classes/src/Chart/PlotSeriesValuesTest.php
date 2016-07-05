<?php

namespace PHPExcel\Chart;

use PHPExcel;
use PHPExcel\Chart;
use PHPExcel\Exception;
use PHPExcel\IOFactory;

class PlotSeriesValuesTest extends \PHPUnit_Framework_TestCase
{
    private $dataTypeValuesSets = [
        [
            'dataSeries' => DataSeriesValues::DATASERIES_TYPE_STRING,
            'xAxis' => DataSeriesValues::DATASERIES_TYPE_NUMBER
        ],
        [
            'dataSeries' => DataSeriesValues::DATASERIES_TYPE_NUMBER,
            'xAxis' => DataSeriesValues::DATASERIES_TYPE_STRING
        ],
        [
            'dataSeries' => DataSeriesValues::DATASERIES_TYPE_NUMBER,
            'xAxis' => DataSeriesValues::DATASERIES_TYPE_NUMBER
        ],
        [
            'dataSeries' => DataSeriesValues::DATASERIES_TYPE_STRING,
            'xAxis' => DataSeriesValues::DATASERIES_TYPE_STRING
        ],
    ];

    public function testPlotSeriesValuesDataType()
    {
        $objWriter = IOFactory::createWriter(new PHPExcel\Spreadsheet, 'Excel2007');
        foreach ($this->dataTypeValuesSets as $dataSeriesTypes) {
            $chart = $this->generateChart($dataSeriesTypes['dataSeries'], $dataSeriesTypes['xAxis']);
            $xmlString = $objWriter->getWriterPart('Chart')->writeChart($chart);
            $this->plotSeriesValuesDataTypeXmlCheck($xmlString, $dataSeriesTypes['dataSeries']);
        }
    }

    /**
     * Generate simple fake line Chart
     * @param string $dataSeriesType
     * @param string $xAxisTickType
     * @return Chart
     * @internal param string $dataSeriesType
     */
    protected function generateChart(
        $dataSeriesType = DataSeriesValues::DATASERIES_TYPE_NUMBER,
        $xAxisTickType = DataSeriesValues::DATASERIES_TYPE_NUMBER

    ) {
        try {
            $layout = new Layout();
            $axisLabels = new Title(null, $layout);
            $Axis = new Axis();
            $majorGridlines = new GridLines();
            $dataSeriesLabels = array(new DataSeriesValues,);

            $dataSeriesValues = array(
                new DataSeriesValues($dataSeriesType)
            );

            $xAxisTickValues = array(
                new DataSeriesValues($xAxisTickType),
            );

            $series = new DataSeries(
                DataSeries::TYPE_LINECHART,        // plotType
                DataSeries::GROUPING_STANDARD,    // plotGrouping
                range(0, mt_rand(0, 200)),            // plotOrder
                $dataSeriesLabels,                                // plotLabel
                $xAxisTickValues,                                // plotCategory
                $dataSeriesValues//,								// plotValues
            );
            $plotArea = new PlotArea($layout, array($series));
            $chart = new Chart(
                'chart',        // name
                new Title(null, $layout),            // title
                new Legend(Legend::POSITION_RIGHT, $layout),        // legend
                $plotArea,
                true,
                0,
                $axisLabels,
                $axisLabels,
                $Axis,
                $Axis,
                $majorGridlines
            );
            $chart->setWorksheet(new \PHPExcel\Worksheet());
            return $chart;
        } catch (Exception $e) {
            $this->fail($e->getMessage());
        }
        return null;
    }

    /**
     * @param $xmlString
     * @param $type
     */
    private function plotSeriesValuesDataTypeXmlCheck($xmlString, $type)
    {
        switch ($type) {
            case DataSeriesValues::DATASERIES_TYPE_NUMBER:
                $dataType = 'num';
                break;
            case DataSeriesValues::DATASERIES_TYPE_STRING:
            default:
                $dataType = 'str';
                break;
        }

        $dom = new \DOMDocument();
        $dom->loadXML($xmlString);
        $xpath = new \DOMXpath($dom);
        $path = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:val/c:' . $dataType . 'Ref';
        $nodeList = $xpath->query($path);
        $this->assertTrue((bool)$nodeList->length, 'Path: '.$path.' failed. DataSeries Type: '.$type);
    }
}
