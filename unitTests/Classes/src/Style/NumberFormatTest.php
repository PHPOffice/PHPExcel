<?php


require_once __DIR__.'/../../../testDataFileIterator.php';

class NumberFormatTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        require_once(__DIR__.'/../../../../src/Autoloader.php');

        \PHPExcel\Shared\String::setDecimalSeparator('.');
        \PHPExcel\Shared\String::setThousandsSeparator(',');
    }

    /**
     * @dataProvider providerNumberFormat
     */
    public function testFormatValueWithMask()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Style\NumberFormat','toFormattedString'), $args);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerNumberFormat()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Style/NumberFormat.data');
    }
}
