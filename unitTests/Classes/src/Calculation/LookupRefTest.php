<?php


require_once __DIR__.'/../../../testDataFileIterator.php';

class LookupRefTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        if (!defined('PHPEXCEL_ROOT')) {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(__DIR__.'/../../../../src/Autoloader.php');

        \PHPExcel\Calculation\Functions::setCompatibilityMode(\PHPExcel\Calculation\Functions::COMPATIBILITY_EXCEL);
    }

    /**
     * @dataProvider providerHLOOKUP
     */
    public function testHLOOKUP()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\LookupRef','HLOOKUP'), $args);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerHLOOKUP()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/LookupRef/HLOOKUP.data');
    }

    /**
     * @dataProvider providerVLOOKUP
     */
    public function testVLOOKUP()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\LookupRef','VLOOKUP'), $args);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerVLOOKUP()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/LookupRef/VLOOKUP.data');
    }
}
