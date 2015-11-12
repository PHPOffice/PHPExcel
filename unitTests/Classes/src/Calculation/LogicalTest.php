<?php


require_once __DIR__.'/../../../testDataFileIterator.php';

class LogicalTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        require_once(__DIR__.'/../../../../src/Autoloader.php');

        \PHPExcel\Calculation\Functions::setCompatibilityMode(\PHPExcel\Calculation\Functions::COMPATIBILITY_EXCEL);
    }

    public function testTRUE()
    {
        $result = \PHPExcel\Calculation\Logical::TRUE();
        $this->assertEquals(true, $result);
    }

    public function testFALSE()
    {
        $result = \PHPExcel\Calculation\Logical::FALSE();
        $this->assertEquals(false, $result);
    }

    /**
     * @dataProvider providerAND
     */
    public function testAND()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Logical','LOGICAL_AND'), $args);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerAND()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Logical/AND.data');
    }

    /**
     * @dataProvider providerOR
     */
    public function testOR()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Logical','LOGICAL_OR'), $args);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerOR()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Logical/OR.data');
    }

    /**
     * @dataProvider providerNOT
     */
    public function testNOT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Logical','NOT'), $args);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerNOT()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Logical/NOT.data');
    }

    /**
     * @dataProvider providerIF
     */
    public function testIF()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Logical','STATEMENT_IF'), $args);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerIF()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Logical/IF.data');
    }

    /**
     * @dataProvider providerIFERROR
     */
    public function testIFERROR()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Logical','IFERROR'), $args);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerIFERROR()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Logical/IFERROR.data');
    }
}
