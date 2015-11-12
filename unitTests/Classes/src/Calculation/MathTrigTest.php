<?php


require_once __DIR__.'/../../../testDataFileIterator.php';

class MathTrigTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        require_once(__DIR__.'/../../../../src/Autoloader.php');

        \PHPExcel\Calculation\Functions::setCompatibilityMode(\PHPExcel\Calculation\Functions::COMPATIBILITY_EXCEL);
    }

    /**
     * @dataProvider providerATAN2
     */
    public function testATAN2()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','ATAN2'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerATAN2()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/ATAN2.data');
    }

    /**
     * @dataProvider providerCEILING
     */
    public function testCEILING()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','CEILING'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerCEILING()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/CEILING.data');
    }

    /**
     * @dataProvider providerCOMBIN
     */
    public function testCOMBIN()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','COMBIN'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerCOMBIN()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/COMBIN.data');
    }

    /**
     * @dataProvider providerEVEN
     */
    public function testEVEN()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','EVEN'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerEVEN()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/EVEN.data');
    }

    /**
     * @dataProvider providerODD
     */
    public function testODD()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','ODD'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerODD()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/ODD.data');
    }

    /**
     * @dataProvider providerFACT
     */
    public function testFACT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','FACT'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerFACT()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/FACT.data');
    }

    /**
     * @dataProvider providerFACTDOUBLE
     */
    public function testFACTDOUBLE()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','FACTDOUBLE'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerFACTDOUBLE()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/FACTDOUBLE.data');
    }

    /**
     * @dataProvider providerFLOOR
     */
    public function testFLOOR()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','FLOOR'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerFLOOR()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/FLOOR.data');
    }

    /**
     * @dataProvider providerGCD
     */
    public function testGCD()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','GCD'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerGCD()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/GCD.data');
    }

    /**
     * @dataProvider providerLCM
     */
    public function testLCM()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','LCM'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerLCM()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/LCM.data');
    }

    /**
     * @dataProvider providerINT
     */
    public function testINT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','INT'), $args);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerINT()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/INT.data');
    }

    /**
     * @dataProvider providerSIGN
     */
    public function testSIGN()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','SIGN'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerSIGN()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/SIGN.data');
    }

    /**
     * @dataProvider providerPOWER
     */
    public function testPOWER()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','POWER'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerPOWER()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/POWER.data');
    }

    /**
     * @dataProvider providerLOG
     */
    public function testLOG()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','LOG_BASE'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerLOG()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/LOG.data');
    }

    /**
     * @dataProvider providerMOD
     */
    public function testMOD()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','MOD'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerMOD()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/MOD.data');
    }

    /**
     * @dataProvider providerMDETERM
     */
    public function testMDETERM()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','MDETERM'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerMDETERM()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/MDETERM.data');
    }

    /**
     * @dataProvider providerMINVERSE
     */
    public function testMINVERSE()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','MINVERSE'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerMINVERSE()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/MINVERSE.data');
    }

    /**
     * @dataProvider providerMMULT
     */
    public function testMMULT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','MMULT'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerMMULT()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/MMULT.data');
    }

    /**
     * @dataProvider providerMULTINOMIAL
     */
    public function testMULTINOMIAL()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','MULTINOMIAL'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerMULTINOMIAL()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/MULTINOMIAL.data');
    }

    /**
     * @dataProvider providerMROUND
     */
    public function testMROUND()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        \PHPExcel\Calculation::setArrayReturnType(\PHPExcel\Calculation::RETURN_ARRAY_AS_VALUE);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','MROUND'), $args);
        \PHPExcel\Calculation::setArrayReturnType(\PHPExcel\Calculation::RETURN_ARRAY_AS_ARRAY);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerMROUND()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/MROUND.data');
    }

    /**
     * @dataProvider providerPRODUCT
     */
    public function testPRODUCT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','PRODUCT'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerPRODUCT()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/PRODUCT.data');
    }

    /**
     * @dataProvider providerQUOTIENT
     */
    public function testQUOTIENT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','QUOTIENT'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerQUOTIENT()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/QUOTIENT.data');
    }

    /**
     * @dataProvider providerROUNDUP
     */
    public function testROUNDUP()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','ROUNDUP'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerROUNDUP()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/ROUNDUP.data');
    }

    /**
     * @dataProvider providerROUNDDOWN
     */
    public function testROUNDDOWN()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','ROUNDDOWN'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerROUNDDOWN()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/ROUNDDOWN.data');
    }

    /**
     * @dataProvider providerSERIESSUM
     */
    public function testSERIESSUM()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','SERIESSUM'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerSERIESSUM()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/SERIESSUM.data');
    }

    /**
     * @dataProvider providerSUMSQ
     */
    public function testSUMSQ()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','SUMSQ'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerSUMSQ()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/SUMSQ.data');
    }

    /**
     * @dataProvider providerTRUNC
     */
    public function testTRUNC()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','TRUNC'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerTRUNC()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/TRUNC.data');
    }

    /**
     * @dataProvider providerROMAN
     */
    public function testROMAN()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','ROMAN'), $args);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerROMAN()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/ROMAN.data');
    }

    /**
     * @dataProvider providerSQRTPI
     */
    public function testSQRTPI()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig','SQRTPI'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerSQRTPI()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/MathTrig/SQRTPI.data');
    }

    /**
     * @dataProvider providerSUMIF
     */
    public function testSUMIF()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\MathTrig', 'SUMIF'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-12);
    }

    public function providerSUMIF()
    {
        return array(
            array(
                array(
                    array(1),
                    array(5),
                    array(10),
                ),
                '>=5',
                15,
            ),
            array(
                array(
                    array('text'),
                    array(2),
                ),
                '=text',
                array(
                    array(10),
                    array(100),
                ),
                10,
            ),
            array(
                array(
                    array('"text with quotes"'),
                    array(2),
                ),
                '="text with quotes"',
                array(
                    array(10),
                    array(100),
                ),
                10,
            ),
            array(
                array(
                    array('"text with quotes"'),
                    array(''),
                ),
                '>"', // Compare to the single characater " (double quote)
                array(
                    array(10),
                    array(100),
                ),
                10
            ),
            array(
                array(
                    array(''),
                    array('anything'),
                ),
                '>"', // Compare to the single characater " (double quote)
                array(
                    array(10),
                    array(100),
                ),
                100
            ),
        );
    }
}
