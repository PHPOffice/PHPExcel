<?php


require_once __DIR__.'/../../../testDataFileIterator.php';

class FinancialTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        require_once(__DIR__.'/../../../../src/Autoloader.php');

        \PHPExcel\Calculation\Functions::setCompatibilityMode(\PHPExcel\Calculation\Functions::COMPATIBILITY_EXCEL);
    }

    /**
     * @dataProvider providerACCRINT
     */
    public function testACCRINT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','ACCRINT'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerACCRINT()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/ACCRINT.data');
    }

    /**
     * @dataProvider providerACCRINTM
     */
    public function testACCRINTM()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','ACCRINTM'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerACCRINTM()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/ACCRINTM.data');
    }

    /**
     * @dataProvider providerAMORDEGRC
     */
    public function testAMORDEGRC()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','AMORDEGRC'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerAMORDEGRC()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/AMORDEGRC.data');
    }

    /**
     * @dataProvider providerAMORLINC
     */
    public function testAMORLINC()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','AMORLINC'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerAMORLINC()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/AMORLINC.data');
    }

    /**
     * @dataProvider providerCOUPDAYBS
     */
    public function testCOUPDAYBS()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','COUPDAYBS'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerCOUPDAYBS()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/COUPDAYBS.data');
    }

    /**
     * @dataProvider providerCOUPDAYS
     */
    public function testCOUPDAYS()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','COUPDAYS'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerCOUPDAYS()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/COUPDAYS.data');
    }

    /**
     * @dataProvider providerCOUPDAYSNC
     */
    public function testCOUPDAYSNC()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','COUPDAYSNC'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerCOUPDAYSNC()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/COUPDAYSNC.data');
    }

    /**
     * @dataProvider providerCOUPNCD
     */
    public function testCOUPNCD()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','COUPNCD'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerCOUPNCD()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/COUPNCD.data');
    }

    /**
     * @dataProvider providerCOUPNUM
     */
    public function testCOUPNUM()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','COUPNUM'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerCOUPNUM()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/COUPNUM.data');
    }

    /**
     * @dataProvider providerCOUPPCD
     */
    public function testCOUPPCD()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','COUPPCD'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerCOUPPCD()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/COUPPCD.data');
    }

    /**
     * @dataProvider providerCUMIPMT
     */
    public function testCUMIPMT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','CUMIPMT'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerCUMIPMT()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/CUMIPMT.data');
    }

    /**
     * @dataProvider providerCUMPRINC
     */
    public function testCUMPRINC()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','CUMPRINC'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerCUMPRINC()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/CUMPRINC.data');
    }

    /**
     * @dataProvider providerDB
     */
    public function testDB()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','DB'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerDB()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/DB.data');
    }

    /**
     * @dataProvider providerDDB
     */
    public function testDDB()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','DDB'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerDDB()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/DDB.data');
    }

    /**
     * @dataProvider providerDISC
     */
    public function testDISC()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','DISC'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerDISC()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/DISC.data');
    }

    /**
     * @dataProvider providerDOLLARDE
     */
    public function testDOLLARDE()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','DOLLARDE'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerDOLLARDE()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/DOLLARDE.data');
    }

    /**
     * @dataProvider providerDOLLARFR
     */
    public function testDOLLARFR()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','DOLLARFR'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerDOLLARFR()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/DOLLARFR.data');
    }

    /**
     * @dataProvider providerEFFECT
     */
    public function testEFFECT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','EFFECT'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerEFFECT()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/EFFECT.data');
    }

    /**
     * @dataProvider providerFV
     */
    public function testFV()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','FV'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerFV()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/FV.data');
    }

    /**
     * @dataProvider providerFVSCHEDULE
     */
    public function testFVSCHEDULE()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','FVSCHEDULE'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerFVSCHEDULE()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/FVSCHEDULE.data');
    }

    /**
     * @dataProvider providerINTRATE
     */
    public function testINTRATE()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','INTRATE'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerINTRATE()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/INTRATE.data');
    }

    /**
     * @dataProvider providerIPMT
     */
    public function testIPMT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','IPMT'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerIPMT()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/IPMT.data');
    }

    /**
     * @dataProvider providerIRR
     */
    public function testIRR()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','IRR'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerIRR()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/IRR.data');
    }

    /**
     * @dataProvider providerISPMT
     */
    public function testISPMT()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','ISPMT'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerISPMT()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/ISPMT.data');
    }

    /**
     * @dataProvider providerMIRR
     */
    public function testMIRR()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','MIRR'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerMIRR()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/MIRR.data');
    }

    /**
     * @dataProvider providerNOMINAL
     */
    public function testNOMINAL()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','NOMINAL'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerNOMINAL()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/NOMINAL.data');
    }

    /**
     * @dataProvider providerNPER
     */
    public function testNPER()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','NPER'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerNPER()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/NPER.data');
    }

    /**
     * @dataProvider providerNPV
     */
    public function testNPV()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','NPV'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerNPV()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/NPV.data');
    }

    /**
     * @dataProvider providerPRICE
     */
    public function testPRICE()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','PRICE'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerPRICE()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/PRICE.data');
    }

    /**
     * @dataProvider providerRATE
     */
    public function testRATE()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','RATE'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerRATE()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/RATE.data');
    }

    /**
     * @dataProvider providerXIRR
     */
    public function testXIRR()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Calculation\Financial','XIRR'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerXIRR()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Calculation/Financial/XIRR.data');
    }
}
