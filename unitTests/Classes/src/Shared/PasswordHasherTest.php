<?php


require_once __DIR__.'/../../../testDataFileIterator.php';

class PasswordHasherTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        require_once(__DIR__.'/../../../../src/Autoloader.php');
    }

    /**
     * @dataProvider providerHashPassword
     */
    public function testHashPassword()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\PHPExcel\Shared\PasswordHasher','hashPassword'), $args);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerHashPassword()
    {
        return new testDataFileIterator(__DIR__.'/../../../rawTestData/Shared/PasswordHashes.data');
    }
}
