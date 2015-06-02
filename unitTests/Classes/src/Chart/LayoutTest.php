<?php


class LayoutTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        if (!defined('PHPEXCEL_ROOT')) {
            define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
        }
        require_once(__DIR__.'/../../../../src/Autoloader.php');
    }

    public function testSetLayoutTarget()
    {
        $LayoutTargetValue = 'String';

        $testInstance = new \PHPExcel\Chart\Layout;

        $result = $testInstance->setLayoutTarget($LayoutTargetValue);
        $this->assertTrue($result instanceof \PHPExcel\Chart\Layout);
    }

    public function testGetLayoutTarget()
    {
        $LayoutTargetValue = 'String';

        $testInstance = new \PHPExcel\Chart\Layout;
        $setValue = $testInstance->setLayoutTarget($LayoutTargetValue);

        $result = $testInstance->getLayoutTarget();
        $this->assertEquals($LayoutTargetValue, $result);
    }
}
