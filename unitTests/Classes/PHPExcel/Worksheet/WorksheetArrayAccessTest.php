<?php

class ArrayAccessTest extends PHPUnit_Framework_TestCase
{

	public function setUp()
	{
		if (!defined('PHPEXCEL_ROOT'))
		{
			define('PHPEXCEL_ROOT', APPLICATION_PATH . '/');
		}
		require_once(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
	}

    public function testArrayAccess(){
        $workbook = new PHPExcel();
        $worksheet = $workbook->getActiveSheet();
        $worksheet->setCellValue('A1', 1);
        $worksheet->setCellValue('A2', 2);

        // reading
        $this->assertEquals($worksheet->getCell('A1')->getValue(), $worksheet['A1']);
        $this->assertEquals(1, $worksheet['A1']);

        // setting up values
        $worksheet['A2'] = $worksheet['A1'];
        $this->assertEquals(1, $worksheet['A2']);
        // exist ?
        $this->assertFalse(isset($worksheet['A5']));
        // unset
        unset($worksheet['A2']);
        $this->assertEmpty($worksheet['A2']);

        $workbook->addSheet(new PHPExcel_Worksheet($workbook, 'TestSheet'));
        $sheet = $workbook->getSheetByName('TestSheet');
        $this->assertNotNull($sheet);
        $this->assertNotNull($workbook['TestSheet']);
        $this->assertEquals($sheet, $workbook['TestSheet']);

        $this->assertTrue(isset($workbook['TestSheet']));
        $this->assertInstanceOf('PHPExcel_Worksheet', $workbook['TestSheet']);

        $newSheet = new PHPExcel_Worksheet($workbook, 'TestSheet2');
        $workbook['TestSheet'] = $newSheet;
        $this->assertFalse(isset($workbook['TestSheet']));
        $this->assertTrue(isset($workbook['TestSheet2']));

        $workbook['TestSheet2']['A1'] = 1;
        $this->assertTrue(isset($workbook['TestSheet2']['A1']));
        $this->assertEquals(1, $workbook['TestSheet2']['A1']);

    }

}
