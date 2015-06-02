<?php

class WorksheetColumnTest extends PHPUnit_Framework_TestCase
{
    public $mockWorksheet;
    public $mockColumn;

    public function setUp()
    {
        require_once(__DIR__.'/../../../../src/Autoloader.php');
        
        $this->mockRow = $this->getMockBuilder('PHPExcel\Worksheet\Row')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet = $this->getMockBuilder('PHPExcel_Worksheet')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet->expects($this->any())
                 ->method('getHighestRow')
                 ->will($this->returnValue(5));
        $this->mockWorksheet->expects($this->any())
                 ->method('current')
                 ->will($this->returnValue($this->mockRow));
    }


    public function testIteratorFullRange()
    {
        $iterator = new PHPExcel\Worksheet\RowIterator($this->mockWorksheet);
        $rowIndexResult = 1;
        $this->assertEquals($rowIndexResult, $iterator->key());
        
        foreach ($iterator as $key => $row) {
            $this->assertEquals($rowIndexResult++, $key);
            $this->assertInstanceOf('PHPExcel\Worksheet\Row', $row);
        }
    }

    public function testIteratorStartEndRange()
    {
        $iterator = new PHPExcel\Worksheet\RowIterator($this->mockWorksheet, 2, 4);
        $rowIndexResult = 2;
        $this->assertEquals($rowIndexResult, $iterator->key());
        
        foreach ($iterator as $key => $row) {
            $this->assertEquals($rowIndexResult++, $key);
            $this->assertInstanceOf('PHPExcel\Worksheet\Row', $row);
        }
    }

    public function testIteratorSeekAndPrev()
    {
        $iterator = new PHPExcel\Worksheet\RowIterator($this->mockWorksheet, 2, 4);
        $columnIndexResult = 4;
        $iterator->seek(4);
        $this->assertEquals($columnIndexResult, $iterator->key());

        for ($i = 1; $i < $columnIndexResult-1; $i++) {
            $iterator->prev();
            $this->assertEquals($columnIndexResult - $i, $iterator->key());
        }
    }

    /**
     * @expectedException PHPExcel_Exception
     */
    public function testSeekOutOfRange()
    {
        $iterator = new PHPExcel\Worksheet\RowIterator($this->mockWorksheet, 2, 4);
        $iterator->seek(1);
    }

    /**
     * @expectedException PHPExcel_Exception
     */
    public function testPrevOutOfRange()
    {
        
        
        $this->mockWorksheet = $this->getMockBuilder('PHPExcel_Worksheet')
            ->disableOriginalConstructor()
            ->getMock();
        $this->mockWorksheet->expects($this->any())
                 ->method('getHighestRow')
                 ->will($this->returnValue(5));
    }


    public function testInstantiateColumnDefault()
    {
        $column = new PHPExcel\Worksheet\Column($this->mockWorksheet);
        $this->assertInstanceOf('PHPExcel\Worksheet\Column', $column);
        $columnIndex = $column->getColumnIndex();
        $this->assertEquals('A', $columnIndex);
    }

    public function testInstantiateColumnSpecified()
    {
        $column = new PHPExcel\Worksheet\Column($this->mockWorksheet, 'E');
        $this->assertInstanceOf('PHPExcel\Worksheet\Column', $column);
        $columnIndex = $column->getColumnIndex();
        $this->assertEquals('E', $columnIndex);
    }

    public function testGetCellIterator()
    {
        $column = new PHPExcel\Worksheet\Column($this->mockWorksheet);
        $cellIterator = $column->getCellIterator();
        $this->assertInstanceOf('PHPExcel\Worksheet\ColumnCellIterator', $cellIterator);
    }
}
