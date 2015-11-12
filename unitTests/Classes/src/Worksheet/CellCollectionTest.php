<?php

namespace PHPExcel\Worksheet;

class CellCollectionTest extends \PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        require_once(__DIR__.'/../../../../src/Autoloader.php');
    }

    public function testCacheLastCell()
    {
        $methods = \PHPExcel\CachedObjectStorageFactory::getCacheStorageMethods();
        foreach ($methods as $method) {
            \PHPExcel\CachedObjectStorageFactory::initialize($method);
            $workbook = new \PHPExcel\Spreadsheet(); 
            $cells = ['A1', 'A2'];
            $worksheet = $workbook->getActiveSheet();
            $worksheet->setCellValue('A1', 1);
            $worksheet->setCellValue('A2', 2);
            $this->assertEquals($cells, $worksheet->getCellCollection(), "Cache method \"$method\".");
            \PHPExcel\CachedObjectStorageFactory::finalize();
        }
    }
}
