<?php

namespace Classes\PHPExcel\Reader;

use PHPExcel_Reader_OOCalc;

class OOCalcTest extends \PHPUnit_Framework_TestCase {

    /**
     * @group oocalc
     */
    public function testSpacesParsing(){

        $reader = new PHPExcel_Reader_OOCalc();
        $file = $reader->load(__DIR__."/../../../rawTestData/Reader/OOCalc/spaces-everywhere.ods");

        $arr = $file->getActiveSheet()->toArray();

        $this->assertEquals([
            ["This has    4 spaces before and 2 after  "],
            ["This only one after "],
            ["Test with DIFFERENT styles     and multiple spaces:  "],
            ["test with new \nLines"],
        ], $arr);
    }
}
