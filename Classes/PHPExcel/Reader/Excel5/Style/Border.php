<?php

class PHPExcel_Reader_Excel5_Style_Border
{
    public static $map = array(
        0x00 => PHPExcel_Style_Border::BORDER_NONE,
        0x01 => PHPExcel_Style_Border::BORDER_THIN,
        0x02 => PHPExcel_Style_Border::BORDER_MEDIUM,
        0x03 => PHPExcel_Style_Border::BORDER_DASHED,
        0x04 => PHPExcel_Style_Border::BORDER_DOTTED,
        0x05 => PHPExcel_Style_Border::BORDER_THICK,
        0x06 => PHPExcel_Style_Border::BORDER_DOUBLE,
        0x07 => PHPExcel_Style_Border::BORDER_HAIR,
        0x08 => PHPExcel_Style_Border::BORDER_MEDIUMDASHED,
        0x09 => PHPExcel_Style_Border::BORDER_DASHDOT,
        0x0A => PHPExcel_Style_Border::BORDER_MEDIUMDASHDOT,
        0x0B => PHPExcel_Style_Border::BORDER_DASHDOTDOT,
        0x0C => PHPExcel_Style_Border::BORDER_MEDIUMDASHDOTDOT,
        0x0D => PHPExcel_Style_Border::BORDER_SLANTDASHDOT,
    );
}