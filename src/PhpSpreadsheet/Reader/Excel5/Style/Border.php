<?php

namespace PhpOffice\PhpExcel\Reader\Excel5\Style;

use \PhpOffice\PhpExcel\Style\Border As BaseBorder;

class Border
{
    protected static $map = array(
        0x00 => BaseBorder::BORDER_NONE,
        0x01 => BaseBorder::BORDER_THIN,
        0x02 => BaseBorder::BORDER_MEDIUM,
        0x03 => BaseBorder::BORDER_DASHED,
        0x04 => BaseBorder::BORDER_DOTTED,
        0x05 => BaseBorder::BORDER_THICK,
        0x06 => BaseBorder::BORDER_DOUBLE,
        0x07 => BaseBorder::BORDER_HAIR,
        0x08 => BaseBorder::BORDER_MEDIUMDASHED,
        0x09 => BaseBorder::BORDER_DASHDOT,
        0x0A => BaseBorder::BORDER_MEDIUMDASHDOT,
        0x0B => BaseBorder::BORDER_DASHDOTDOT,
        0x0C => BaseBorder::BORDER_MEDIUMDASHDOTDOT,
        0x0D => BaseBorder::BORDER_SLANTDASHDOT,
    );

    /**
     * Map border style
     * OpenOffice documentation: 2.5.11
     *
     * @param int $index
     * @return string
     */
    public static function lookup($index)
    {
        if (isset(self::$map[$index])) {
            return self::$map[$index];
        }
        return BaseBorder::BORDER_NONE;
    }
}