<?php
/**
 * PHPExcel
 *
 * Copyright (c) 2006 - 2013 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel\Shared
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */


namespace PHPExcel;

/**
 * PHPExcel\Shared_CodePage
 *
 * @category   PHPExcel
 * @package    PHPExcel\Shared
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class Shared_CodePage
{
    protected static $_codepages = array(
        0 => 'CP1252',                //    CodePage is not always correctly set when the xls file was saved by Apple's Numbers program, so default it
        367 => 'ASCII',               //    ASCII
        437 => 'CP437',               //    OEM US
        737 => 'CP737',               //    OEM Greek
        775 => 'CP775',               //    OEM Baltic
        850 => 'CP850',               //    OEM Latin I
        852 => 'CP852',               //    OEM Latin II (Central European)
        855 => 'CP855',               //    OEM Cyrillic
        857 => 'CP857',               //    OEM Turkish
        858 => 'CP858',               //    OEM Multilingual Latin I with Euro
        860 => 'CP860',               //    OEM Portugese
        861 => 'CP861',               //    OEM Icelandic
        862 => 'CP862',               //    OEM Hebrew
        863 => 'CP863',               //    OEM Canadian (French)
        864 => 'CP864',               //    OEM Arabic
        865 => 'CP865',               //    OEM Nordic
        866 => 'CP866',               //    OEM Cyrillic (Russian)
        869 => 'CP869',               //    OEM Greek (Modern)
        874 => 'CP874',               //    ANSI Thai
        932 => 'CP932',               //    ANSI Japanese Shift-JIS
        936 => 'CP936',               //    ANSI Chinese Simplified GBK
        949 => 'CP949',               //    ANSI Korean (Wansung)
        950 => 'CP950',               //    ANSI Chinese Traditional BIG5
        1200 => 'UTF-16LE',           //    UTF-16 (BIFF8)
        1250 => 'CP1250',             //    ANSI Latin II (Central European)
        1251 => 'CP1251',             //    ANSI Cyrillic
        1252 => 'CP1252',             //    ANSI Latin I (BIFF4-BIFF7)
        1253 => 'CP1253',             //    ANSI Greek
        1254 => 'CP1254',             //    ANSI Turkish
        1255 => 'CP1255',             //    ANSI Hebrew
        1256 => 'CP1256',             //    ANSI Arabic
        1257 => 'CP1257',             //    ANSI Baltic
        1258 => 'CP1258',             //    ANSI Vietnamese
        1361 => 'CP1361',             //    ANSI Korean (Johab)
        10000 => 'MAC',               //    Apple Roman
        10006 => 'MACGREEK',          //    Macintosh Greek
        10007 => 'MACCYRILLIC',       //    Macintosh Cyrillic
        10029 => 'MACCENTRALEUROPE',  //    Macintosh Central Europe
        10079 => 'MACICELAND',        //    Macintosh Icelandic
        10081 => 'MACTURKISH',        //    Macintosh Turkish
        32768 => 'MAC',               //    Apple Roman
        65000 => 'UTF-7',             //    Unicode (UTF-7)
        65001 => 'UTF-8',             //    Unicode (UTF-8)
    );

    /**
     * Convert Microsoft Code Page Identifier to Code Page Name which iconv
     * and mbstring understands
     *
     * @param integer $codePage Microsoft Code Page Indentifier
     * @return string Code Page Name
     * @throws PHPExcel\Exception
     */
    public static function NumberToName($codePage = 1252)
    {
        switch ($codePage) {
            case 720:   //    OEM Arabic
                throw new Exception('Code page 720 not supported.');
                break;
            case 32769: //    ANSI Latin I (BIFF2-BIFF3)
                throw new Exception('Code page 32769 not supported.');
                break;
            default:
                if (isset(self::$_codepages[$codePage])) {
                    return self::$_codepages[$codePage];
                }
        }

        throw new Exception('Unknown codepage: ' . $codePage);
    }
}
