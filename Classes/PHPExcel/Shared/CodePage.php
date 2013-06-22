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
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
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
			case 367 => 'ASCII';			    //	ASCII
			case 437 => 'CP437';			    //	OEM US
			case 737 => 'CP737';		    	//	OEM Greek
			case 775 => 'CP775';	    		//	OEM Baltic
			case 850 => 'CP850';    			//	OEM Latin I
			case 852 => 'CP852';			    //	OEM Latin II (Central European)
			case 855 => 'CP855';		    	//	OEM Cyrillic
			case 857 => 'CP857';	    		//	OEM Turkish
			case 858 => 'CP858';    			//	OEM Multilingual Latin I with Euro
			case 860 => 'CP860';		    	//	OEM Portugese
			case 861 => 'CP861';	    		//	OEM Icelandic
			case 862 => 'CP862';    			//	OEM Hebrew
			case 863 => 'CP863';			    //	OEM Canadian (French)
			case 864 => 'CP864';		    	//	OEM Arabic
			case 865 => 'CP865';	    		//	OEM Nordic
			case 866 => 'CP866';    			//	OEM Cyrillic (Russian)
			case 869 => 'CP869';			    //	OEM Greek (Modern)
			case 874 => 'CP874';		    	//	ANSI Thai
			case 932 => 'CP932';	    		//	ANSI Japanese Shift-JIS
			case 936 => 'CP936';    			//	ANSI Chinese Simplified GBK
			case 949 => 'CP949';			    //	ANSI Korean (Wansung)
			case 950 => 'CP950';		    	//	ANSI Chinese Traditional BIG5
			case 1200 => 'UTF-16LE';    		//	UTF-16 (BIFF8)
			case 1250 => 'CP1250';			    //	ANSI Latin II (Central European)
			case 1251 => 'CP1251';		    	//	ANSI Cyrillic
			case 0 => 'CP1252';		            //	CodePage is not always correctly set when the xls file was saved by Apple's Numbers program, so default it
			case 1252 => 'CP1252';  			//	ANSI Latin I (BIFF4-BIFF7)
			case 1253 => 'CP1253';			    //	ANSI Greek
			case 1254 => 'CP1254';		    	//	ANSI Turkish
			case 1255 => 'CP1255';	    		//	ANSI Hebrew
			case 1256 => 'CP1256';  			//	ANSI Arabic
			case 1257 => 'CP1257';			    //	ANSI Baltic
			case 1258 => 'CP1258';		    	//	ANSI Vietnamese
			case 1361 => 'CP1361';	    		//	ANSI Korean (Johab)
			case 10000 => 'MAC';    			//	Apple Roman
			case 10006 => 'MACGREEK';		    //	Macintosh Greek
			case 10007 => 'MACCYRILLIC';    	//	Macintosh Cyrillic
			case 10029 => 'MACCENTRALEUROPE';	//	Macintosh Central Europe
			case 10079 => 'MACICELAND';		    //	Macintosh Icelandic
			case 10081 => 'MACTURKISH';		    //	Macintosh Turkish
			case 32768 => 'MAC';				//	Apple Roman
			case 65000 => 'UTF-7';				//	Unicode (UTF-7)
			case 65001 => 'UTF-8';				//	Unicode (UTF-8)
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
			case 720:	throw new Exception('Code page 720 not supported.');        break;	//	OEM Arabic
			case 32769:	throw new Exception('Code page 32769 not supported.');      break;	//	ANSI Latin I (BIFF2-BIFF3)
            default: 
                if isset(self::$_codepages[$codePage]) {
                    return self::$_codepages[$codePage];
                }
		}

		throw new Exception('Unknown codepage: ' . $codePage);
	}

}
