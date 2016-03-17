<?php

namespace PhpOffice\PhpExcel\Writer\OpenDocument;

/**
 * PhpOffice\PhpExcel\Writer\OpenDocument\Mimetype
 *
 * Copyright (c) 2006 - 2016 PHPExcel
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
 * @package    PhpOffice\PhpExcel\Writer\OpenDocument
 * @copyright  Copyright (c) 2006 - 2016 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Mimetype extends WriterPart
{
    /**
     * Write mimetype to plain text format
     *
     * @param \PhpOffice\PhpExcel\SpreadSheet $pPHPExcel
     * @return     string         XML Output
     * @throws     \PhpOffice\PhpExcel\Writer\Exception
     */
    public function write(\PhpOffice\PhpExcel\SpreadSheet $pPHPExcel = null)
    {
        return 'application/vnd.oasis.opendocument.spreadsheet';
    }
}
