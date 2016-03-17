<?php

namespace PhpOffice\PhpExcel\Writer\Excel2007;

/**
 * PhpOffice\PhpExcel\Writer\Excel2007\WriterPart
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
 * @package    PhpOffice\PhpExcel\Writer\Excel2007
 * @copyright  Copyright (c) 2006 - 2016 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
abstract class WriterPart
{
    /**
     * Parent IWriter object
     *
     * @var \PhpOffice\PhpExcel\Writer\IWriter
     */
    private $parentWriter;

    /**
     * Set parent IWriter object
     *
     * @param \PhpOffice\PhpExcel\Writer\IWriter    $pWriter
     * @throws \PhpOffice\PhpExcel\Writer\Exception
     */
    public function setParentWriter(\PhpOffice\PhpExcel\Writer\IWriter $pWriter = null)
    {
        $this->parentWriter = $pWriter;
    }

    /**
     * Get parent IWriter object
     *
     * @return \PhpOffice\PhpExcel\Writer\IWriter
     * @throws \PhpOffice\PhpExcel\Writer\Exception
     */
    public function getParentWriter()
    {
        if (!is_null($this->parentWriter)) {
            return $this->parentWriter;
        } else {
            throw new \PhpOffice\PhpExcel\Writer\Exception("No parent \\PhpOffice\\PhpExcel\\Writer\\IWriter assigned.");
        }
    }

    /**
     * Set parent IWriter object
     *
     * @param \PhpOffice\PhpExcel\Writer\IWriter    $pWriter
     * @throws \PhpOffice\PhpExcel\Writer\Exception
     */
    public function __construct(\PhpOffice\PhpExcel\Writer\IWriter $pWriter = null)
    {
        if (!is_null($pWriter)) {
            $this->parentWriter = $pWriter;
        }
    }
}
