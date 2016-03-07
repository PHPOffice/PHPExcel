<?php

namespace PhpOffice\PhpExcel\Worksheet;

/**
 * PHPExcel_Worksheet_ColumnIterator
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
 * @package    PHPExcel_Worksheet
 * @copyright  Copyright (c) 2006 - 2016 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class ColumnIterator implements \Iterator
{
    /**
     * \PhpOffice\PhpExcel\Worksheet to iterate
     *
     * @var \PhpOffice\PhpExcel\Worksheet
     */
    private $subject;

    /**
     * Current iterator position
     *
     * @var int
     */
    private $position = 0;

    /**
     * Start position
     *
     * @var int
     */
    private $startColumn = 0;


    /**
     * End position
     *
     * @var int
     */
    private $endColumn = 0;


    /**
     * Create a new column iterator
     *
     * @param    \PhpOffice\PhpExcel\Worksheet   $subject       The worksheet to iterate over
     * @param    string                $startColumn   The column address at which to start iterating
     * @param    string                $endColumn     Optionally, the column address at which to stop iterating
     */
    public function __construct(\PhpOffice\PhpExcel\Worksheet $subject, $startColumn = 'A', $endColumn = null)
    {
        // Set subject
        $this->subject = $subject;
        $this->resetEnd($endColumn);
        $this->resetStart($startColumn);
    }

    /**
     * Destructor
     */
    public function __destruct()
    {
        unset($this->subject);
    }

    /**
     * (Re)Set the start column and the current column pointer
     *
     * @param integer    $startColumn    The column address at which to start iterating
     * @return ColumnIterator
     * @throws PHPExcel_Exception
     */
    public function resetStart($startColumn = 'A')
    {
        $startColumnIndex = \PhpOffice\PhpExcel\Cell::columnIndexFromString($startColumn) - 1;
        if ($startColumnIndex > PHPExcel_Cell::columnIndexFromString($this->subject->getHighestColumn()) - 1) {
            throw new PHPExcel_Exception("Start column ({$startColumn}) is beyond highest column ({$this->subject->getHighestColumn()})");
        }

        $this->startColumn = $startColumnIndex;
        if ($this->endColumn < $this->startColumn) {
            $this->endColumn = $this->startColumn;
        }
        $this->seek($startColumn);

        return $this;
    }

    /**
     * (Re)Set the end column
     *
     * @param string    $endColumn    The column address at which to stop iterating
     * @return ColumnIterator
     */
    public function resetEnd($endColumn = null)
    {
        $endColumn = ($endColumn) ? $endColumn : $this->subject->getHighestColumn();
        $this->endColumn = \PhpOffice\PhpExcel\Cell::columnIndexFromString($endColumn) - 1;

        return $this;
    }

    /**
     * Set the column pointer to the selected column
     *
     * @param   string    $column    The column address to set the current pointer at
     * @return  ColumnIterator
     * @throws  \PhpOffice\PhpExcel\Exception
     */
    public function seek($column = 'A')
    {
        $column = \PhpOffice\PhpExcel\Cell::columnIndexFromString($column) - 1;
        if (($column < $this->startColumn) || ($column > $this->endColumn)) {
            throw new \PhpOffice\PhpExcel\Exception("Column $column is out of range ({$this->startColumn} - {$this->endColumn})");
        }
        $this->position = $column;

        return $this;
    }

    /**
     * Rewind the iterator to the starting column
     */
    public function rewind()
    {
        $this->position = $this->startColumn;
    }

    /**
     * Return the current column in this worksheet
     *
     * @return Column
     */
    public function current()
    {
        return new Column($this->subject, \PhpOffice\PhpExcel\Cell::stringFromColumnIndex($this->position));
    }

    /**
     * Return the current iterator key
     *
     * @return string
     */
    public function key()
    {
        return \PhpOffice\PhpExcel\Cell::stringFromColumnIndex($this->position);
    }

    /**
     * Set the iterator to its next value
     */
    public function next()
    {
        ++$this->position;
    }

    /**
     * Set the iterator to its previous value
     *
     * @throws \PhpOffice\PhpExcel\Exception
     */
    public function prev()
    {
        if ($this->position <= $this->startColumn) {
            throw new \PhpOffice\PhpExcel\Exception(
                "Column is already at the beginning of range (" .
                \PhpOffice\PhpExcel\Cell::stringFromColumnIndex($this->endColumn) . " - " .
                \PhpOffice\PhpExcel\Cell::stringFromColumnIndex($this->endColumn) . ")"
            );
        }

        --$this->position;
    }

    /**
     * Indicate if more columns exist in the worksheet range of columns that we're iterating
     *
     * @return boolean
     */
    public function valid()
    {
        return $this->position <= $this->endColumn;
    }
}
