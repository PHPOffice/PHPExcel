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
 * @package    PHPExcel\Worksheet
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */


namespace PHPExcel;

/**
 * PHPExcel\Worksheet_CellIterator
 *
 * Used to iterate rows in a PHPExcel\Worksheet
 *
 * @category   PHPExcel
 * @package    PHPExcel\Worksheet
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class Worksheet_CellIterator implements \Iterator
{
    /**
     * PHPExcel\Worksheet to iterate
     *
     * @var PHPExcel\Worksheet
     */
    protected $_subject;

    /**
     * Row index
     *
     * @var int
     */
    protected $_rowIndex;

    /**
     * Current iterator position
     *
     * @var int
     */
    protected $_position = 0;

    /**
     * Loop only existing cells
     *
     * @var boolean
     */
    protected $_onlyExistingCells = true;

    /**
     * Create a new cell iterator
     *
     * @param PHPExcel\Worksheet         $subject
     * @param int                        $rowIndex
     */
    public function __construct(Worksheet $subject = null, $rowIndex = 1) {
        // Set subject and row index
        $this->_subject     = $subject;
        $this->_rowIndex     = $rowIndex;
    }

    /**
     * Destructor
     */
    public function __destruct() {
        unset($this->_subject);
    }

    /**
     * Rewind iterator
     */
    public function rewind() {
        $this->_position = 0;
    }

    /**
     * Current PHPExcel\Cell
     *
     * @return PHPExcel\Cell
     */
    public function current() {
        return $this->_subject->getCellByColumnAndRow($this->_position, $this->_rowIndex);
    }

    /**
     * Current key
     *
     * @return int
     */
    public function key() {
        return $this->_position;
    }

    /**
     * Next value
     */
    public function next() {
        ++$this->_position;
    }

    /**
     * Are there any more PHPExcel\Cell instances available?
     *
     * @return boolean
     */
    public function valid() {
        // columnIndexFromString() returns an index based at one,
        // treat it as a count when comparing it to the base zero
        // position.
        $columnCount = Cell::columnIndexFromString($this->_subject->getHighestColumn());

        if ($this->_onlyExistingCells) {
            // If we aren't looking at an existing cell, either
            // because the first column doesn't exist or next() has
            // been called onto a nonexistent cell, then loop until we
            // find one, or pass the last column.
            while ($this->_position < $columnCount &&
                   !$this->_subject->cellExistsByColumnAndRow($this->_position, $this->_rowIndex)) {
                ++$this->_position;
            }
        }

        return $this->_position < $columnCount;
    }

    /**
     * Get loop only existing cells
     *
     * @return boolean
     */
    public function getIterateOnlyExistingCells() {
        return $this->_onlyExistingCells;
    }

    /**
     * Set the iterator to loop only existing cells
     *
     * @param    boolean        $value
     */
    public function setIterateOnlyExistingCells($value = true) {
        $this->_onlyExistingCells = $value;
    }
}
