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
 * PHPExcel\Worksheet_RowDimension
 *
 * @category   PHPExcel
 * @package    PHPExcel\Worksheet
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class Worksheet_RowDimension
{
    /**
     * Row index
     *
     * @var int
     */
    protected $_rowIndex;

    /**
     * Row height (in pt)
     *
     * When this is set to a negative value, the row height should be ignored by IWriter
     *
     * @var double
     */
    protected $_rowHeight        = -1;

     /**
     * ZeroHeight for Row?
     *
     * @var bool
     */
    protected $_zeroHeight    = false;

    /**
     * Visible?
     *
     * @var bool
     */
    protected $_visible        = true;

    /**
     * Outline level
     *
     * @var int
     */
    protected $_outlineLevel    = 0;

    /**
     * Collapsed
     *
     * @var bool
     */
    protected $_collapsed        = false;

    /**
     * Index to cellXf. Null value means row has no explicit cellXf format.
     *
     * @var int|null
     */
    protected $_xfIndex;

    /**
     * Create a new PHPExcel\Worksheet_RowDimension
     *
     * @param int $pIndex Numeric row index
     */
    public function __construct($pIndex = 0)
    {
        // Initialise values
        $this->_rowIndex        = $pIndex;

        // set row dimension as unformatted by default
        $this->_xfIndex = null;
    }

    /**
     * Get Row Index
     *
     * @return int
     */
    public function getRowIndex() {
        return $this->_rowIndex;
    }

    /**
     * Set Row Index
     *
     * @param int $pValue
     * @return PHPExcel\Worksheet_RowDimension
     */
    public function setRowIndex($pValue) {
        $this->_rowIndex = $pValue;
        return $this;
    }

    /**
     * Get Row Height
     *
     * @return double
     */
    public function getRowHeight() {
        return $this->_rowHeight;
    }

    /**
     * Set Row Height
     *
     * @param double $pValue
     * @return PHPExcel\Worksheet_RowDimension
     */
    public function setRowHeight($pValue = -1) {
        $this->_rowHeight = $pValue;
        return $this;
    }

    /**
     * Get ZeroHeight
     *
     * @return bool
     */
    public function getzeroHeight() {
        return $this->_zeroHeight;
    }

    /**
     * Set ZeroHeight
     *
     * @param bool $pValue
     * @return PHPExcel\Worksheet_RowDimension
     */
    public function setzeroHeight($pValue = false) {
        $this->_zeroHeight = $pValue;
        return $this;
    }

    /**
     * Get Visible
     *
     * @return bool
     */
    public function getVisible() {
        return $this->_visible;
    }

    /**
     * Set Visible
     *
     * @param bool $pValue
     * @return PHPExcel\Worksheet_RowDimension
     */
    public function setVisible($pValue = true) {
        $this->_visible = $pValue;
        return $this;
    }

    /**
     * Get Outline Level
     *
     * @return int
     */
    public function getOutlineLevel() {
        return $this->_outlineLevel;
    }

    /**
     * Set Outline Level
     *
     * Value must be between 0 and 7
     *
     * @param int $pValue
     * @throws PHPExcel\Exception
     * @return PHPExcel\Worksheet_RowDimension
     */
    public function setOutlineLevel($pValue) {
        if ($pValue < 0 || $pValue > 7) {
            throw new Exception("Outline level must range between 0 and 7.");
        }

        $this->_outlineLevel = $pValue;
        return $this;
    }

    /**
     * Get Collapsed
     *
     * @return bool
     */
    public function getCollapsed() {
        return $this->_collapsed;
    }

    /**
     * Set Collapsed
     *
     * @param bool $pValue
     * @return PHPExcel\Worksheet_RowDimension
     */
    public function setCollapsed($pValue = true) {
        $this->_collapsed = $pValue;
        return $this;
    }

    /**
     * Get index to cellXf
     *
     * @return int
     */
    public function getXfIndex()
    {
        return $this->_xfIndex;
    }

    /**
     * Set index to cellXf
     *
     * @param int $pValue
     * @return PHPExcel\Worksheet_RowDimension
     */
    public function setXfIndex($pValue = 0)
    {
        $this->_xfIndex = $pValue;
        return $this;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone() {
        $vars = get_object_vars($this);
        foreach ($vars as $key => $value) {
            if (is_object($value)) {
                $this->$key = clone $value;
            } else {
                $this->$key = $value;
            }
        }
    }
}
