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
 * PHPExcel\Worksheet_Protection
 *
 * @category   PHPExcel
 * @package    PHPExcel\Worksheet
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class Worksheet_Protection
{
    /**
     * Sheet
     *
     * @var boolean
     */
    protected $_sheet                    = false;

    /**
     * Objects
     *
     * @var boolean
     */
    protected $_objects                = false;

    /**
     * Scenarios
     *
     * @var boolean
     */
    protected $_scenarios                = false;

    /**
     * Format cells
     *
     * @var boolean
     */
    protected $_formatCells            = false;

    /**
     * Format columns
     *
     * @var boolean
     */
    protected $_formatColumns            = false;

    /**
     * Format rows
     *
     * @var boolean
     */
    protected $_formatRows            = false;

    /**
     * Insert columns
     *
     * @var boolean
     */
    protected $_insertColumns            = false;

    /**
     * Insert rows
     *
     * @var boolean
     */
    protected $_insertRows            = false;

    /**
     * Insert hyperlinks
     *
     * @var boolean
     */
    protected $_insertHyperlinks        = false;

    /**
     * Delete columns
     *
     * @var boolean
     */
    protected $_deleteColumns            = false;

    /**
     * Delete rows
     *
     * @var boolean
     */
    protected $_deleteRows            = false;

    /**
     * Select locked cells
     *
     * @var boolean
     */
    protected $_selectLockedCells        = false;

    /**
     * Sort
     *
     * @var boolean
     */
    protected $_sort                    = false;

    /**
     * AutoFilter
     *
     * @var boolean
     */
    protected $_autoFilter            = false;

    /**
     * Pivot tables
     *
     * @var boolean
     */
    protected $_pivotTables            = false;

    /**
     * Select unlocked cells
     *
     * @var boolean
     */
    protected $_selectUnlockedCells    = false;

    /**
     * Password
     *
     * @var string
     */
    protected $_password                = '';

    /**
     * Create a new PHPExcel\Worksheet_Protection
     */
    public function __construct()
    {
    }

    /**
     * Is some sort of protection enabled?
     *
     * @return boolean
     */
    function isProtectionEnabled() {
        return     $this->_sheet ||
                $this->_objects ||
                $this->_scenarios ||
                $this->_formatCells ||
                $this->_formatColumns ||
                $this->_formatRows ||
                $this->_insertColumns ||
                $this->_insertRows ||
                $this->_insertHyperlinks ||
                $this->_deleteColumns ||
                $this->_deleteRows ||
                $this->_selectLockedCells ||
                $this->_sort ||
                $this->_autoFilter ||
                $this->_pivotTables ||
                $this->_selectUnlockedCells;
    }

    /**
     * Get Sheet
     *
     * @return boolean
     */
    function getSheet() {
        return $this->_sheet;
    }

    /**
     * Set Sheet
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setSheet($pValue = false) {
        $this->_sheet = $pValue;
        return $this;
    }

    /**
     * Get Objects
     *
     * @return boolean
     */
    function getObjects() {
        return $this->_objects;
    }

    /**
     * Set Objects
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setObjects($pValue = false) {
        $this->_objects = $pValue;
        return $this;
    }

    /**
     * Get Scenarios
     *
     * @return boolean
     */
    function getScenarios() {
        return $this->_scenarios;
    }

    /**
     * Set Scenarios
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setScenarios($pValue = false) {
        $this->_scenarios = $pValue;
        return $this;
    }

    /**
     * Get FormatCells
     *
     * @return boolean
     */
    function getFormatCells() {
        return $this->_formatCells;
    }

    /**
     * Set FormatCells
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setFormatCells($pValue = false) {
        $this->_formatCells = $pValue;
        return $this;
    }

    /**
     * Get FormatColumns
     *
     * @return boolean
     */
    function getFormatColumns() {
        return $this->_formatColumns;
    }

    /**
     * Set FormatColumns
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setFormatColumns($pValue = false) {
        $this->_formatColumns = $pValue;
        return $this;
    }

    /**
     * Get FormatRows
     *
     * @return boolean
     */
    function getFormatRows() {
        return $this->_formatRows;
    }

    /**
     * Set FormatRows
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setFormatRows($pValue = false) {
        $this->_formatRows = $pValue;
        return $this;
    }

    /**
     * Get InsertColumns
     *
     * @return boolean
     */
    function getInsertColumns() {
        return $this->_insertColumns;
    }

    /**
     * Set InsertColumns
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setInsertColumns($pValue = false) {
        $this->_insertColumns = $pValue;
        return $this;
    }

    /**
     * Get InsertRows
     *
     * @return boolean
     */
    function getInsertRows() {
        return $this->_insertRows;
    }

    /**
     * Set InsertRows
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setInsertRows($pValue = false) {
        $this->_insertRows = $pValue;
        return $this;
    }

    /**
     * Get InsertHyperlinks
     *
     * @return boolean
     */
    function getInsertHyperlinks() {
        return $this->_insertHyperlinks;
    }

    /**
     * Set InsertHyperlinks
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setInsertHyperlinks($pValue = false) {
        $this->_insertHyperlinks = $pValue;
        return $this;
    }

    /**
     * Get DeleteColumns
     *
     * @return boolean
     */
    function getDeleteColumns() {
        return $this->_deleteColumns;
    }

    /**
     * Set DeleteColumns
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setDeleteColumns($pValue = false) {
        $this->_deleteColumns = $pValue;
        return $this;
    }

    /**
     * Get DeleteRows
     *
     * @return boolean
     */
    function getDeleteRows() {
        return $this->_deleteRows;
    }

    /**
     * Set DeleteRows
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setDeleteRows($pValue = false) {
        $this->_deleteRows = $pValue;
        return $this;
    }

    /**
     * Get SelectLockedCells
     *
     * @return boolean
     */
    function getSelectLockedCells() {
        return $this->_selectLockedCells;
    }

    /**
     * Set SelectLockedCells
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setSelectLockedCells($pValue = false) {
        $this->_selectLockedCells = $pValue;
        return $this;
    }

    /**
     * Get Sort
     *
     * @return boolean
     */
    function getSort() {
        return $this->_sort;
    }

    /**
     * Set Sort
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setSort($pValue = false) {
        $this->_sort = $pValue;
        return $this;
    }

    /**
     * Get AutoFilter
     *
     * @return boolean
     */
    function getAutoFilter() {
        return $this->_autoFilter;
    }

    /**
     * Set AutoFilter
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setAutoFilter($pValue = false) {
        $this->_autoFilter = $pValue;
        return $this;
    }

    /**
     * Get PivotTables
     *
     * @return boolean
     */
    function getPivotTables() {
        return $this->_pivotTables;
    }

    /**
     * Set PivotTables
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setPivotTables($pValue = false) {
        $this->_pivotTables = $pValue;
        return $this;
    }

    /**
     * Get SelectUnlockedCells
     *
     * @return boolean
     */
    function getSelectUnlockedCells() {
        return $this->_selectUnlockedCells;
    }

    /**
     * Set SelectUnlockedCells
     *
     * @param boolean $pValue
     * @return PHPExcel\Worksheet_Protection
     */
    function setSelectUnlockedCells($pValue = false) {
        $this->_selectUnlockedCells = $pValue;
        return $this;
    }

    /**
     * Get Password (hashed)
     *
     * @return string
     */
    function getPassword() {
        return $this->_password;
    }

    /**
     * Set Password
     *
     * @param string     $pValue
     * @param boolean     $pAlreadyHashed If the password has already been hashed, set this to true
     * @return PHPExcel\Worksheet_Protection
     */
    function setPassword($pValue = '', $pAlreadyHashed = false) {
        if (!$pAlreadyHashed) {
            $pValue = Shared_PasswordHasher::hashPassword($pValue);
        }
        $this->_password = $pValue;
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
