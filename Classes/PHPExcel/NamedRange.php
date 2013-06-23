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
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */


namespace PHPExcel;

/**
 * PHPExcel\NamedRange
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class NamedRange
{
    /**
     * Range name
     *
     * @var string
     */
    protected $_name;

    /**
     * Worksheet on which the named range can be resolved
     *
     * @var PHPExcel\Worksheet
     */
    protected $_worksheet;

    /**
     * Range of the referenced cells
     *
     * @var string
     */
    protected $_range;

    /**
     * Is the named range local? (i.e. can only be used on $this->_worksheet)
     *
     * @var bool
     */
    protected $_localOnly;

    /**
     * Scope
     *
     * @var PHPExcel\Worksheet
     */
    protected $_scope;

    /**
     * Create a new NamedRange
     *
     * @param string $pName
     * @param PHPExcel\Worksheet $pWorksheet
     * @param string $pRange
     * @param bool $pLocalOnly
     * @param PHPExcel\Worksheet|null $pScope    Scope. Only applies when $pLocalOnly = true. Null for global scope.
     * @throws PHPExcel\Exception
     */
    public function __construct($pName = null, Worksheet $pWorksheet, $pRange = 'A1', $pLocalOnly = false, $pScope = null)
    {
        // Validate data
        if (($pName === null) || ($pWorksheet === null) || ($pRange === null)) {
            throw new Exception('Parameters can not be null.');
        }

        // Set local members
        $this->_name         = $pName;
        $this->_worksheet     = $pWorksheet;
        $this->_range         = $pRange;
        $this->_localOnly     = $pLocalOnly;
        $this->_scope         = ($pLocalOnly == true) ?
                                (($pScope == null) ? $pWorksheet : $pScope) : null;
    }

    /**
     * Get name
     *
     * @return string
     */
    public function getName() {
        return $this->_name;
    }

    /**
     * Set name
     *
     * @param string $value
     * @return PHPExcel\NamedRange
     */
    public function setName($value = null) {
        if ($value !== null) {
            // Old title
            $oldTitle = $this->_name;

            // Re-attach
            if ($this->_worksheet !== null) {
                $this->_worksheet->getParent()->removeNamedRange($this->_name,$this->_worksheet);
            }
            $this->_name = $value;

            if ($this->_worksheet !== null) {
                $this->_worksheet->getParent()->addNamedRange($this);
            }

            // New title
            $newTitle = $this->_name;
            ReferenceHelper::getInstance()->updateNamedFormulas($this->_worksheet->getParent(), $oldTitle, $newTitle);
        }
        return $this;
    }

    /**
     * Get worksheet
     *
     * @return PHPExcel\Worksheet
     */
    public function getWorksheet() {
        return $this->_worksheet;
    }

    /**
     * Set worksheet
     *
     * @param PHPExcel\Worksheet $worksheet
     * @return PHPExcel\NamedRange
     */
    public function setWorksheet(Worksheet $worksheet = null) {
        if ($worksheet !== null) {
            $this->_worksheet = $worksheet;
        }
        return $this;
    }

    /**
     * Get range
     *
     * @return string
     */
    public function getRange() {
        return $this->_range;
    }

    /**
     * Set range
     *
     * @param string $value
     * @return PHPExcel\NamedRange
     */
    public function setRange($value = null) {
        if ($value !== null) {
            $this->_range = $value;
        }
        return $this;
    }

    /**
     * Get localOnly
     *
     * @return bool
     */
    public function getLocalOnly() {
        return $this->_localOnly;
    }

    /**
     * Set localOnly
     *
     * @param bool $value
     * @return PHPExcel\NamedRange
     */
    public function setLocalOnly($value = false) {
        $this->_localOnly = $value;
        $this->_scope = $value ? $this->_worksheet : null;
        return $this;
    }

    /**
     * Get scope
     *
     * @return PHPExcel\Worksheet|null
     */
    public function getScope() {
        return $this->_scope;
    }

    /**
     * Set scope
     *
     * @param PHPExcel\Worksheet|null $worksheet
     * @return PHPExcel\NamedRange
     */
    public function setScope(Worksheet $worksheet = null) {
        $this->_scope = $worksheet;
        $this->_localOnly = ($worksheet === null) ? false : true;
        return $this;
    }

    /**
     * Resolve a named range to a regular cell range
     *
     * @param string $pNamedRange Named range
     * @param PHPExcel\Worksheet|null $pSheet Scope. Use null for global scope
     * @return PHPExcel\NamedRange
     */
    public static function resolveRange($pNamedRange = '', Worksheet $pSheet) {
        return $pSheet->getParent()->getNamedRange($pNamedRange, $pSheet);
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
