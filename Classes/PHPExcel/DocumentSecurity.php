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
 * PHPExcel\DocumentSecurity
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class DocumentSecurity
{
    /**
     * LockRevision
     *
     * @var boolean
     */
    protected $_lockRevision = false;

    /**
     * LockStructure
     *
     * @var boolean
     */
    protected $_lockStructure = false;

    /**
     * LockWindows
     *
     * @var boolean
     */
    protected $_lockWindows = false;

    /**
     * RevisionsPassword
     *
     * @var string
     */
    protected $_revisionsPassword = '';

    /**
     * WorkbookPassword
     *
     * @var string
     */
    protected $_workbookPassword = '';

    /**
     * Create a new PHPExcel\DocumentSecurity
     */
    public function __construct()
    {
    }

    /**
     * Is some sort of dcument security enabled?
     *
     * @return boolean
     */
    function isSecurityEnabled() {
        return     $this->_lockRevision ||
                $this->_lockStructure ||
                $this->_lockWindows;
    }

    /**
     * Get LockRevision
     *
     * @return boolean
     */
    function getLockRevision() {
        return $this->_lockRevision;
    }

    /**
     * Set LockRevision
     *
     * @param boolean $pValue
     * @return PHPExcel\DocumentSecurity
     */
    function setLockRevision($pValue = false) {
        $this->_lockRevision = $pValue;
        return $this;
    }

    /**
     * Get LockStructure
     *
     * @return boolean
     */
    function getLockStructure() {
        return $this->_lockStructure;
    }

    /**
     * Set LockStructure
     *
     * @param boolean $pValue
     * @return PHPExcel\DocumentSecurity
     */
    function setLockStructure($pValue = false) {
        $this->_lockStructure = $pValue;
        return $this;
    }

    /**
     * Get LockWindows
     *
     * @return boolean
     */
    function getLockWindows() {
        return $this->_lockWindows;
    }

    /**
     * Set LockWindows
     *
     * @param boolean $pValue
     * @return PHPExcel\DocumentSecurity
     */
    function setLockWindows($pValue = false) {
        $this->_lockWindows = $pValue;
        return $this;
    }

    /**
     * Get RevisionsPassword (hashed)
     *
     * @return string
     */
    function getRevisionsPassword() {
        return $this->_revisionsPassword;
    }

    /**
     * Set RevisionsPassword
     *
     * @param string     $pValue
     * @param boolean     $pAlreadyHashed If the password has already been hashed, set this to true
     * @return PHPExcel\DocumentSecurity
     */
    function setRevisionsPassword($pValue = '', $pAlreadyHashed = false) {
        if (!$pAlreadyHashed) {
            $pValue = Shared_PasswordHasher::hashPassword($pValue);
        }
        $this->_revisionsPassword = $pValue;
        return $this;
    }

    /**
     * Get WorkbookPassword (hashed)
     *
     * @return string
     */
    function getWorkbookPassword() {
        return $this->_workbookPassword;
    }

    /**
     * Set WorkbookPassword
     *
     * @param string     $pValue
     * @param boolean     $pAlreadyHashed If the password has already been hashed, set this to true
     * @return PHPExcel\DocumentSecurity
     */
    function setWorkbookPassword($pValue = '', $pAlreadyHashed = false) {
        if (!$pAlreadyHashed) {
            $pValue = Shared_PasswordHasher::hashPassword($pValue);
        }
        $this->_workbookPassword = $pValue;
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
