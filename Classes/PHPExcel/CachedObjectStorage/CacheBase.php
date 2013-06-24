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
 * @package    PHPExcel\CachedObjectStorage
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */


namespace PHPExcel;
 
/**
 * PHPExcel\CachedObjectStorage_CacheBase
 *
 * @category   PHPExcel
 * @package    PHPExcel\CachedObjectStorage
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
abstract class CachedObjectStorage_CacheBase
{
    /**
     * Parent worksheet
     *
     * @var PHPExcel\Worksheet
     */
    protected $parent;

    /**
     * The currently active Cell
     *
     * @var PHPExcel\Cell
     */
    protected $currentObject = null;

    /**
     * Coordinate address of the currently active Cell
     *
     * @var string
     */
    protected $currentObjectID = null;

    /**
     * Flag indicating whether the currently active Cell requires saving
     *
     * @var boolean
     */
    protected $currentCellIsDirty = true;

    /**
     * An array of cells or cell pointers for the worksheet cells held in this cache,
     *        and indexed by their coordinate address within the worksheet
     *
     * @var array of mixed
     */
    protected $cellCache = array();


    /**
     * Initialise this new cell collection
     *
     * @param    PHPExcel\Worksheet    $parent        The worksheet for this cell collection
     */
    public function __construct(Worksheet $parent)
    {
        //    Set our parent worksheet.
        //    This is maintained within the cache controller to facilitate re-attaching it to PHPExcel\Cell objects when
        //        they are woken from a serialized state
        $this->parent = $parent;
    }

    /**
     * Return the parent worksheet for this cell collection
     *
     * @return    PHPExcel\Worksheet
     */
    public function getParent()
    {
        return $this->parent;
    }

    /**
     * Is a value set in the current PHPExcel\CachedObjectStorage_ICache for an indexed cell?
     *
     * @param    string        $pCoord        Coordinate address of the cell to check
     * @return    boolean
     */
    public function isDataSet($pCoord)
    {
        if ($pCoord === $this->currentObjectID) {
            return true;
        }
        //    Check if the requested entry exists in the cache
        return isset($this->cellCache[$pCoord]);
    }

    /**
     * Move a cell object from one address to another
     *
     * @param    string        $fromAddress    Current address of the cell to move
     * @param    string        $toAddress        Destination address of the cell to move
     * @return    boolean
     */
    public function moveCell($fromAddress, $toAddress)
    {
        if ($fromAddress === $this->currentObjectID) {
            $this->currentObjectID = $toAddress;
        }
        $this->currentCellIsDirty = true;
        if (isset($this->cellCache[$fromAddress])) {
            $this->cellCache[$toAddress] = &$this->cellCache[$fromAddress];
            unset($this->cellCache[$fromAddress]);
        }

        return true;
    }

    /**
     * Add or Update a cell in cache
     *
     * @param    PHPExcel\Cell    $cell        Cell to update
     * @return    void
     * @throws    PHPExcel\Exception
     */
    public function updateCacheData(Cell $cell)
    {
        return $this->addCacheData($cell->getCoordinate(), $cell);
    }

    /**
     * Delete a cell in cache identified by coordinate address
     *
     * @param    string            $pCoord        Coordinate address of the cell to delete
     * @throws    PHPExcel\Exception
     */
    public function deleteCacheData($pCoord)
    {
        if ($pCoord === $this->currentObjectID) {
            $this->currentObject->detach();
            $this->currentObjectID = $this->currentObject = null;
        }

        if (is_object($this->cellCache[$pCoord])) {
            $this->cellCache[$pCoord]->detach();
            unset($this->cellCache[$pCoord]);
        }
        $this->currentCellIsDirty = false;
    }

    /**
     * Get a list of all cell addresses currently held in cache
     *
     * @return    array of string
     */
    public function getCellList()
    {
        return array_keys($this->cellCache);
    }

    /**
     * Sort the list of all cell addresses currently held in cache by row and column
     *
     * @return    void
     */
    public function getSortedCellList()
    {
        $sortKeys = array();
        foreach ($this->getCellList() as $coord) {
            sscanf($coord, '%[A-Z]%d', $column, $row);
            $sortKeys[sprintf('%09d%3s', $row, $column)] = $coord;
        }
        ksort($sortKeys);

        return array_values($sortKeys);
    }

    /**
     * Get highest worksheet column and highest row that have cell records
     *
     * @return array Highest column name and highest row number
     */
    public function getHighestRowAndColumn()
    {
        // Lookup highest column and highest row
        $col = array('A' => '1A');
        $row = array(1);
        foreach ($this->getCellList() as $coord) {
            sscanf($coord, '%[A-Z]%d', $c, $r);
            $row[$r] = $r;
            $col[$c] = strlen($c).$c;
        }
        if (!empty($row)) {
            // Determine highest column and row
            $highestRow = max($row);
            $highestColumn = substr(max($col), 1);
        }

        return array(
            'row'    => $highestRow,
            'column' => $highestColumn
        );
    }

    /**
     * Return the cell address of the currently active cell object
     *
     * @return    string
     */
    public function getCurrentAddress()
    {
        return $this->currentObjectID;
    }

    /**
     * Return the column address of the currently active cell object
     *
     * @return    string
     */
    public function getCurrentColumn()
    {
        sscanf($this->currentObjectID, '%[A-Z]%d', $column, $row);
        return $column;
    }

    /**
     * Return the row address of the currently active cell object
     *
     * @return    string
     */
    public function getCurrentRow()
    {
        sscanf($this->currentObjectID, '%[A-Z]%d', $column, $row);
        return $row;
    }

    /**
     * Get highest worksheet column
     *
     * @return string Highest column name
     */
    public function getHighestColumn()
    {
        $colRow = $this->getHighestRowAndColumn();
        return $colRow['column'];
    }

    /**
     * Get highest worksheet row
     *
     * @return int Highest row number
     */
    public function getHighestRow()
    {
        $colRow = $this->getHighestRowAndColumn();
        return $colRow['row'];
    }

    /**
     * Generate a unique ID for cache referencing
     *
     * @return string Unique Reference
     */
    protected function getUniqueID()
    {
        if (function_exists('posix_getpid')) {
            $baseUnique = posix_getpid();
        } else {
            $baseUnique = mt_rand();
        }
        return uniqid($baseUnique, true);
    }

    /**
     * Clone the cell collection
     *
     * @param    PHPExcel\Worksheet    $parent        The new worksheet
     * @return    void
     */
    public function copyCellCollection(Worksheet $parent)
    {
        $this->currentCellIsDirty;
        $this->storeData();

        $this->parent = $parent;
        if (($this->currentObject !== null) && (is_object($this->currentObject))) {
            $this->currentObject->attach($this);
        }
    }

    /**
     * Identify whether the caching method is currently available
     * Some methods are dependent on the availability of certain extensions being enabled in the PHP build
     *
     * @return    boolean
     */
    public static function cacheMethodIsAvailable()
    {
        return true;
    }
}
