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
 * PHPExcel\CachedObjectStorage_Memory
 *
 * @category   PHPExcel
 * @package    PHPExcel\CachedObjectStorage
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class CachedObjectStorage_Memory extends CachedObjectStorage_CacheBase implements CachedObjectStorage_ICache
{
    /**
     * Dummy method callable from CacheBase, but unused by Memory cache
     *
     * @return    void
     */
    protected function storeData()
    {
    }

    /**
     * Add or Update a cell in cache identified by coordinate address
     *
     * @param    string            $pCoord        Coordinate address of the cell to update
     * @param    PHPExcel\Cell    $cell        Cell to update
     * @return    PHPExcel\Cell
     * @throws    PHPExcel\Exception
     */
    public function addCacheData($pCoord, Cell $cell)
    {
        $this->cellCache[$pCoord] = $cell;

        //    Set current entry to the new/updated entry
        $this->currentObjectID = $pCoord;

        return $cell;
    }

    /**
     * Get cell at a specific coordinate
     *
     * @param     string             $pCoord        Coordinate of the cell
     * @throws     PHPExcel\Exception
     * @return     PHPExcel\Cell     Cell that was found, or null if not found
     */
    public function getCacheData($pCoord)
    {
        //    Check if the entry that has been requested actually exists
        if (!isset($this->cellCache[$pCoord])) {
            $this->currentObjectID = null;
            //    Return null if requested entry doesn't exist in cache
            return null;
        }

        //    Set current entry to the requested entry
        $this->currentObjectID = $pCoord;

        //    Return requested entry
        return $this->cellCache[$pCoord];
    }

    /**
     * Clone the cell collection
     *
     * @param    PHPExcel\Worksheet    $parent        The new worksheet
     * @return    void
     */
    public function copyCellCollection(Worksheet $parent)
    {
        parent::copyCellCollection($parent);

        $newCollection = array();
        foreach ($this->cellCache as $k => &$cell) {
            $newCollection[$k] = clone $cell;
            $newCollection[$k]->attach($this);
        }

        $this->cellCache = $newCollection;
    }

    /**
     * Clear the cell collection and disconnect from our parent
     *
     * @return    void
     */
    public function unsetWorksheetCells()
    {
        //    Because cells are all stored as intact objects in memory, we need to detach each one from the parent
        foreach ($this->cellCache as $k => &$cell) {
            $cell->detach();
            $this->cellCache[$k] = null;
        }
        unset($cell);

        $this->cellCache = array();

        //    detach ourself from the worksheet, so that it can then delete this object successfully
        $this->parent = null;
    }
}
