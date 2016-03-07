<?php

namespace PhpOffice\PhpExcel\CachedObjectStorage;

/**
 * PHPExcel_CachedObjectStorage_Memory
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
 * @package    PHPExcel_CachedObjectStorage
 * @copyright  Copyright (c) 2006 - 2016 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Memory extends CacheBase implements ICache
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
     * @param   string            $pCoord        Coordinate address of the cell to update
     * @param   \PhpOffice\PhpExcel\Cell    $cell        Cell to update
     * @return  \PhpOffice\PhpExcel\Cell
     * @throws  \PhpOffice\PhpExcel\Exception
     */
    public function addCacheData($pCoord, \PhpOffice\PhpExcel\Cell $cell)
    {
        $this->cellCache[$pCoord] = $cell;

        //    Set current entry to the new/updated entry
        $this->currentObjectID = $pCoord;

        return $cell;
    }


    /**
     * Get cell at a specific coordinate
     *
     * @param   string             $pCoord        Coordinate of the cell
     * @throws  \PhpOffice\PhpExcel\Exception
     * @return  \PhpOffice\PhpExcel\Cell     Cell that was found, or null if not found
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
     * @param  \PhpOffice\PhpExcel\Worksheet    $parent        The new worksheet that we're copying to
     */
    public function copyCellCollection(\PhpOffice\PhpExcel\Worksheet $parent)
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
     */
    public function unsetWorksheetCells()
    {
        // Because cells are all stored as intact objects in memory, we need to detach each one from the parent
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
