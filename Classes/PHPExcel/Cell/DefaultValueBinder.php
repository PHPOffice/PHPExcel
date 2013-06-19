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
 * @package    PHPExcel\Cell
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */


namespace PHPExcel;

/**
 * PHPExcel\Cell_DefaultValueBinder
 *
 * @category   PHPExcel
 * @package    PHPExcel\Cell
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class Cell_DefaultValueBinder implements Cell_IValueBinder
{
    /**
     * Bind value to a cell
     *
     * @param  PHPExcel\Cell  $cell   Cell to bind value to
     * @param  mixed          $value  Value to bind in cell
     * @return boolean
     */
    public function bindValue(Cell $cell, $value = null)
    {
        // sanitize UTF-8 strings
        if (is_string($value)) {
            $value = Shared_String::SanitizeUTF8($value);
        }

        // Set value explicit
        $cell->setValueExplicit( $value, self::dataTypeForValue($value) );

        // Done!
        return TRUE;
    }

    /**
     * DataType for value
     *
     * @param   mixed  $pValue
     * @return  string
     */
    public static function dataTypeForValue($pValue = null) {
        // Match the value against a few data types
        if (is_null($pValue)) {
            return Cell_DataType::TYPE_NULL;

        } elseif ($pValue === '') {
            return Cell_DataType::TYPE_STRING;

        } elseif ($pValue instanceof RichText) {
            return Cell_DataType::TYPE_INLINE;

        } elseif ($pValue{0} === '=' && strlen($pValue) > 1) {
            return Cell_DataType::TYPE_FORMULA;

        } elseif (is_bool($pValue)) {
            return Cell_DataType::TYPE_BOOL;

        } elseif (is_float($pValue) || is_int($pValue)) {
            return Cell_DataType::TYPE_NUMERIC;

        } elseif (preg_match('/^\-?([0-9]+\\.?[0-9]*|[0-9]*\\.?[0-9]+)$/', $pValue)) {
            return Cell_DataType::TYPE_NUMERIC;

        } elseif (is_string($pValue) && array_key_exists($pValue, Cell_DataType::getErrorCodes())) {
            return Cell_DataType::TYPE_ERROR;

        } else {
            return Cell_DataType::TYPE_STRING;

        }
    }
}
