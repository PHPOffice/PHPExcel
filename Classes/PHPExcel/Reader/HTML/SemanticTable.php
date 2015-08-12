<?php
/**
 * PHPExcel_Reader_HTML_SemanticTable
 *
 * Parse each table element in HTML documents as a worksheet and honor
 * semantic markup.
 *
 * The elements are treated as:
 * title - Document title
 * table - worksheet
 * table > caption - worksheet title
 * table > thead - Header rows (formatted bold)
 * table > tbody - Data rows (no formatting)
 *
 * Copyright (c) 2015 Wine Logistix GmbH
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
 * @package    PHPExcel_Reader_HTML_SemanticTable
 * @copyright  Copyright (c) 2015 Wine Logistix (http://www.wine-logistix.de)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class PHPExcel_Reader_HTML_SemanticTable extends PHPExcel_Reader_HTML_Abstract
{

    /**
     * @var \PHPExcel
     */
    protected $excel;

    /**
     * Write cell content at specified position to active sheet.
     * @param int $row
     * @param string $column
     * @param string $cellContent
     */
    protected function flushCell($column, $row, &$cellContent)
    {
        if (is_string($cellContent)) {
            $cellContent = trim($cellContent);
            if ($cellContent !== '') {
                $this->excel->getActiveSheet()->setCellValue($column.$row, $cellContent);
            }
        }
    }

    /**
     * Handler for elements with no explicit handler.
     * @param \DOMNode $element
     * @param int $row
     * @param string $column
     * @param string $cellContent
     */
    protected function defaultElementHandler(\DOMNode $element, &$row, &$column, &$cellContent)
    {
        // This implementation doesn't care about any element except the ones
        // for which an explicit handler is defined. To get to these elements
        // though, children of the other elements need to be traversed.
        return \PHPExcel_Reader_HTML_Abstract::TRAVERSE_CHILDS;
    }

    /**
     * Handler for DOMText elements.
     * @param \DOMNode $element
     * @param int $row
     * @param string $column
     * @param string $cellContent
     */
    protected function textElementHandler(\DOMNode $element, &$row, &$column, &$cellContent)
    {
    }

    /**
     * Handler which is executed after loading the HTML file and before
     * traversing elements.
     * @param \PHPExcel $objPHPExcel
     */
    protected function loadHandler(\PHPExcel $objPHPExcel)
    {
        $this->excel = $objPHPExcel;
        // Remove first sheet because if no table elements are occured
        // in document, then it's an error in the source file.
        $this->excel->removeSheetByIndex(0);
    }

    protected function finishHandler()
    {
        if ($this->excel->getSheetCount() > 0) {
            // This is cosmetic; during processing a worksheet was created
            // for each table and the last created is set active. When opening
            // the file in GUI, the last worksheet would open, but it's most
            // likely desired to view the first worksheet first.
            $this->excel->setActiveSheetIndex(0);
        }
    }

    /**
     * Set document title.
     * @param \DOMNode $element
     * @param int $row
     * @param string $column
     * @param string $cellContent
     */
    protected function titleElementHandler(\DOMNode $element, &$row, &$column, &$cellContent)
    {
        $this->excel->getProperties()->setTitle($element->textContent);
    }

    /**
     * Create a new worksheet and use it as active sheet.
     * @param \DOMNode $element
     * @param int $row
     * @param string $column
     * @param string $cellContent
     */
    protected function tableElementHandler(\DOMNode $element, &$row, &$column, &$cellContent)
    {
        $sheetNum = $this->excel->getSheetCount();
        $this->excel->createSheet();
        $this->excel->setActiveSheetIndex($sheetNum);
        // Row and column need to be reset.
        $row = 1;
        $column = 'A';
        return PHPExcel_Reader_HTML_Abstract::TRAVERSE_CHILDS;
    }

    /**
     * Set title of current active sheet.
     * @param \DOMNode $element
     * @param int $row
     * @param string $column
     * @param string $cellContent
     */
    protected function captionElementHandler(\DOMNode $element, &$row, &$column, &$cellContent)
    {
        $this->excel->getActiveSheet()->setTitle($element->textContent);
    }

    /**
     * For each header row in thead, create a row with bold formatted columns.
     * @param \DOMNode $element
     * @param int $row
     * @param string $column
     * @param string $cellContent
     */
    protected function theadElementHandler(\DOMNode $element, &$row, &$column, &$cellContent)
    {
        foreach ($element->childNodes as $child) {
            if ($this->isElement($child, "tr")) {
                $this->createHeaderRow($child, $row);
                $row += 1;
            }
        }
        // Don't traverse childs as they are already traversed in here.
    }

    protected function tbodyElementHandler(\DOMNode $element, &$row, &$column, &$cellContent)
    {
        foreach ($element->childNodes as $child) {
            if ($this->isElement($child, "tr")) {
                $this->createDataRow($child, $row);
                $row += 1;
            }
        }
        // Don't traverse childs as they are already traversed in here.
    }

    protected function createHeaderRow(\DOMNode $theadRow, $row)
    {
        $column = 'A';
        foreach ($theadRow->childNodes as $child) {
            if ($this->isElement($child, "th")) {
                $this->flushCell($column, $row, $child->textContent);
                $column++;
            }
        }
        // Formatting headers by using range is faster than doing it in the loop.
        $range = sprintf('A%d:%s%d', $row, $column, $row);
        $this->excel->getActiveSheet()->getStyle($range)->getFont()->setBold(true);
    }

    protected function createDataRow(\DOMNode $tbodyRow, $row)
    {
        $column = 'A';
        foreach ($tbodyRow->childNodes as $child) {
            if ($this->isElement($child, "td")) {
                $this->flushCell($column, $row, $child->textContent);
                $column++;
            }
        }
    }

    private function isElement($el, $name) {
        return $el instanceof \DOMNode && $el->nodeName === $name;
    }

}
