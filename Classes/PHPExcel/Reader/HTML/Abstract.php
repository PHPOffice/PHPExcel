<?php

/**
 * PHPExcel_Reader_HTML
 *
 * Generic reader of HTML files which sub classes can extend
 * to read HTML files to their need.
 *
 * Copyright (c) 2006 - 2015 PHPExcel
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
 * @package    PHPExcel_Reader_HTML
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @copyright  Copyright (c) 2015 Wine Logistix (http://www.wine-logistix.de)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
abstract class PHPExcel_Reader_HTML_Abstract extends PHPExcel_Reader_Abstract implements PHPExcel_Reader_IReader
{

    /**
     * Write cell content at specified position to active sheet.
     * @param int $row
     * @param string $column
     * @param string $cellContent
     */
    protected abstract function flushCell($row, $column, &$cellContent);

    /**
     * Handler for elements with no explicit handler.
     * @param \DOMNode $element
     * @param int $row
     * @param string $column
     * @param string $cellContent
     */
    protected abstract function defaultElementHandler(\DOMNode $element, &$row, &$column, &$cellContent);

    /**
     * Handler for DOMText elements.
     * @param \DOMNode $element
     * @param int $row
     * @param string $column
     * @param string $cellContent
     */
    protected abstract function textElementHandler(\DOMNode $element, &$row, &$column, &$cellContent);

    /**
     * Handler which is executed after loading the HTML file and before
     * traversing elements.
     * @param \PHPExcel $objPHPExcel
     */
    protected abstract function loadHandler(\PHPExcel $objPHPExcel);

    /**
     * Loads PHPExcel from file.
     * @param  string $pFilename
     * @return PHPExcel
     * @throws PHPExcel_Reader_Exception
     */
    public function load($pFilename)
    {
        // Create new PHPExcel
        $objPHPExcel = new PHPExcel();
        // Load into this instance
        return $this->loadIntoExisting($pFilename, $objPHPExcel);
    }

    /**
     * Loads PHPExcel from file into PHPExcel instance.
     *
     * @param  string                    $pFilename
     * @param  PHPExcel                  $objPHPExcel
     * @return PHPExcel
     * @throws PHPExcel_Reader_Exception
     */
    public function loadIntoExisting($pFilename, PHPExcel $objPHPExcel)
    {
        // Open file to validate
        $this->openFile($pFilename);
        if (!$this->isValidFormat()) {
            fclose($this->fileHandle);
            throw new PHPExcel_Reader_Exception($pFilename . " is an invalid HTML file.");
        }
        //    Close after validating
        fclose($this->fileHandle);

        //    Create a new DOM object
        $dom = new DOMDocument();
        //    Reload the HTML file into the DOM object
        $loaded = $dom->loadHTML($this->securityScanFile($pFilename));
        if ($loaded === false) {
            throw new PHPExcel_Reader_Exception('Failed to load ', $pFilename, ' as a DOM Document');
        }

        //    Discard white space
        $dom->preserveWhiteSpace = false;

        $row = 0;
        $column = 'A';
        $content = '';

        // Allow implementation specific initalization after load.
        $this->loadHandler($objPHPExcel);

        $this->processDomElement($dom, $row, $column, $content);

        // Return
        return $objPHPExcel;
    }

    /**
     * Validate that the current file is an HTML file
     *
     * @return boolean
     */
    protected function isValidFormat()
    {
        //    Reading 2048 bytes should be enough to validate that the format is HTML
        $data = fread($this->fileHandle, 2048);
        if ((strpos($data, '<') !== false) &&
                (strlen($data) !== strlen(strip_tags($data)))) {
            return true;
        }

        return false;
    }

    protected function processDomElement(DOMNode $element, &$row, &$column, &$cellContent)
    {
        foreach ($element->childNodes as $child) {
            if ($child instanceof \DOMText) {
                $this->textElementHandler($child, $row, $column, $cellContent);
            } elseif ($child instanceof \DOMElement) {
                // For each element a handler is invoked dynamically. If you don't want to use
                // dynamic dispatch, use defaultElementHandler.
                $nodeName = $this->cleanNodeName($child->nodeName);
                $handlerName = $nodeName . "ElementHandler";
                if (method_exists($this, $handlerName)) {
                    $this->{$handlerName}($child, $row, $column, $cellContent);
                } else {
                    $this->defaultElementHandler($child, $row, $column, $cellContent);
                }
            }
        }
    }

    protected function cleanNodeName($elementName)
    {
        return strtolower(preg_replace('/[^a-zA-Z0-9]/u', '', $elementName));
    }
}
