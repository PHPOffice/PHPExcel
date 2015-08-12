<?php

/**
 * PHPExcel_Reader_HTML_Abstract
 *
 * Generic reader of HTML files which sub classes can extend
 * to read HTML files to their need.
 *
 * When loading a document, the DOM is traversed with its top-level elements.
 * A handler is invoked for each element. By default it is defaultElementHandler,
 * though explicit handlers may be defined in the subclass in form
 * <element_name>ElementHandler where <element_name> is lowercase element name.
 * Explicit handlers must accept same arguments as defaultElementHandler.
 *
 * Other handlers exist which facilitate implementation specific behavior:
 *
 * flushCell - Write a cell value
 * textElementHandler - Invoked for DOMText elements.
 * loadHandler - Invoked before traversing the DOM.
 * finishHandler - Invoked after traversing the DOM.
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
     * Tell processDomElement to traverse child elements of the current child
     * element recursively.
     * @var int
     */
    const TRAVERSE_CHILDS = 1;

    /**
     * Write cell content at specified position to active sheet.
     * @param string $column
     * @param int $row
     * @param string $cellContent
     */
    protected abstract function flushCell($column, $row, &$cellContent);

    /**
     * Handler for elements with no explicit handler.
     * @param DOMNode $element
     * @param int $row
     * @param string $column
     * @param string $cellContent
     * @return int|null TRAVERSE_CHILDS or null
     */
    protected abstract function defaultElementHandler(DOMNode $element, &$row, &$column, &$cellContent);

    /**
     * Handler for DOMText elements.
     * @param DOMNode $element
     * @param int $row
     * @param string $column
     * @param string $cellContent
     */
    protected abstract function textElementHandler(DOMNode $element, &$row, &$column, &$cellContent);

    /**
     * Handler which is executed after loading the HTML file and before
     * traversing elements.
     * @param PHPExcel $objPHPExcel
     */
    protected abstract function loadHandler(PHPExcel $objPHPExcel);

    /**
     * Handler which is executed after traversing elements and before
     * returning from load method.
     */
    protected abstract function finishHandler();

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
        // Open file to validate
        $this->openFile($pFilename);
        if (!$this->isValidFileFormat()) {
            fclose($this->fileHandle);
            throw new PHPExcel_Reader_Exception($pFilename . " is an invalid HTML file.");
        }
        //    Close after validating
        fclose($this->fileHandle);
        // Load into this instance
        return $this->loadIntoExisting($pFilename, $objPHPExcel);
    }

    /**
     * Loads PHPExcel from string.
     * @param string $content HTML content
     * @return PHPExcel
     * @throws PHPExcel_Reader_Exception
     */
    public function loadFromString($content)
    {
        $objPHPExcel = new PHPExcel();
        if (!$this->isValidFormat($content)) {
            throw new PHPExcel_Reader_Exception("HTML content is invalid");
        }
        $html = $this->securityScan($content);
        return $this->loadIntoExistingFromString($html, $objPHPExcel);
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
        $html = $this->securityScanFile($pFilename);
        return $this->loadIntoExistingFromString($html, $objPHPExcel);
    }

    /**
     * Loads PHPExcel from string into PHPExcel instance.
     */
    protected function loadIntoExistingFromString($content, PHPExcel $objPHPExcel)
    {
        // This method is protected as it doesn't do the security scan on content.
        //    Create a new DOM object
        $dom = new DOMDocument();
        //    Reload the HTML file into the DOM object
        $loaded = $dom->loadHTML($content);
        if ($loaded === false) {
            throw new PHPExcel_Reader_Exception('Failed to load ', $pFilename, ' as a DOM Document');
        }

        //    Discard white space
        $dom->preserveWhiteSpace = false;

        $row = 1;
        $column = 'A';
        $content = '';

        // Allow implementation specific initalization after load.
        $this->loadHandler($objPHPExcel);

        $this->processDomElement($dom, $row, $column, $content);

        // Allow implementation specific operation after processing.
        $this->finishHandler();

        // Return
        return $objPHPExcel;
    }

    /**
     * Validate that data contains HTML.
     * @return boolean
     */
    protected function isValidFormat(&$data)
    {
        if ((strpos($data, '<') !== false) &&
                (strlen($data) !== strlen(strip_tags($data)))) {
            return true;
        }
        return false;
    }

    /**
     * Validate that the current file is an HTML file
     *
     * @return boolean
     */
    protected function isValidFileFormat()
    {
        //    Reading 2048 bytes should be enough to validate that the format is HTML
        $data = fread($this->fileHandle, 2048);
        return $this->isValidFormat($data);
    }

    /**
     * Traverse elements in DOM and invoke handler.
     * A handler method in own object with name <element_name>ElementHandler
     * is invoked if the method exists, or defaultElementHandler if not.
     * Handlers can indicate whether to traverse child elements, by returning
     * TRAVERSE_CHILDS. Childs are traversed recursively.
     * @param DOMNode $element Element of which childs are traversed.
     * @param int $row Row number
     * @param string $column Excel style column name
     * @param $cellContent A buffer which can be used by implementation to store temporary cell content before flushing to cell.
     */
    protected function processDomElement(DOMNode $element, &$row, &$column, &$cellContent)
    {
        foreach ($element->childNodes as $child) {
            if ($child instanceof DOMText) {
                $this->textElementHandler($child, $row, $column, $cellContent);
            } elseif ($child instanceof DOMElement) {
                // For each element a handler is invoked dynamically. If you
                // don't want to use dynamic dispatch, use defaultElementHandler.
                $nodeName = $this->cleanNodeName($child->nodeName);
                $handlerName = $nodeName . "ElementHandler";
                $continueWith = (method_exists($this, $handlerName)
                    ? $this->{$handlerName}($child, $row, $column, $cellContent)
                    : $this->defaultElementHandler($child, $row, $column, $cellContent));
                if ($continueWith === self::TRAVERSE_CHILDS && $child->hasChildNodes()) {
                    // Handlers may traverse the DOM themselves. To avoid
                    // unnecessary traversing in here, by default no childs of
                    // the child are traversed. If however indicated by handler
                    // to traverse childs, then do so.
                    $this->processDomElement($child, $row, $column, $cellContent);
                }
            }
        }
    }

    protected function cleanNodeName($elementName)
    {
        return strtolower(preg_replace('/[^a-zA-Z0-9]/u', '', $elementName));
    }

    /**
     * Scan theXML for use of <!ENTITY to prevent XXE/XEE attacks
     *
     * @param     string         $xml
     * @throws PHPExcel_Reader_Exception
     */
    public function securityScan($xml)
    {
        $pattern = '/\\0?' . implode('\\0?', str_split('<!ENTITY')) . '\\0?/';
        if (preg_match($pattern, $xml)) {
            throw new PHPExcel_Reader_Exception('Detected use of ENTITY in XML, spreadsheet file load() aborted to prevent XXE/XEE attacks');
        }
        return $xml;
    }
}
