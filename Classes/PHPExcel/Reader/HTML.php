<?php

if (!defined('PHPEXCEL_ROOT')) {
    /**
     * @ignore
     */
    define('PHPEXCEL_ROOT', dirname(__FILE__) . '/../../');
    require(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
}

/**
 * PHPExcel_Reader_HTML
 *
 * Copyright (c) 2006 - 2015 PHPExcel
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
 * @package    PHPExcel_Reader
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
/** PHPExcel root directory */
class PHPExcel_Reader_HTML extends PHPExcel_Reader_HTML_Abstract
{

    /**
     * Input encoding
     *
     * @var string
     */
    protected $inputEncoding = 'ANSI';

    /**
     * Sheet index to read
     *
     * @var int
     */
    protected $sheetIndex = 0;

    /**
     * Formats
     *
     * @var array
     */
    protected $formats = array(
        'h1' => array(
            'font' => array(
                'bold' => true,
                'size' => 24,
            ),
        ), //    Bold, 24pt
        'h2' => array(
            'font' => array(
                'bold' => true,
                'size' => 18,
            ),
        ), //    Bold, 18pt
        'h3' => array(
            'font' => array(
                'bold' => true,
                'size' => 13.5,
            ),
        ), //    Bold, 13.5pt
        'h4' => array(
            'font' => array(
                'bold' => true,
                'size' => 12,
            ),
        ), //    Bold, 12pt
        'h5' => array(
            'font' => array(
                'bold' => true,
                'size' => 10,
            ),
        ), //    Bold, 10pt
        'h6' => array(
            'font' => array(
                'bold' => true,
                'size' => 7.5,
            ),
        ), //    Bold, 7.5pt
        'a' => array(
            'font' => array(
                'underline' => true,
                'color' => array(
                    'argb' => PHPExcel_Style_Color::COLOR_BLUE,
                ),
            ),
        ), //    Blue underlined
        'hr' => array(
            'borders' => array(
                'bottom' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array(
                        PHPExcel_Style_Color::COLOR_BLACK,
                    ),
                ),
            ),
        ), //    Bottom border
    );

    protected $rowspan = array();

    /**
     * Create a new PHPExcel_Reader_HTML
     */
    public function __construct()
    {
        $this->readFilter = new PHPExcel_Reader_DefaultReadFilter();
    }

    /**
     * Set input encoding
     *
     * @param string $pValue Input encoding
     */
    public function setInputEncoding($pValue = 'ANSI')
    {
        $this->inputEncoding = $pValue;

        return $this;
    }

    /**
     * Get input encoding
     *
     * @return string
     */
    public function getInputEncoding()
    {
        return $this->inputEncoding;
    }

    //    Data Array used for testing only, should write to PHPExcel object on completion of tests
    protected $dataArray = array();
    protected $tableLevel = 0;
    protected $nestedColumn = array('A');

    /**
     * Active Worksheet which is used for writing to.
     * @var \PHPExcel_Worksheet
     */
    protected $sheet;

    protected function setTableStartColumn($column)
    {
        if ($this->tableLevel == 0) {
            $column = 'A';
        }
        ++$this->tableLevel;
        $this->nestedColumn[$this->tableLevel] = $column;

        return $this->nestedColumn[$this->tableLevel];
    }

    protected function getTableStartColumn()
    {
        return $this->nestedColumn[$this->tableLevel];
    }

    protected function releaseTableStartColumn()
    {
        --$this->tableLevel;

        return array_pop($this->nestedColumn);
    }

    protected function loadHandler(\PHPExcel $objPHPExcel)
    {
        // Create new PHPExcel worksheets.
        while ($objPHPExcel->getSheetCount() <= $this->sheetIndex) {
            $objPHPExcel->createSheet();
        }
        $objPHPExcel->setActiveSheetIndex($this->sheetIndex);
        $this->sheet = $objPHPExcel->getActiveSheet();
    }

    protected function finishHandler()
    {
    }

    protected function flushCell($column, $row, &$cellContent)
    {
        if (is_string($cellContent)) {
            //    Simple String content
            if (trim($cellContent) > '') {
                //    Only actually write it if there's content in the string
//                echo 'FLUSH CELL: ' , $column , $row , ' => ' , $cellContent , '<br />';
                //    Write to worksheet to be done here...
                //    ... we return the cell so we can mess about with styles more easily
                $this->sheet->setCellValue($column . $row, $cellContent, true);
                $this->dataArray[$row][$column] = $cellContent;
            }
        } else {
            //    We have a Rich Text run
            //    TODO
            $this->dataArray[$row][$column] = 'RICH TEXT: ' . $cellContent;
        }
        $cellContent = (string) '';
    }

    protected function textElementHandler(\DOMNode $element, &$row, &$column, &$cellContent)
    {
        $domText = preg_replace('/\s+/u', ' ', trim($element->nodeValue));
        if (is_string($cellContent)) {
            //    simply append the text if the cell content is a plain text string
            $cellContent .= $domText;
        } else {
            //    but if we have a rich text run instead, we need to append it correctly
            //    TODO
        }
    }

    protected function defaultElementHandler(DOMNode $element, &$row, &$column, &$cellContent, $format = null)
    {
//                echo '<b>DOM ELEMENT: </b>' , strtoupper($element->nodeName) , '<br />';

                $attributeArray = array();
                foreach ($element->attributes as $attribute) {
//                    echo '<b>ATTRIBUTE: </b>' , $attribute->name , ' => ' , $attribute->value , '<br />';
                    $attributeArray[$attribute->name] = $attribute->value;
                }

                switch ($element->nodeName) {
                    case 'meta':
                        foreach ($attributeArray as $attributeName => $attributeValue) {
                            switch ($attributeName) {
                                case 'content':
                                    //    TODO
                                    //    Extract character set, so we can convert to UTF-8 if required
                                    break;
                            }
                        }
                        $this->processDomElement($element, $row, $column, $cellContent);
                        break;
                    case 'title':
                        $this->processDomElement($element, $row, $column, $cellContent);
                        $this->sheet->setTitle($cellContent);
                        $cellContent = '';
                        break;
                    case 'span':
                    case 'div':
                    case 'font':
                    case 'i':
                    case 'em':
                    case 'strong':
                    case 'b':
//                        echo 'STYLING, SPAN OR DIV<br />';
                        if ($cellContent > '') {
                            $cellContent .= ' ';
                        }
                        $this->processDomElement($element, $row, $column, $cellContent);
                        if ($cellContent > '') {
                            $cellContent .= ' ';
                        }
//                        echo 'END OF STYLING, SPAN OR DIV<br />';
                        break;
                    case 'hr':
                        $this->flushCell($column, $row, $cellContent);
                        ++$row;
                        if (isset($this->formats[$element->nodeName])) {
                            $this->sheet->getStyle($column . $row)->applyFromArray($this->formats[$element->nodeName]);
                        } else {
                            $cellContent = '----------';
                            $this->flushCell($column, $row, $cellContent);
                        }
                        ++$row;
                        // Add a break after a horizontal rule, simply by allowing the code to dropthru
                    case 'br':
                        if ($this->tableLevel > 0) {
                            //    If we're inside a table, replace with a \n
                            $cellContent .= "\n";
                        } else {
                            //    Otherwise flush our existing content and move the row cursor on
                            $this->flushCell($column, $row, $cellContent);
                            ++$row;
                        }
//                        echo 'HARD LINE BREAK: ' , '<br />';
                        break;
                    case 'a':
//                        echo 'START OF HYPERLINK: ' , '<br />';
                        foreach ($attributeArray as $attributeName => $attributeValue) {
                            switch ($attributeName) {
                                case 'href':
//                                    echo 'Link to ' , $attributeValue , '<br />';
                                    $this->sheet->getCell($column . $row)->getHyperlink()->setUrl($attributeValue);
                                    if (isset($this->formats[$element->nodeName])) {
                                        $this->sheet->getStyle($column . $row)->applyFromArray($this->formats[$element->nodeName]);
                                    }
                                    break;
                            }
                        }
                        $cellContent .= ' ';
                        $this->processDomElement($element, $row, $column, $cellContent);
//                        echo 'END OF HYPERLINK:' , '<br />';
                        break;
                    case 'h1':
                    case 'h2':
                    case 'h3':
                    case 'h4':
                    case 'h5':
                    case 'h6':
                    case 'ol':
                    case 'ul':
                    case 'p':
                        if ($this->tableLevel > 0) {
                            //    If we're inside a table, replace with a \n
                            $cellContent .= "\n";
//                            echo 'LIST ENTRY: ' , '<br />';
                            $this->processDomElement($element, $row, $column, $cellContent);
//                            echo 'END OF LIST ENTRY:' , '<br />';
                        } else {
                            if ($cellContent > '') {
                                $this->flushCell($column, $row, $cellContent);
                                $row++;
                            }
//                            echo 'START OF PARAGRAPH: ' , '<br />';
                            $this->processDomElement($element, $row, $column, $cellContent);
//                            echo 'END OF PARAGRAPH:' , '<br />';
                            $this->flushCell($column, $row, $cellContent);

                            if (isset($this->formats[$element->nodeName])) {
                                $this->sheet->getStyle($column . $row)->applyFromArray($this->formats[$element->nodeName]);
                            }

                            $row++;
                            $column = 'A';
                        }
                        break;
                    case 'li':
                        if ($this->tableLevel > 0) {
                            //    If we're inside a table, replace with a \n
                            $cellContent .= "\n";
//                            echo 'LIST ENTRY: ' , '<br />';
                            $this->processDomElement($element, $row, $column, $cellContent);
//                            echo 'END OF LIST ENTRY:' , '<br />';
                        } else {
                            if ($cellContent > '') {
                                $this->flushCell($column, $row, $cellContent);
                            }
                            ++$row;
//                            echo 'LIST ENTRY: ' , '<br />';
                            $this->processDomElement($element, $row, $column, $cellContent);
//                            echo 'END OF LIST ENTRY:' , '<br />';
                            $this->flushCell($column, $row, $cellContent);
                            $column = 'A';
                        }
                        break;
                    case 'table':
                        $this->flushCell($column, $row, $cellContent);
                        $column = $this->setTableStartColumn($column);
//                        echo 'START OF TABLE LEVEL ' , $this->tableLevel , '<br />';
                        if ($this->tableLevel > 1) {
                            --$row;
                        }
                        $this->processDomElement($element, $row, $column, $cellContent);
//                        echo 'END OF TABLE LEVEL ' , $this->tableLevel , '<br />';
                        $column = $this->releaseTableStartColumn();
                        if ($this->tableLevel > 1) {
                            ++$column;
                        } else {
                            ++$row;
                        }
                        break;
                    case 'thead':
                    case 'tbody':
                        $this->processDomElement($element, $row, $column, $cellContent);
                        break;
                    case 'tr':
                        $column = $this->getTableStartColumn();
                        $cellContent = '';
//                        echo 'START OF TABLE ' , $this->tableLevel , ' ROW<br />';
                        $this->processDomElement($element, $row, $column, $cellContent);
                        ++$row;
//                        echo 'END OF TABLE ' , $this->tableLevel , ' ROW<br />';
                        break;
                    case 'th':
                    case 'td':
//                        echo 'START OF TABLE ' , $this->tableLevel , ' CELL<br />';
                        $this->processDomElement($element, $row, $column, $cellContent);
//                        echo 'END OF TABLE ' , $this->tableLevel , ' CELL<br />';

                        while (isset($this->rowspan[$column . $row])) {
                            ++$column;
                        }

                        $this->flushCell($column, $row, $cellContent);

//                        if (isset($attributeArray['style']) && !empty($attributeArray['style'])) {
//                            $styleAry = $this->getPhpExcelStyleArray($attributeArray['style']);
//
//                            if (!empty($styleAry)) {
//                                $this->sheet->getStyle($column . $row)->applyFromArray($styleAry);
//                            }
//                        }

                        if (isset($attributeArray['rowspan']) && isset($attributeArray['colspan'])) {
                            //create merging rowspan and colspan
                            $columnTo = $column;
                            for ($i = 0; $i < $attributeArray['colspan'] - 1; $i++) {
                                ++$columnTo;
                            }
                            $range = $column . $row . ':' . $columnTo . ($row + $attributeArray['rowspan'] - 1);
                            foreach (\PHPExcel_Cell::extractAllCellReferencesInRange($range) as $value) {
                                $this->rowspan[$value] = true;
                            }
                            $this->sheet->mergeCells($range);
                            $column = $columnTo;
                        } elseif (isset($attributeArray['rowspan'])) {
                            //create merging rowspan
                            $range = $column . $row . ':' . $column . ($row + $attributeArray['rowspan'] - 1);
                            foreach (\PHPExcel_Cell::extractAllCellReferencesInRange($range) as $value) {
                                $this->rowspan[$value] = true;
                            }
                            $this->sheet->mergeCells($range);
                        } elseif (isset($attributeArray['colspan'])) {
                            //create merging colspan
                            $columnTo = $column;
                            for ($i = 0; $i < $attributeArray['colspan'] - 1; $i++) {
                                ++$columnTo;
                            }
                            $this->sheet->mergeCells($column . $row . ':' . $columnTo . $row);
                            $column = $columnTo;
                        }
                        ++$column;
                        break;
                    case 'body':
                        $row = 1;
                        $column = 'A';
                        $content = '';
                        $this->tableLevel = 0;
                        $this->processDomElement($element, $row, $column, $cellContent);
                        break;
                    default:
                        $this->processDomElement($element, $row, $column, $cellContent);
                }
                // This method does all traversing itself, no TRAVERSE_CHILD hint needed.
    }

    /**
     * Get sheet index
     *
     * @return int
     */
    public function getSheetIndex()
    {
        return $this->sheetIndex;
    }

    /**
     * Set sheet index
     *
     * @param  int                  $pValue Sheet index
     * @return PHPExcel_Reader_HTML
     */
    public function setSheetIndex($pValue = 0)
    {
        $this->sheetIndex = $pValue;

        return $this;
    }


}
