<?php

namespace PhpOffice\PhpExcel;

/**
 * PHPExcel_RichText
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
 * @package    PHPExcel_RichText
 * @copyright  Copyright (c) 2006 - 2016 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class RichText implements IComparable
{
    /**
     * Rich text elements
     *
     * @var RichText\ITextElement[]
     */
    private $richTextElements;

    /**
     * Create a new RichText instance
     *
     * @param Cell $pCell
     * @throws Exception
     */
    public function __construct(Cell $pCell = null)
    {
        // Initialise variables
        $this->richTextElements = array();

        // Rich-Text string attached to cell?
        if ($pCell !== null) {
            // Add cell text and style
            if ($pCell->getValue() != "") {
                $objRun = new RichText\Run($pCell->getValue());
                $objRun->setFont(clone $pCell->getParent()->getStyle($pCell->getCoordinate())->getFont());
                $this->addText($objRun);
            }

            // Set parent value
            $pCell->setValueExplicit($this, Cell\DataType::TYPE_STRING);
        }
    }

    /**
     * Add text
     *
     * @param RichText\ITextElement $pText Rich text element
     * @throws Exception
     * @return RichText
     */
    public function addText(RichText\ITextElement $pText = null)
    {
        $this->richTextElements[] = $pText;
        return $this;
    }

    /**
     * Create text
     *
     * @param string $pText Text
     * @return RichText\TextElement
     * @throws Exception
     */
    public function createText($pText = '')
    {
        $objText = new RichText\TextElement($pText);
        $this->addText($objText);
        return $objText;
    }

    /**
     * Create text run
     *
     * @param string $pText Text
     * @return RichText\Run
     * @throws Exception
     */
    public function createTextRun($pText = '')
    {
        $objText = new RichText\Run($pText);
        $this->addText($objText);
        return $objText;
    }

    /**
     * Get plain text
     *
     * @return string
     */
    public function getPlainText()
    {
        // Return value
        $returnValue = '';

        // Loop through all RichText\ITextElements
        foreach ($this->richTextElements as $text) {
            $returnValue .= $text->getText();
        }

        return $returnValue;
    }

    /**
     * Convert to string
     *
     * @return string
     */
    public function __toString()
    {
        return $this->getPlainText();
    }

    /**
     * Get Rich Text elements
     *
     * @return RichText\ITextElement[]
     */
    public function getRichTextElements()
    {
        return $this->richTextElements;
    }

    /**
     * Set Rich Text elements
     *
     * @param RichText\ITextElement[] $pElements Array of elements
     * @throws Exception
     * @return RichText
     */
    public function setRichTextElements($pElements = null)
    {
        if (is_array($pElements)) {
            $this->richTextElements = $pElements;
        } else {
            throw new Exception("Invalid \PhpOffice\PhpExcel\RichText\ITextElement[] array passed.");
        }
        return $this;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        $hashElements = '';
        foreach ($this->richTextElements as $element) {
            $hashElements .= $element->getHashCode();
        }

        return md5(
            $hashElements .
            __CLASS__
        );
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
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
