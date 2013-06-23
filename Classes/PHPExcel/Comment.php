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
 * PHPExcel\Comment
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class Comment implements IComparable
{
    /**
     * Author
     *
     * @var string
     */
    protected $_author;

    /**
     * Rich text comment
     *
     * @var PHPExcel\RichText
     */
    protected $_text;

    /**
     * Comment width (CSS style, i.e. XXpx or YYpt)
     *
     * @var string
     */
    protected $_width = '96pt';

    /**
     * Left margin (CSS style, i.e. XXpx or YYpt)
     *
     * @var string
     */
    protected $_marginLeft = '59.25pt';

    /**
     * Top margin (CSS style, i.e. XXpx or YYpt)
     *
     * @var string
     */
    protected $_marginTop = '1.5pt';

    /**
     * Visible
     *
     * @var boolean
     */
    protected $_visible = false;

    /**
     * Comment height (CSS style, i.e. XXpx or YYpt)
     *
     * @var string
     */
    protected $_height = '55.5pt';

    /**
     * Comment fill color
     *
     * @var PHPExcel\Style_Color
     */
    protected $_fillColor;

    /**
     * Alignment
     *
     * @var string
     */
    protected $_alignment;

    /**
     * Create a new PHPExcel\Comment
     *
     * @throws PHPExcel\Exception
     */
    public function __construct()
    {
        // Initialise variables
        $this->_author        = 'Author';
        $this->_text        = new RichText();
        $this->_fillColor    = new Style_Color('FFFFFFE1');
        $this->_alignment    = Style_Alignment::HORIZONTAL_GENERAL;
    }

    /**
     * Get Author
     *
     * @return string
     */
    public function getAuthor() {
        return $this->_author;
    }

    /**
     * Set Author
     *
     * @param string $pValue
     * @return PHPExcel\Comment
     */
    public function setAuthor($pValue = '') {
        $this->_author = $pValue;
        return $this;
    }

    /**
     * Get Rich text comment
     *
     * @return PHPExcel\RichText
     */
    public function getText() {
        return $this->_text;
    }

    /**
     * Set Rich text comment
     *
     * @param PHPExcel\RichText $pValue
     * @return PHPExcel\Comment
     */
    public function setText(RichText $pValue) {
        $this->_text = $pValue;
        return $this;
    }

    /**
     * Get comment width (CSS style, i.e. XXpx or YYpt)
     *
     * @return string
     */
    public function getWidth() {
        return $this->_width;
    }

    /**
     * Set comment width (CSS style, i.e. XXpx or YYpt)
     *
     * @param string $value
     * @return PHPExcel\Comment
     */
    public function setWidth($value = '96pt') {
        $this->_width = $value;
        return $this;
    }

    /**
     * Get comment height (CSS style, i.e. XXpx or YYpt)
     *
     * @return string
     */
    public function getHeight() {
        return $this->_height;
    }

    /**
     * Set comment height (CSS style, i.e. XXpx or YYpt)
     *
     * @param string $value
     * @return PHPExcel\Comment
     */
    public function setHeight($value = '55.5pt') {
        $this->_height = $value;
        return $this;
    }

    /**
     * Get left margin (CSS style, i.e. XXpx or YYpt)
     *
     * @return string
     */
    public function getMarginLeft() {
        return $this->_marginLeft;
    }

    /**
     * Set left margin (CSS style, i.e. XXpx or YYpt)
     *
     * @param string $value
     * @return PHPExcel\Comment
     */
    public function setMarginLeft($value = '59.25pt') {
        $this->_marginLeft = $value;
        return $this;
    }

    /**
     * Get top margin (CSS style, i.e. XXpx or YYpt)
     *
     * @return string
     */
    public function getMarginTop() {
        return $this->_marginTop;
    }

    /**
     * Set top margin (CSS style, i.e. XXpx or YYpt)
     *
     * @param string $value
     * @return PHPExcel\Comment
     */
    public function setMarginTop($value = '1.5pt') {
        $this->_marginTop = $value;
        return $this;
    }

    /**
     * Is the comment visible by default?
     *
     * @return boolean
     */
    public function getVisible() {
        return $this->_visible;
    }

    /**
     * Set comment default visibility
     *
     * @param boolean $value
     * @return PHPExcel\Comment
     */
    public function setVisible($value = false) {
        $this->_visible = $value;
        return $this;
    }

    /**
     * Get fill color
     *
     * @return PHPExcel\Style_Color
     */
    public function getFillColor() {
        return $this->_fillColor;
    }

    /**
     * Set Alignment
     *
     * @param string $pValue
     * @return PHPExcel\Comment
     */
    public function setAlignment($pValue = Style_Alignment::HORIZONTAL_GENERAL) {
        $this->_alignment = $pValue;
        return $this;
    }

    /**
     * Get Alignment
     *
     * @return string
     */
    public function getAlignment() {
        return $this->_alignment;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode() {
        return md5(
              $this->_author
            . $this->_text->getHashCode()
            . $this->_width
            . $this->_height
            . $this->_marginLeft
            . $this->_marginTop
            . ($this->_visible ? 1 : 0)
            . $this->_fillColor->getHashCode()
            . $this->_alignment
            . __CLASS__
        );
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

    /**
     * Convert to string
     *
     * @return string
     */
    public function __toString() {
        return $this->_text->getPlainText();
    }
}
