<?php
/**
 * PHPExcel
 *
 * Copyright (c) 2006 - 2014 PHPExcel
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
 * @package    PHPExcel_Worksheet
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPExcel_Worksheet_HeaderFooter
 *
 * <code>
 * Header/Footer Formatting Syntax taken from Office Open XML Part 4 - Markup Language Reference, page 1970:
 *
 * There are a number of formatting codes that can be written inline with the actual header / footer text, which
 * affect the formatting in the header or footer.
 *
 * Example: This example shows the text "Center Bold Header" on the first line (center section), and the date on
 * the second line (center section).
 * 		&CCenter &"-,Bold"Bold&"-,Regular"Header_x000A_&D
 *
 * General Rules:
 * There is no required order in which these codes must appear.
 *
 * The first occurrence of the following codes turns the formatting ON, the second occurrence turns it OFF again:
 * - strikethrough
 * - superscript
 * - subscript
 * Superscript and subscript cannot both be ON at same time. Whichever comes first wins and the other is ignored,
 * while the first is ON.
 * &L - code for "left section" (there are three header / footer locations, "left", "center", and "right"). When
 * two or more occurrences of this section marker exist, the contents from all markers are concatenated, in the
 * order of appearance, and placed into the left section.
 * &P - code for "current page #"
 * &N - code for "total pages"
 * &font size - code for "text font size", where font size is a font size in points.
 * &K - code for "text font color"
 * RGB Color is specified as RRGGBB
 * Theme Color is specifed as TTSNN where TT is the theme color Id, S is either "+" or "-" of the tint/shade
 * value, NN is the tint/shade value.
 * &S - code for "text strikethrough" on / off
 * &X - code for "text super script" on / off
 * &Y - code for "text subscript" on / off
 * &C - code for "center section". When two or more occurrences of this section marker exist, the contents
 * from all markers are concatenated, in the order of appearance, and placed into the center section.
 *
 * &D - code for "date"
 * &T - code for "time"
 * &G - code for "picture as background"
 * &U - code for "text single underline"
 * &E - code for "double underline"
 * &R - code for "right section". When two or more occurrences of this section marker exist, the contents
 * from all markers are concatenated, in the order of appearance, and placed into the right section.
 * &Z - code for "this workbook's file path"
 * &F - code for "this workbook's file name"
 * &A - code for "sheet tab name"
 * &+ - code for add to page #.
 * &- - code for subtract from page #.
 * &"font name,font type" - code for "text font name" and "text font type", where font name and font type
 * are strings specifying the name and type of the font, separated by a comma. When a hyphen appears in font
 * name, it means "none specified". Both of font name and font type can be localized values.
 * &"-,Bold" - code for "bold font style"
 * &B - also means "bold font style".
 * &"-,Regular" - code for "regular font style"
 * &"-,Italic" - code for "italic font style"
 * &I - also means "italic font style"
 * &"-,Bold Italic" code for "bold italic font style"
 * &O - code for "outline style"
 * &H - code for "shadow style"
 * </code>
 *
 * @category   PHPExcel
 * @package    PHPExcel_Worksheet
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Worksheet_HeaderFooter
{
	/* Header/footer image location */
	const IMAGE_HEADER_LEFT							= 'LH';
        const IMAGE_HEADER_LEFT_FIRST						= 'LHFIRST';
	const IMAGE_HEADER_CENTER						= 'CH';
        const IMAGE_HEADER_CENTER_FIRST						= 'CHFIRST';
	const IMAGE_HEADER_RIGHT						= 'RH';
        const IMAGE_HEADER_RIGHT_FIRST						= 'RHFIRST';
	const IMAGE_FOOTER_LEFT							= 'LF';
        const IMAGE_FOOTER_LEFT_FIRST						= 'LFFIRST';
	const IMAGE_FOOTER_CENTER						= 'CF';
        const IMAGE_FOOTER_CENTER_FIRST						= 'CFFIRST';
	const IMAGE_FOOTER_RIGHT						= 'RF';
        const IMAGE_FOOTER_RIGHT_FIRST						= 'RFFIRST';

	/**
	 * OddHeader
	 *
	 * @var string
	 */
	private $_oddHeader			= '';

	/**
	 * OddFooter
	 *
	 * @var string
	 */
	private $_oddFooter			= '';

	/**
	 * EvenHeader
	 *
	 * @var string
	 */
	private $_evenHeader		= '';

	/**
	 * EvenFooter
	 *
	 * @var string
	 */
	private $_evenFooter		= '';

	/**
	 * FirstHeader
	 *
	 * @var string
	 */
	private $_firstHeader		= '';

	/**
	 * FirstFooter
	 *
	 * @var string
	 */
	private $_firstFooter		= '';

	/**
	 * Different header for Odd/Even, defaults to false
	 *
	 * @var boolean
	 */
	private $_differentOddEven	= false;

	/**
	 * Different header for first page, defaults to false
	 *
	 * @var boolean
	 */
	private $_differentFirst	= false;

	/**
	 * Scale with document, defaults to true
	 *
	 * @var boolean
	 */
	private $_scaleWithDocument	= true;

	/**
	 * Align with margins, defaults to true
	 *
	 * @var boolean
	 */
	private $_alignWithMargins	= true;

	/**
	 * Header/footer images
	 *
	 * @var PHPExcel_Worksheet_HeaderFooterDrawing[]
	 */
	private $_headerFooterImages = array();

    /**
     * Create a new PHPExcel_Worksheet_HeaderFooter
     */
    public function __construct()
    {
    }

    /**
     * Get OddHeader
     *
     * @return string
     */
    public function getOddHeader() {
    	return $this->_oddHeader;
    }

    /**
     * Set OddHeader
     *
     * @param string $pValue
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function setOddHeader($pValue) {
    	$this->_oddHeader = $pValue;
    	return $this;
    }

    /**
     * Get OddFooter
     *
     * @return string
     */
    public function getOddFooter() {
    	return $this->_oddFooter;
    }

    /**
     * Set OddFooter
     *
     * @param string $pValue
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function setOddFooter($pValue) {
    	$this->_oddFooter = $pValue;
    	return $this;
    }

    /**
     * Get EvenHeader
     *
     * @return string
     */
    public function getEvenHeader() {
    	return $this->_evenHeader;
    }

    /**
     * Set EvenHeader
     *
     * @param string $pValue
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function setEvenHeader($pValue) {
    	$this->_evenHeader = $pValue;
    	return $this;
    }

    /**
     * Get EvenFooter
     *
     * @return string
     */
    public function getEvenFooter() {
    	return $this->_evenFooter;
    }

    /**
     * Set EvenFooter
     *
     * @param string $pValue
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function setEvenFooter($pValue) {
    	$this->_evenFooter = $pValue;
    	return $this;
    }

    /**
     * Get FirstHeader
     *
     * @return string
     */
    public function getFirstHeader() {
    	return $this->_firstHeader;
    }

    /**
     * Set FirstHeader
     *
     * @param string $pValue
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function setFirstHeader($pValue) {
    	$this->_firstHeader = $pValue;
    	return $this;
    }

    /**
     * Get FirstFooter
     *
     * @return string
     */
    public function getFirstFooter() {
    	return $this->_firstFooter;
    }

    /**
     * Set FirstFooter
     *
     * @param string $pValue
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function setFirstFooter($pValue) {
    	$this->_firstFooter = $pValue;
    	return $this;
    }

    /**
     * Get DifferentOddEven
     *
     * @return boolean
     */
    public function getDifferentOddEven() {
    	return $this->_differentOddEven;
    }

    /**
     * Set DifferentOddEven
     *
     * @param boolean $pValue
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function setDifferentOddEven($pValue = false) {
    	$this->_differentOddEven = $pValue;
    	return $this;
    }

    /**
     * Get DifferentFirst
     *
     * @return boolean
     */
    public function getDifferentFirst() {
    	return $this->_differentFirst;
    }

    /**
     * Set DifferentFirst
     *
     * @param boolean $pValue
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function setDifferentFirst($pValue = false) {
    	$this->_differentFirst = $pValue;
    	return $this;
    }

    /**
     * Get ScaleWithDocument
     *
     * @return boolean
     */
    public function getScaleWithDocument() {
    	return $this->_scaleWithDocument;
    }

    /**
     * Set ScaleWithDocument
     *
     * @param boolean $pValue
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function setScaleWithDocument($pValue = true) {
    	$this->_scaleWithDocument = $pValue;
    	return $this;
    }

    /**
     * Get AlignWithMargins
     *
     * @return boolean
     */
    public function getAlignWithMargins() {
    	return $this->_alignWithMargins;
    }

    /**
     * Set AlignWithMargins
     *
     * @param boolean $pValue
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function setAlignWithMargins($pValue = true) {
    	$this->_alignWithMargins = $pValue;
    	return $this;
    }

    /**
     * Add header/footer image
     *
     * @param PHPExcel_Worksheet_HeaderFooterDrawing $image
     * @param string $location
     * @throws PHPExcel_Exception
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function addImage(PHPExcel_Worksheet_HeaderFooterDrawing $image = null, $location = self::IMAGE_HEADER_LEFT) {
    	$this->_headerFooterImages[$location] = $image;
    	return $this;
    }

    /**
     * Remove header/footer image
     *
     * @param string $location
     * @throws PHPExcel_Exception
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function removeImage($location = self::IMAGE_HEADER_LEFT) {
    	if (isset($this->_headerFooterImages[$location])) {
    		unset($this->_headerFooterImages[$location]);
    	}
    	return $this;
    }

    /**
     * Set header/footer images
     *
     * @param PHPExcel_Worksheet_HeaderFooterDrawing[] $images
     * @throws PHPExcel_Exception
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function setImages($images) {
    	if (!is_array($images)) {
    		throw new PHPExcel_Exception('Invalid parameter!');
    	}

    	$this->_headerFooterImages = $images;
    	return $this;
    }

    /**
     * Get header/footer images
     *
     * @return PHPExcel_Worksheet_HeaderFooterDrawing[]
     */
    public function getImages() {
    	// Sort array
    	$images = array();
    	if (isset($this->_headerFooterImages[self::IMAGE_HEADER_LEFT])) 	$images[self::IMAGE_HEADER_LEFT] = 		$this->_headerFooterImages[self::IMAGE_HEADER_LEFT];
        if (isset($this->_headerFooterImages[self::IMAGE_HEADER_LEFT_FIRST])) 	$images[self::IMAGE_HEADER_LEFT_FIRST] = 	$this->_headerFooterImages[self::IMAGE_HEADER_LEFT_FIRST];
    	if (isset($this->_headerFooterImages[self::IMAGE_HEADER_CENTER])) 	$images[self::IMAGE_HEADER_CENTER] =            $this->_headerFooterImages[self::IMAGE_HEADER_CENTER];
        if (isset($this->_headerFooterImages[self::IMAGE_HEADER_CENTER_FIRST])) $images[self::IMAGE_HEADER_CENTER_FIRST] = 	$this->_headerFooterImages[self::IMAGE_HEADER_CENTER_FIRST];
    	if (isset($this->_headerFooterImages[self::IMAGE_HEADER_RIGHT])) 	$images[self::IMAGE_HEADER_RIGHT] =             $this->_headerFooterImages[self::IMAGE_HEADER_RIGHT];
        if (isset($this->_headerFooterImages[self::IMAGE_HEADER_RIGHT_FIRST])) 	$images[self::IMAGE_HEADER_RIGHT_FIRST] = 	$this->_headerFooterImages[self::IMAGE_HEADER_RIGHT_FIRST];
    	if (isset($this->_headerFooterImages[self::IMAGE_FOOTER_LEFT])) 	$images[self::IMAGE_FOOTER_LEFT] = 		$this->_headerFooterImages[self::IMAGE_FOOTER_LEFT];
        if (isset($this->_headerFooterImages[self::IMAGE_FOOTER_LEFT_FIRST])) 	$images[self::IMAGE_FOOTER_LEFT_FIRST] = 	$this->_headerFooterImages[self::IMAGE_FOOTER_LEFT_FIRST];
    	if (isset($this->_headerFooterImages[self::IMAGE_FOOTER_CENTER])) 	$images[self::IMAGE_FOOTER_CENTER] =            $this->_headerFooterImages[self::IMAGE_FOOTER_CENTER];
        if (isset($this->_headerFooterImages[self::IMAGE_FOOTER_CENTER_FIRST])) $images[self::IMAGE_FOOTER_CENTER_FIRST] = 	$this->_headerFooterImages[self::IMAGE_FOOTER_CENTER_FIRST];
    	if (isset($this->_headerFooterImages[self::IMAGE_FOOTER_RIGHT])) 	$images[self::IMAGE_FOOTER_RIGHT] =             $this->_headerFooterImages[self::IMAGE_FOOTER_RIGHT];
        if (isset($this->_headerFooterImages[self::IMAGE_FOOTER_RIGHT_FIRST])) 	$images[self::IMAGE_FOOTER_RIGHT_FIRST] = 	$this->_headerFooterImages[self::IMAGE_FOOTER_RIGHT_FIRST];
    	$this->_headerFooterImages = $images;

    	return $this->_headerFooterImages;
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
}
