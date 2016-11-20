<?php

class PHPExcel_Adapter_Spreadsheet_Excel_Writer_Format extends PEAR
{
	private $_sewWorkbook;
	private $_peWorkbook;
	private $_peStyle;

	/**
	 * Constructor
	 *
	 * @access private
	 * @param PHPExcel_Adapter_Spreadsheet_Excel_Writer	$workbook	The parent workbook
	 * @param array	  $properties array with properties to be set on initialization.
	 */
	function PHPExcel_Adapter_Spreadsheet_Excel_Writer_Format($workbook, $properties = array())
	{
		$this->_sewWorkbook = $workbook;
		$this->_peWorkbook = $workbook->_getPhpExcelWorkbook();
		$this->_peStyle = clone $this->_peWorkbook->getDefaultStyle();
		$this->_peWorkbook->addCellXf($this->_peStyle);

		foreach ($properties as $property => $value)
		{
			if (method_exists($this, 'set' . ucwords($property)))
			{
				$method_name = 'set' . ucwords($property);
				$this->$method_name($value);
			}
		}
	}

	function _getPhpExcelStyle()
	{
		return $this->_peStyle;
	}

	/**
	 * Set cell alignment.
	 *
	 * @access public
	 * @param string $location alignment for the cell ('left', 'right', etc...).
	 */
	function setAlign($location)
	{
		$this->setHAlign($location);
		$this->setVAlign($location);
	}

	/**
	 * Set cell horizontal alignment.
	 *
	 * @access public
	 * @param string $location alignment for the cell ('left', 'right', etc...).
	 */
	function setHAlign($location)
	{
		$location = strtolower($location);

		$halign = null;

		if ($location == 'left') {
			$halign = PHPExcel_Style_Alignment::HORIZONTAL_LEFT;

		} elseif ($location == 'centre' || $location == 'center') {
			$halign = PHPExcel_Style_Alignment::HORIZONTAL_CENTER;

		} elseif ($location == 'right') {
			$halign = PHPExcel_Style_Alignment::HORIZONTAL_RIGHT;

		} elseif ($location == 'fill') {
			$halign = PHPExcel_Style_Alignment::HORIZONTAL_CENTER_CONTINUOUS;

		} elseif ($location == 'justify') {
			$halign = PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY;

		}

		if ($halign != null) {
			$this->_peStyle->getAlignment()->setHorizontal($halign);
		}
	}

	/**
	 * Set cell vertical alignment.
	 *
	 * @access public
	 * @param string $location alignment for the cell ('top', 'vleft', 'vright', etc...).
	 */
	function setVAlign($location)
	{
		$location = strtolower($location);

		$valign = null;

		if ($location == 'top') {
			$valign = PHPExcel_Style_Alignment::VERTICAL_TOP;

		} elseif ($location == 'vcentre' || $location == 'vcenter') {
			$valign = PHPExcel_Style_Alignment::VERTICAL_CENTER;

		} elseif ($location == 'bottom') {
			$valign = PHPExcel_Style_Alignment::VERTICAL_BOTTOM;

		} elseif ($location == 'vjustelseify') {
			$valign = PHPExcel_Style_Alignment::VERTICAL_JUSTIFY;
		}

		if ($valign != null) {
			$this->_peStyle->getAlignment()->setVertical($valign);
		}
	}

	/**
	 * This is an alias for the unintuitive setAlign('merge')
	 *
	 * @access public
	 */
	function setMerge()
	{
		return $this->raiseError('Not implemented');
	}

	/**
	 * Sets the boldness of the text.
	 *
	 * @access public
	 * @param integer $weight Weight for the text, 0 is normal text, 1 is bold.
	 */
	function setBold($weight = 1)
	{
		if ($weight == 1) {
			$this->_peStyle->getFont()->setBold(true);
		} else if ($weight == 0) {
			$this->_peStyle->getFont()->setBold(false);
		}
	}

	function _convertBorderStyle($style)
	{
		if ($style == 1) {
			return PHPExcel_Style_Border::BORDER_THIN;
		} elseif ($style == 2) {
			return PHPExcel_Style_Border::BORDER_THICK;
		}
	}

	/**
	 * Sets the width for the bottom border of the cell
	 *
	 * @access public
	 * @param integer $style style of the cell border. 1 => thin, 2 => thick.
	 */
	function setBottom($style)
	{
		$this->_peStyle->getBorders()->getBottom()->setBorderStyle($this->_convertBorderStyle($style));
	}

	/**
	 * Sets the width for the top border of the cell
	 *
	 * @access public
	 * @param integer $style style of the cell top border. 1 => thin, 2 => thick.
	 */
	function setTop($style)
	{
		$this->_peStyle->getBorders()->getTop()->setBorderStyle($this->_convertBorderStyle($style));
	}

	/**
	 * Sets the width for the left border of the cell
	 *
	 * @access public
	 * @param integer $style style of the cell left border. 1 => thin, 2 => thick.
	 */
	function setLeft($style)
	{
		$this->_peStyle->getBorders()->getLeft()->setBorderStyle($this->_convertBorderStyle($style));
	}

	/**
	 * Sets the width for the right border of the cell
	 *
	 * @access public
	 * @param integer $style style of the cell right border. 1 => thin, 2 => thick.
	 */
	function setRight($style)
	{
		$this->_peStyle->getBorders()->getRight()->setBorderStyle($this->_convertBorderStyle($style));
	}

	/**
	 * Set cells borders to the same style
	 *
	 * @access public
	 * @param integer $style style to apply for all cell borders. 1 => thin, 2 => thick.
	 */
	function setBorder($style)
	{
		$this->setBottom($style);
		$this->setTop($style);
		$this->setLeft($style);
		$this->setRight($style);
	}

	private $_colors = array(
		'aqua'	  => 'FF00FFFF',
		'cyan'	  => 'FF00FFFF',
		'black'	  => 'FF000000',
		'blue'	  => 'FF0000FF',
		'brown'	  => 'FFA52A2A',
		'magenta' => 'FFFF00FF',
		'fuchsia' => 'FFFF00FF',
		'gray'	  => 'FF808080',
		'grey'	  => 'FF808080',
		'green'	  => 'FF008000',
		'lime'	  => 'FF00FF00',
		'navy'	  => 'FF000080',
		'orange'  => 'FFFFA500',
		'purple'  => 'FF800080',
		'red'	  => 'FFFF0000',
		'silver'  => 'FFC0C0C0',
		'white'	  => 'FFFFFFFF',
		'yellow'  => 'FFFFFF00',
		);

	function _cvtColor($color)
	{
		if(isset($this->_colors[$color])) {
			return $this->_colors[$color];
		}

		return $this->_sewWorkbook->_getCustomColor($color);
	}

	/**
	 * Sets all the cell's borders to the same color
	 *
	 * @access public
	 * @param mixed $color The color we are setting. Either a string (like 'blue'),
	 *					   or an integer (range is [8...63]).
	 */
	function setBorderColor($color)
	{
		$this->setBottomColor($color);
		$this->setTopColor($color);
		$this->setLeftColor($color);
		$this->setRightColor($color);
	}

	/**
	 * Sets the cell's bottom border color
	 *
	 * @access public
	 * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
	 */
	function setBottomColor($color)
	{
		$this->_peStyle->getBorders()->getBottom()->setColor($this->_cvtColor($color));
	}

	/**
	 * Sets the cell's top border color
	 *
	 * @access public
	 * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
	 */
	function setTopColor($color)
	{
		$this->_peStyle->getBorders()->getTop()->setColor($this->_cvtColor($color));
	}

	/**
	 * Sets the cell's left border color
	 *
	 * @access public
	 * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
	 */
	function setLeftColor($color)
	{
		$this->_peStyle->getBorders()->getLeft()->setColor($this->_cvtColor($color));
	}

	/**
	 * Sets the cell's right border color
	 *
	 * @access public
	 * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
	 */
	function setRightColor($color)
	{
		$this->_peStyle->getBorders()->getRight()->setColor($this->_cvtColor($color));
	}

	/**
	 * Sets the cell's foreground color
	 *
	 * @access public
	 * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
	 */
	function setFgColor($color)
	{
		$this->setBgColor($color);
	}

	/**
	 * Sets the cell's background color
	 *
	 * @access public
	 * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
	 */
	function setBgColor($color)
	{
		$this->_peStyle->getFill()->
			setFillType(PHPExcel_Style_Fill::FILL_SOLID)->
			getStartColor()->
			setARGB($this->_cvtColor($color));
	}

	/**
	 * Sets the cell's color
	 *
	 * @access public
	 * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
	 */
	function setColor($color)
	{
		$this->_peStyle->getFont()->getColor()->setARGB($this->_cvtColor($color));
	}

	/**
	 * Sets the fill pattern attribute of a cell
	 *
	 * @access public
	 * @param integer $arg Optional. Defaults to 1. Meaningful values are: 0-18,
	 *					   0 meaning no background.
	 */
	function setPattern($arg = 1)
	{
		return $this->raiseError('Not implemented');
	}

	/**
	 * Sets the underline of the text
	 *
	 * @access public
	 * @param integer $underline The value for underline. Possible values are:
	 *							1 => underline, 2 => double underline.
	 */
	function setUnderline($underline)
	{
		if ($underline == 1) {
			$this->_peStyle->getFont()->setUnderline(PHPExcel_Style_Font::UNDERLINE_SINGLE);
		} else if ($underline == 2) {
			$this->_peStyle->getFont()->setUnderline(PHPExcel_Style_Font::UNDERLINE_DOUBLE);
		}
	}

	/**
	 * Sets the font style as italic
	 *
	 * @access public
	 */
	function setItalic()
	{
		$this->_peStyle->getFont()->setItalic(true);
	}

	/**
	 * Sets the font size
	 *
	 * @access public
	 * @param integer $size The font size.
	 */
	function setSize($size)
	{
		$this->_peStyle->getFont()->setSize($size);
	}

	/**
	 * Sets text wrapping
	 *
	 * @access public
	 */
	function setTextWrap()
	{
		$this->_peStyle->getAlignment()->setWrapText(true);
	}

	/**
	 * Sets the orientation of the text
	 *
	 * @access public
	 * @param integer $angle The rotation angle for the text (clockwise). Possible
	 values are: 0, 90, 270 and -1 for stacking top-to-bottom.
	*/
	function setTextRotation($angle)
	{
		$this->_peStyle->getAlignment()->setTextRotation($angle);
	}

	/**
	 * Sets the numeric format.
	 * It can be date, time, currency, etc...
	 *
	 * @access public
	 * @param integer $num_format The numeric format.
	 */
	function setNumFormat($num_format)
	{
		$this->_peStyle->getNumberFormat()->setFormatCode($num_format);
	}

	/**
	 * Sets font as strikeout.
	 *
	 * @access public
	 */
	function setStrikeOut()
	{
		$this->_peStyle->getFont()->setStrikethrough(true);
	}

	/**
	 * Sets outlining for a font.
	 *
	 * @access public
	 */
	function setOutLine()
	{
		return $this->raiseError('Not implemented');
	}

	/**
	 * Sets font as shadow.
	 *
	 * @access public
	 */
	function setShadow()
	{
		return $this->raiseError('Not implemented');
	}

	/**
	 * Sets the script type of the text
	 *
	 * @access public
	 * @param integer $script The value for script type. Possible values are:
	 *						  1 => superscript, 2 => subscript.
	 */
	function setScript($script)
	{
		if ($script == 1) {
			$this->_peStyle->getFont()->setSuperScript(true);
		} else if ($script == 2) {
			$this->_peStyle->getFont()->setSubScript(true);
		}
	}

	/**
	 * Locks a cell.
	 *
	 * @access public
	 */
	function setLocked()
	{
		return $this->raiseError('Not implemented');
	}

	/**
	 * Unlocks a cell. Useful for unprotecting particular cells of a protected sheet.
	 *
	 * @access public
	 */
	function setUnLocked()
	{
		return $this->raiseError('Not implemented');
	}

	/**
	 * Sets the font family name.
	 *
	 * @access public
	 * @param string $fontfamily The font family name. Possible values are:
	 *							 'Times New Roman', 'Arial', 'Courier'.
	 */
	function setFontFamily($font_family)
	{
		$this->_peStyle->getFont()->setName($font_family);
	}
}
