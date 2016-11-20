<?php

class PHPExcel_Adapter_Spreadsheet_Excel_Writer_Worksheet extends PEAR
{
	private $_sewWorkbook;
	private $_peWorkbook;
	private $_peWorksheet;
	private $_cacheController;
	private $_checkNotify;

	/**
	 * Constructor
	 *
	 * @access private
	 * @param PHPExcel_Adapter_Spreadsheet_Excel_Writer	$workbook	The parent workbook
	 * @param string	$name	The name of the new worksheet
	 * @param boolean	$firstSheet	First sheet flag
	 */
	function PHPExcel_Adapter_Spreadsheet_Excel_Writer_Worksheet($workbook, $name, $firstSheet)
	{
		$this->_sewWorkbook = $workbook;
		$this->_peWorkbook = $workbook->_getPhpExcelWorkbook();

		if ($firstSheet) {
			$this->_peWorksheet = $this->_peWorkbook->getActiveSheet();
		} else {
			$this->_peWorksheet = $this->_peWorkbook->createSheet();
		}

		if (method_exists($this->_peWorksheet, 'setNotifyCacheControllerEnabled')) {
			$this->_peWorksheet->setNotifyCacheControllerEnabled(false);
		}

		if ($name != null && $name != '') {
			$tname = $name;
			// trim long worksheet title
			if (PHPExcel_Shared_String::CountCharacters($name) > 31) {
				$tname = PHPExcel_Shared_String::Substring($tname, 0, 28) . '...';
			}
			$this->_peWorksheet->setTitle($tname);
		}

		$this->_cacheController = $this->_peWorksheet->getCellCacheController();
		$this->_checkNotify = method_exists($this->_peWorksheet, 'setNotifyCacheControllerEnabled');
	}

	function _getPhpExcelWorksheet()
	{
		return $this->_peWorksheet;
	}

	function close($sheetnames)
	{
		// skip
	}

	/**
	 * Retrieve the worksheet name.
	 * This is usefull when creating worksheets without a name.
	 *
	 * @access public
	 * @return string The worksheet's name
	 */
	function getName()
	{
		return $this->_peWorksheet->getTitle();
	}

	/**
	 * Sets a merged cell range
	 *
	 * @access public
	 * @param integer $first_row First row of the area to merge
	 * @param integer $first_col First column of the area to merge
	 * @param integer $last_row	 Last row of the area to merge
	 * @param integer $last_col	 Last column of the area to merge
	 */
	function setMerge($first_row, $first_col, $last_row, $last_col)
	{
		$this->mergeCells($first_row, $first_col, $last_row, $last_col);
	}

	/**
	 * Set this worksheet as a selected worksheet,
	 * i.e. the worksheet has its tab highlighted.
	 *
	 * @access public
	 */
	function select()
	{
		$this->activate();
	}

	/**
	 * Set this worksheet as the active worksheet,
	 * i.e. the worksheet that is displayed when the workbook is opened.
	 * Also set it as selected.
	 *
	 * @access public
	 */
	function activate()
	{
		$index = $this->_peWorkbook->getIndex($this->_peWorksheet);
		$this->_peWorkbook->setActiveSheetIndex($index);
	}

	/**
	 * Set this worksheet as the first visible sheet.
	 * This is necessary when there are a large number of worksheets and the
	 * activated worksheet is not visible on the screen.
	 *
	 * @access public
	 */
	function setFirstSheet()
	{
		return $this->raiseError('Not implemented');
	}

	/**
	 * Set the worksheet protection flag
	 * to prevent accidental modification and to
	 * hide formulas if the locked and hidden format properties have been set.
	 *
	 * @access public
	 * @param string $password The password to use for protecting the sheet.
	 */
	function protect($password)
	{
		return $this->raiseError('Not implemented');
	}

	/**
	 * Set the width of a single column or a range of columns.
	 *
	 * @access public
	 * @param integer $firstcol first column on the range
	 * @param integer $lastcol	last column on the range
	 * @param integer $width	width to set
	 * @param mixed	  $format	The optional XF format to apply to the columns
	 * @param integer $hidden	The optional hidden atribute
	 * @param integer $level	The optional outline level
	 */
	function setColumn($firstcol, $lastcol, $width, $format = 0, $hidden = 0, $level = 0)
	{
		for ($col = $firstcol; $col <= $lastcol; $col++) {
			$colDimension = $this->_peWorksheet->getColumnDimensionByColumn($col);
			if ($width > 0) {
				$colDimension->setWidth($width * (7 + 5.0 / $width) / 7.0);
			} else {
				$colDimension->setWidth(0);
			}
			$colDimension->setVisible(!$hidden);
			$colDimension->setOutlineLevel($level);
		}
	}

	/**
	 * Set which cell or cells are selected in a worksheet
	 *
	 * @access public
	 * @param integer $first_row	first row in the selected quadrant
	 * @param integer $first_column first column in the selected quadrant
	 * @param integer $last_row		last row in the selected quadrant
	 * @param integer $last_column	last column in the selected quadrant
	 */
	function setSelection($first_row,$first_column,$last_row,$last_column)
	{
		return $this->raiseError('Not implemented');
	}

	/**
	 * Set panes and mark them as frozen.
	 *
	 * @access public
	 * @param array $panes This is the only parameter received and is composed of the following:
	 *					   0 => Vertical split position,
	 *					   1 => Horizontal split position
	 *					   2 => Top row visible
	 *					   3 => Leftmost column visible
	 *					   4 => Active pane
	 */
	function freezePanes($panes)
	{
		$this->_peWorksheet->freezePaneByColumnAndRow($panes[0], $panes[1] + 1);
	}

	/**
	 * Set panes and mark them as unfrozen.
	 *
	 * @access public
	 * @param array $panes This is the only parameter received and is composed of the following:
	 *					   0 => Vertical split position,
	 *					   1 => Horizontal split position
	 *					   2 => Top row visible
	 *					   3 => Leftmost column visible
	 *					   4 => Active pane
	 */
	function thawPanes($panes)
	{
		$this->_peWorksheet->unfreezePane();
	}

	/**
	 * Set the page orientation as portrait.
	 *
	 * @access public
	 */
	function setPortrait()
	{
		$pageSetup = $this->_peWorksheet->getPageSetup();
		$pageSetup->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
	}

	/**
	 * Set the page orientation as landscape.
	 *
	 * @access public
	 */
	function setLandscape()
	{
		$pageSetup = $this->_peWorksheet->getPageSetup();
		$pageSetup->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
	}

	/**
	 * Set the paper type. Ex. 1 = US Letter, 9 = A4
	 *
	 * @access public
	 * @param integer $size The type of paper size to use
	 */
	function setPaper($size = 0)
	{
		$pageSetup = $this->_peWorksheet->getPageSetup();
		$pageSetup->setPaperSize($size);
	}

	/**
	 * Set the page header caption and optional margin.
	 *
	 * @access public
	 * @param string $string The header text
	 * @param float	 $margin optional head margin in inches.
	 */
	function setHeader($string, $margin = 0.5)
	{
		$headerFooter = $this->_peWorksheet->getHeaderFooter();
		$headerFooter->setOddHeader($string);
		$this->_peWorksheet->getPageMargins()->setHeader($margin);
	}

	/**
	 * Set the page footer caption and optional margin.
	 *
	 * @access public
	 * @param string $string The footer text
	 * @param float	 $margin optional foot margin in inches.
	 */
	function setFooter($string, $margin = 0.5)
	{
		$headerFooter = $this->_peWorksheet->getHeaderFooter();
		$headerFooter->setOddFooter($string);
		$this->_peWorksheet->getPageMargins()->setFooter($margin);
	}

	/**
	 * Center the page horinzontally.
	 *
	 * @access public
	 * @param integer $center the optional value for centering. Defaults to 1 (center).
	 */
	function centerHorizontally($center = 1)
	{
		$pageSetup = $this->_peWorksheet->getPageSetup();
		$pageSetup->setHorizontalCentered(($center > 0) ? true : false);
	}

	/**
	 * Center the page vertically.
	 *
	 * @access public
	 * @param integer $center the optional value for centering. Defaults to 1 (center).
	 */
	function centerVertically($center = 1)
	{
		$pageSetup = $this->_peWorksheet->getPageSetup();
		$pageSetup->setVerticalCentered(($center > 0) ? true : false);
	}

	/**
	 * Set all the page margins to the same value in inches.
	 *
	 * @access public
	 * @param float $margin The margin to set in inches
	 */
	function setMargins($margin)
	{
		$this->setMarginLeft($margin);
		$this->setMarginRight($margin);
		$this->setMarginTop($margin);
		$this->setMarginBottom($margin);
	}

	/**
	 * Set the left and right margins to the same value in inches.
	 *
	 * @access public
	 * @param float $margin The margin to set in inches
	 */
	function setMargins_LR($margin)
	{
		$this->setMarginLeft($margin);
		$this->setMarginRight($margin);
	}

	/**
	 * Set the top and bottom margins to the same value in inches.
	 *
	 * @access public
	 * @param float $margin The margin to set in inches
	 */
	function setMargins_TB($margin)
	{
		$this->setMarginTop($margin);
		$this->setMarginBottom($margin);
	}

	/**
	 * Set the left margin in inches.
	 *
	 * @access public
	 * @param float $margin The margin to set in inches
	 */
	function setMarginLeft($margin = 0.75)
	{
		$this->_peWorksheet->getPageMargins()->setLeft($margin);
	}

	/**
	 * Set the right margin in inches.
	 *
	 * @access public
	 * @param float $margin The margin to set in inches
	 */
	function setMarginRight($margin = 0.75)
	{
		$this->_peWorksheet->getPageMargins()->setRight($margin);
	}

	/**
	 * Set the top margin in inches.
	 *
	 * @access public
	 * @param float $margin The margin to set in inches
	 */
	function setMarginTop($margin = 1.00)
	{
		$this->_peWorksheet->getPageMargins()->setTop($margin);
	}

	/**
	 * Set the bottom margin in inches.
	 *
	 * @access public
	 * @param float $margin The margin to set in inches
	 */
	function setMarginBottom($margin = 1.00)
	{
		$this->_peWorksheet->getPageMargins()->setBottom($margin);
	}

	/**
	 * Set the rows to repeat at the top of each printed page.
	 *
	 * @access public
	 * @param integer $first_row First row to repeat
	 * @param integer $last_row	 Last row to repeat. Optional.
	 */
	function repeatRows($first_row, $last_row = NULL)
	{
		if (!isset($last_row)) {
			$last_row  = $first_row;
		}

		$pageSetup = $this->_peWorksheet->getPageSetup();
		$pageSetup->setRowsToRepeatAtTopByStartAndEnd($first_row + 1, $last_row + 1);
	}

	/**
	 * Set the columns to repeat at the left hand side of each printed page.
	 *
	 * @access public
	 * @param integer $first_col First column to repeat
	 * @param integer $last_col	 Last column to repeat. Optional.
	 */
	function repeatColumns($first_col, $last_col = NULL)
	{
		if (!isset($last_col)) {
			$last_col  = $first_col;
		}

		$pageSetup = $this->_peWorksheet->getPageSetup();
		$pageSetup->setColumnsToRepeatAtLeftByStartAndEnd(
			PHPExcel_Cell::stringFromColumnIndex($first_col),
			PHPExcel_Cell::stringFromColumnIndex($last_col)
			);
	}

	/**
	 * Set the area of each worksheet that will be printed.
	 *
	 * @access public
	 * @param integer $first_row First row of the area to print
	 * @param integer $first_col First column of the area to print
	 * @param integer $last_row	 Last row of the area to print
	 * @param integer $last_col	 Last column of the area to print
	 */
	function printArea($first_row, $first_col, $last_row, $last_col)
	{
		$pageSetup = $this->_peWorksheet->getPageSetup();
		$pageSetup->setPrintAreaByColumnAndRow($first_col, $first_row + 1, $last_col, $last_row + 1);
	}

	/**
	 * Set the option to hide gridlines on the printed page.
	 *
	 * @access public
	 */
	function hideGridlines()
	{
		$this->_peWorksheet->setPrintGridlines(false);
	}

	/**
	 * Set the option to hide gridlines on the worksheet (as seen on the screen).
	 *
	 * @access public
	 */
	function hideScreenGridlines()
	{
		$this->_peWorksheet->setShowGridlines(false);
	}

	/**
	 * Set the option to print the row and column headers on the printed page.
	 *
	 * @access public
	 * @param integer $print Whether to print the headers or not. Defaults to 1 (print).
	 */
	function printRowColHeaders($print = 1)
	{
		return $this->raiseError('Not implemented');
	}

	/**
	 * Set the vertical and horizontal number of pages that will define the maximum area printed.
	 * It doesn't seem to work with OpenOffice.
	 *
	 * @access public
	 * @param  integer $width  Maximun width of printed area in pages
	 * @param  integer $height Maximun heigth of printed area in pages
	 * @see setPrintScale()
	 */
	function fitToPages($width, $height)
	{
		$pageSetup = $this->_peWorksheet->getPageSetup();
		$pageSetup->setFitToWidth($width);
		$pageSetup->setFitToHeight($height);
	}

	/**
	 * Store the horizontal page breaks on a worksheet (for printing).
	 * The breaks represent the row after which the break is inserted.
	 *
	 * @access public
	 * @param array $breaks Array containing the horizontal page breaks
	 */
	function setHPagebreaks($breaks)
	{
		foreach($breaks as $row) {
			$this->_phpExcelWorkshee->setBreakByColumnAndRow(0, $row + 1, PHPExcel_Worksheet::BREAK_ROW);
		}
	}

	/**
	 * Store the vertical page breaks on a worksheet (for printing).
	 * The breaks represent the column after which the break is inserted.
	 *
	 * @access public
	 * @param array $breaks Array containing the vertical page breaks
	 */
	function setVPagebreaks($breaks)
	{
		foreach($breaks as $col) {
			$this->_phpExcelWorkshee->setBreakByColumnAndRow($col, 1, PHPExcel_Worksheet::BREAK_COLUMN);
		}
	}

	/**
	 * Set the worksheet zoom factor.
	 *
	 * @access public
	 * @param integer $scale The zoom factor
	 */
	function setZoom($scale = 100)
	{
		$this->_peWorksheet->getSheetView()->setZoomScale($scale);
	}

	/**
	 * Set the scale factor for the printed page.
	 * It turns off the "fit to page" option
	 *
	 * @access public
	 * @param integer $scale The optional scale factor. Defaults to 100
	 */
	function setPrintScale($scale = 100)
	{
		$pageSetup = $this->_peWorksheet->getPageSetup();
		$pageSetup->setScale($scale);
	}

	/**
	 * Map to the appropriate write method acording to the token recieved.
	 *
	 * @access public
	 * @param integer $row	  The row of the cell we are writing to
	 * @param integer $col	  The column of the cell we are writing to
	 * @param mixed	  $token  What we are writing
	 * @param mixed	  $format The optional format to apply to the cell
	 */
	function write($row, $col, $str, $format = null)
	{
		$dataType = null;

		if ($str === null) {
			$dataType = PHPExcel_Cell_DataType::TYPE_NULL;

		} elseif ($str === '') {
			$dataType = PHPExcel_Cell_DataType::TYPE_STRING;

		} elseif ($str{0} === '0' && strlen($str) > 1 && $str{1} !== '.') {
			$dataType = PHPExcel_Cell_DataType::TYPE_STRING;

		} elseif (is_numeric($str)) {
			$dataType = PHPExcel_Cell_DataType::TYPE_NUMERIC;

		} elseif ($str{0} === '=' && strlen($str) > 1) {
			$dataType = PHPExcel_Cell_DataType::TYPE_FORMULA;

		} elseif (is_string($str)) {
			$dataType = PHPExcel_Cell_DataType::TYPE_STRING;

		} elseif (is_bool($str)) {
			$dataType = PHPExcel_Cell_DataType::TYPE_BOOL;

		} elseif ($str instanceof PHPExcel_RichText) {
			$dataType = PHPExcel_Cell_DataType::TYPE_STRING;

		} else {
			$dataType = PHPExcel_Cell_DataType::TYPE_STRING;

		}

		$cell = $this->_getPhpExcelCell($row, $col, $format);
		$cell->setValueExplicit($str, $dataType);
		$this->_releasePhpExcelCell($cell);
	}

	/**
	 * Write a double to the specified row and column (zero indexed).
	 * An integer can be written as a double. Excel will display an
	 * integer. $format is optional.
	 *
	 * Returns	0 : normal termination
	 *		   -2 : row or column out of range
	 *
	 * @access public
	 * @param integer $row	  Zero indexed row
	 * @param integer $col	  Zero indexed column
	 * @param float	  $num	  The number to write
	 * @param mixed	  $format The optional XF format
	 * @return integer
	 */
	function writeNumber($row, $col, $num, $format = null)
	{
		$cell = $this->_getPhpExcelCell($row, $col, $format);
		$cell->setValueExplicit($num, PHPExcel_Cell_DataType::TYPE_NUMERIC);
		$this->_releasePhpExcelCell($cell);
	}

	/**
	 * Write a string to the specified row and column (zero indexed).
	 * NOTE: there is an Excel 5 defined limit of 255 characters.
	 * $format is optional.
	 * Returns	0 : normal termination
	 *		   -2 : row or column out of range
	 *		   -3 : long string truncated to 255 chars
	 *
	 * @access public
	 * @param integer $row	  Zero indexed row
	 * @param integer $col	  Zero indexed column
	 * @param string  $str	  The string to write
	 * @param mixed	  $format The XF format for the cell
	 * @return integer
	 */
	function writeString($row, $col, $str, $format = null)
	{
		$cell = $this->_getPhpExcelCell($row, $col, $format);
		$cell->setValueExplicit($str, PHPExcel_Cell_DataType::TYPE_STRING);
		$this->_releasePhpExcelCell($cell);
	}

	/**
	 * Writes a note associated with the cell given by the row and column.
	 * NOTE records don't have a length limit.
	 *
	 * @access public
	 * @param integer $row	  Zero indexed row
	 * @param integer $col	  Zero indexed column
	 * @param string  $note	  The note to write
	 */
	function writeNote($row, $col, $note)
	{
		return $this->raiseError('Not implemented');
	}

	/**
	 * Write a blank cell to the specified row and column (zero indexed).
	 * A blank cell is used to specify formatting without adding a string
	 * or a number.
	 *
	 * A blank cell without a format serves no purpose. Therefore, we don't write
	 * a BLANK record unless a format is specified.
	 *
	 * Returns	0 : normal termination (including no format)
	 *		   -1 : insufficient number of arguments
	 *		   -2 : row or column out of range
	 *
	 * @access public
	 * @param integer $row	  Zero indexed row
	 * @param integer $col	  Zero indexed column
	 * @param mixed	  $format The XF format
	 */
	function writeBlank($row, $col, $format = null)
	{
		$cell = $this->_getPhpExcelCell($row, $col, $format);
		$cell->setValueExplicit(null, PHPExcel_Cell_DataType::TYPE_NULL);
		$this->_releasePhpExcelCell($cell);
	}

	/**
	 * Write a formula to the specified row and column (zero indexed).
	 * The textual representation of the formula is passed to the parser in
	 * Parser.php which returns a packed binary string.
	 *
	 * Returns	0 : normal termination
	 *		   -1 : formula errors (bad formula)
	 *		   -2 : row or column out of range
	 *
	 * @access public
	 * @param integer $row	   Zero indexed row
	 * @param integer $col	   Zero indexed column
	 * @param string  $formula The formula text string
	 * @param mixed	  $format  The optional XF format
	 * @return integer
	 */
	function writeFormula($row, $col, $formula, $format = null)
	{
		$cell = $this->_getPhpExcelCell($row, $col, $format);
		$cell->setValueExplicit($formula, PHPExcel_Cell_DataType::TYPE_FORMULA);
		$this->_releasePhpExcelCell($cell);
	}

	/**
	 * Write a hyperlink.
	 * This is comprised of two elements: the visible label and
	 * the invisible link. The visible label is the same as the link unless an
	 * alternative string is specified. The label is written using the
	 * writeString() method. Therefore the 255 characters string limit applies.
	 * $string and $format are optional.
	 *
	 * The hyperlink can be to a http, ftp, mail, internal sheet (not yet), or external
	 * directory url.
	 *
	 * Returns	0 : normal termination
	 *		   -2 : row or column out of range
	 *		   -3 : long string truncated to 255 chars
	 *
	 * @access public
	 * @param integer $row	  Row
	 * @param integer $col	  Column
	 * @param string  $url	  URL string
	 * @param string  $string Alternative label
	 * @param mixed	  $format The cell format
	 * @return integer
	 */
	function writeUrl($row, $col, $url, $string = '', $format = null)
	{
		return $this->raiseError('Not implemented');
	}

	function _getPhpExcelCell($row, $col, $format = null)
	{
		$cell = $this->_peWorksheet->getCellByColumnAndRow($col, $row + 1);

		if ($format != null) {
			$style = $format->_getPhpExcelStyle();
			$cell->setXfIndex($style->getIndex());
		}

		return $cell;
	}

	function _releasePhpExcelCell($cell)
	{
		if ($this->_checkNotify) {
			if (!$this->_peWorksheet->getNotifyCacheControllerEnabled()) {
				$this->_cacheController->addCacheData($cell->getCoordinate(), $cell);
			}
		}
	}

	/**
	 * Write an array of values as a row
	 *
	 * @access public
	 * @param integer $row	  The row we are writing to
	 * @param integer $col	  The first col (leftmost col) we are writing to
	 * @param array	  $val	  The array of values to write
	 * @param mixed	  $format The optional format to apply to the cell
	 * @return mixed PEAR_Error on failure
	 */
	function writeRow($row, $col, $val, $format=0)
	{
		$retval = '';
		if (is_array($val)) {
			foreach($val as $v) {
				if (is_array($v)) {
					$this->writeCol($row, $col, $v, $format);
				} else {
					$this->write($row, $col, $v, $format);
				}
				$col++;
			}
		} else {
			$retval = new PEAR_Error('$val needs to be an array');
		}
		return($retval);
	}

	/**
	 * Write an array of values as a column
	 *
	 * @access public
	 * @param integer $row	  The first row (uppermost row) we are writing to
	 * @param integer $col	  The col we are writing to
	 * @param array	  $val	  The array of values to write
	 * @param mixed	  $format The optional format to apply to the cell
	 * @return mixed PEAR_Error on failure
	 */
	function writeCol($row, $col, $val, $format=0)
	{
		$retval = '';
		if (is_array($val)) {
			foreach($val as $v) {
				$this->write($row, $col, $v, $format);
				$row++;
			}
		} else {
			$retval = new PEAR_Error('$val needs to be an array');
		}
		return($retval);
	}

	/**
	* This method sets the properties for outlining and grouping. The defaults
	* correspond to Excel's defaults.
	*
	* @param bool $visible
	* @param bool $symbols_below
	* @param bool $symbols_right
	* @param bool $auto_style
	*/
	function setOutline($visible = true, $symbols_below = true, $symbols_right = true, $auto_style = false)
	{
		// $this->_peWorksheet->setShowSummaryBelow($symbols_below);
		// $this->_peWorksheet->setShowSummaryRight($symbols_right);
		return $this->raiseError('Not implemented');
	}

	/**
	 * Merges the area given by its arguments.
	 * This is an Excel97/2000 method. It is required to perform more complicated
	 * merging than the normal setAlign('merge').
	 *
	 * @access public
	 * @param integer $first_row First row of the area to merge
	 * @param integer $first_col First column of the area to merge
	 * @param integer $last_row	 Last row of the area to merge
	 * @param integer $last_col	 Last column of the area to merge
	 */
	function mergeCells($first_row, $first_col, $last_row, $last_col)
	{
		$this->_peWorksheet->mergeCellsByColumnAndRow($first_col, $first_row + 1, $last_col, $last_row + 1);
	}

	function keepLeadingZeros($keep = true)
	{
		// skip
	}

	/**
	* This method is used to set the height and format for a row.
	*
	* @access public
	* @param integer $row	 The row to set
	* @param integer $height Height we are giving to the row.
	*						 Use null to set XF without setting height
	* @param mixed	 $format XF format we are giving to the row
	* @param bool	 $hidden The optional hidden attribute
	* @param integer $level	 The optional outline level for row, in range [0,7]
	*/
	function setRow($row, $height, $format = 0, $hidden = false, $level = 0)
	{
		$rowDimension = $this->_peWorksheet->getRowDimension($row + 1);
		$rowDimension->setRowHeight($height);
		$rowDimension->setVisible(!$hidden);
		$rowDimension->setOutlineLevel($level);
	}

	/**
	 * Insert a 24bit bitmap image in a worksheet.
	 *
	 * @access public
	 * @param integer $row	   The row we are going to insert the bitmap into
	 * @param integer $col	   The column we are going to insert the bitmap into
	 * @param string  $bitmap  The bitmap filename
	 * @param integer $x	   The horizontal position (offset) of the image inside the cell.
	 * @param integer $y	   The vertical position (offset) of the image inside the cell.
	 * @param integer $scale_x The horizontal scale
	 * @param integer $scale_y The vertical scale
	 */
	function insertBitmap($row, $col, $bitmap, $x = 0, $y = 0, $scale_x = 1, $scale_y = 1)
	{
		$img = PHPExcel_Shared_Drawing::imagecreatefrombmp($bitmap);
		$objDrawing = new PHPExcel_Worksheet_MemoryDrawing();
		$objDrawing->setCoordinates(PHPExcel_Cell::stringFromColumnIndex($col) . ($row + 1));
		$objDrawing->setImageResource($img);
		$objDrawing->setWorksheet($this->_peWorksheet);
	}
}
