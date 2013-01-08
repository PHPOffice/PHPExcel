<?php

require_once('PEAR.php');

/**
 * PHPExcel_Adapter_Spreadsheet_Excel_Writer
 *
 * Implements Spreadsheet_Excel_Writer interface (http://pear.php.net/package/Spreadsheet_Excel_Writer/)
 *
 * Patch your old reports:
 *
 * <pre>
 * -require_once 'Spreadsheet/Excel/Writer.php';
 * +require_once 'PHPExcel.php';
 *
 * -$workbook = new Spreadsheet_Excel_Writer();
 * +$workbook = new PHPExcel_Adapter_Spreadsheet_Excel_Writer();
 * </pre>
 *
 * And test them, implementation is partial.
 *
 * @category   PHPExcel
 * @package    PHPExcel_Adapter_Spreadsheet_Excel_Writer
 */
class PHPExcel_Adapter_Spreadsheet_Excel_Writer extends PEAR
{
	private $_peWorkbook;
	private $_filename;
	private $_writerType;
	private $_firstSheet = true;
	private $_worksheets = array();

	private $_customColors = array();

	/**
	 * Constructor
	 *
	 * With $peSettings parameter it is possible to change types of
	 * CachedObjectStorage and Writer:
	 *
	 * <code>
	 * $peSettings = array(
	 *	   'CachedObjectStorage' => array(
	 *		   'type' => 'Memory',
	 *		   'args' => array(),
	 *		   ),
	 *	   'Writer' => array(
	 *		   'type' => 'Excel2007'
	 *		   ),
	 *	   );
	 *
	 * $workbook = new PHPExcel_Adapter_Spreadsheet_Excel_Writer($filename, $peSettings);
	 * </code>
	 *
	 * @param string $filename for storing the workbook. "-" for writing to stdout.
	 * @param array	 $peSettings PHPExcel settings
	 * @access public
	 */
	function PHPExcel_Adapter_Spreadsheet_Excel_Writer($filename = '', $peSettings = array())
	{
		if (isset($peSettings['CachedObjectStorage']['type'])) {
			$cacheType = $peSettings['CachedObjectStorage']['type'];
			$cacheArgs = array();
			if (isset($peSettings['CachedObjectStorage']['args'])) {
				$cacheArgs = $peSettings['CachedObjectStorage']['args'];
			}
		} else {
			$cacheType = defined('PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp_bucket') ?
				PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp_bucket :
				PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
			$cacheArgs = array('memoryCacheSize'  => '4MB');
		}

		PHPExcel_Settings::setCacheStorageMethod($cacheType, $cacheArgs);

		if (isset($peSettings['Writer']['type'])) {
			$this->_writerType = $peSettings['Writer']['type'];
		} else {
			$this->_writerType = 'Excel5';
		}

		$this->_peWorkbook = new PHPExcel();

		$defaultStyle = $this->_peWorkbook->getDefaultStyle();
		$defaultStyle->getFont()->setName('Arial');
		$defaultStyle->getFont()->setSize(10);

		$this->_filename = $filename;
	}

	function _getPhpExcelWorkbook()
	{
		return $this->_peWorkbook;
	}

	/**
	 * Calls finalization methods.
	 * This method should always be the last one to be called on every workbook
	 *
	 * @access public
	 * @return mixed true on success. PEAR_Error on failure
	 */
	function close()
	{
		$filename = $this->_filename;
		$tmpfile = false;

		if ($this->_filename == '' ||
			$this->_filename == '-' ||
			strtolower($this->_filename) == 'php://output' ||
			strtolower($this->_filename) == 'php://stdout')
		{
			$filename = tempnam('/tmp', 'PHPExcel_Adapter_Spreadsheet_Excel_Writer');
			$tmpfile = true;
		}

		if (method_exists('PHPExcel_Worksheet', 'setNotifyCacheControllerEnabled')) {
			foreach ($this->_peWorkbook->getWorksheetIterator() as $sheet) {
				$sheet->setNotifyCacheControllerEnabled(true);
			}
		}

		if (!class_exists('ZipArchive')) {
			PHPExcel_Settings::setZipClass(PHPExcel_Settings::PCLZIP);
		}

		$objWriter = PHPExcel_IOFactory::createWriter($this->_peWorkbook, $this->_writerType);

		$objWriter->setPreCalculateFormulas(false);
		if (method_exists($objWriter, 'setGarbageCollectEnabled')) {
			$objWriter->setGarbageCollectEnabled(false);
		}
		if (method_exists($objWriter, 'setSortCellCollectionEnabled')) {
			$objWriter->setSortCellCollectionEnabled(false);
		}

		$objWriter->save($filename);

		if ($tmpfile) {
			$fp = fopen($filename, 'r');
			fpassthru($fp);
			unlink($filename);
		}

		return true;
	}

	/**
	 * An accessor for the _worksheets[] array
	 * Returns an array of the worksheet objects in a workbook
	 * It actually calls to worksheets()
	 *
	 * @access public
	 * @see worksheets()
	 * @return array
	 */
	function sheets()
	{
		return $this->worksheets();
	}

	/**
	 * An accessor for the _worksheets[] array.
	 * Returns an array of the worksheet objects in a workbook
	 *
	 * @access public
	 * @return array
	 */
	function worksheets()
	{
		return $this->_worksheets;
	}

	function setVersion()
	{
		// skip
	}

	function setBIFF8InputEncoding($encoding)
	{
		// skip
	}

	/**
	 * Set the country identifier for the workbook
	 *
	 * @access public
	 * @param integer $code Is the international calling country code for the
	 *						chosen country.
	 */
	function setCountry($code)
	{
		return $this->raiseError('Not implemented');
	}

	/**
	 * Add a new worksheet to the Excel workbook.
	 * If no name is given the name of the worksheet will be Sheet$i, with
	 * $i in [1..].
	 *
	 * @access public
	 * @param string $name the optional name of the worksheet
	 * @return mixed worksheet object on success, PEAR_Error on failure
	 */
	function addWorksheet($name = '')
	{
		if ($name == '') {
			$name = 'Sheet' . (count($this->_worksheets) + 1);
		}

		$worksheet = new PHPExcel_Adapter_Spreadsheet_Excel_Writer_Worksheet($this, $name,
																			 $this->_firstSheet);
		$this->_firstSheet = false;
		$this->_worksheets[] = $worksheet;

		return $worksheet;
	}

	/**
	 * Add a new format to the Excel workbook.
	 * Also, pass any properties to the Format constructor.
	 *
	 * @access public
	 * @param array $properties array with properties for initializing the format.
	 * @return PHPExcel_Adapter_Spreadsheet_Excel_Writer_Format format object
	 */
	function addFormat($properties = array())
	{
		return new PHPExcel_Adapter_Spreadsheet_Excel_Writer_Format($this, $properties);
	}

	function addValidator()
	{
		return $this->raiseError('Not implemented');
	}

	function _getCustomColor($index)
	{
		if(!isset($this->_customColors[$index])) {
			return $this->raiseError("Unknown custom color index: " . $index);
		}

		return $this->_customColors[$index];
	}

	/**
	 * Change the RGB components of the elements in the colour palette.
	 *
	 * @access public
	 * @param integer $index colour index
	 * @param integer $red	 red RGB value [0-255]
	 * @param integer $green green RGB value [0-255]
	 * @param integer $blue	 blue RGB value [0-255]
	 * @return integer The palette index for the custom color
	 */
	function setCustomColor($index, $red, $green, $blue)
	{
		if ($index < 8 || $index > 64) {
			return $this->raiseError("Color index $index outside range: 8 <= index <= 64");
		}

		if (($red	< 0 || $red	  > 255) ||
			($green < 0 || $green > 255) ||
			($blue	< 0 || $blue  > 255))
		{
			return $this->raiseError("Color component outside range: 0 <= color <= 255");
		}

		$this->_customColors[$index] = sprintf("FF%X%X%X", $red, $green, $blue);

		return $index;
	}

	/**
	 * Send HTTP headers for the Excel file.
	 *
	 * @param string $filename The filename to use for HTTP headers
	 * @access public
	 */
	function send($filename)
	{
		header("Content-type: application/vnd.ms-excel");
		header("Content-Disposition: attachment; filename=$filename");
		header("Expires: 0");
		header("Cache-Control: must-revalidate, post-check=0,pre-check=0");
		header("Pragma: public");
	}

	/**
	 * Utility function for writing formulas
	 * Converts a cell's coordinates to the A1 format.
	 *
	 * @access public
	 * @static
	 * @param integer $row Row for the cell to convert (0-indexed).
	 * @param integer $col Column for the cell to convert (0-indexed).
	 * @return string The cell identifier in A1 format
	 */
	function rowcolToCell($row, $col)
	{
		return PHPExcel_Cell::stringFromColumnIndex($col) . ($row + 1);
	}
}
