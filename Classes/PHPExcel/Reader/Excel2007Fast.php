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
 * @package    PHPExcel_Reader
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.8.0, 2014-03-02
 */


/** PHPExcel root directory */
if (!defined('PHPEXCEL_ROOT')) {
	/**
	 * @ignore
	 */
	define('PHPEXCEL_ROOT', dirname(__FILE__) . '/../../');
	require(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
}

/**
 * PHPExcel_Reader_Excel2007Fast
 *
 * @category	PHPExcel
 * @package	PHPExcel_Reader
 * @copyright	Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Reader_Excel2007Fast extends PHPExcel_Reader_Abstract implements PHPExcel_Reader_IReader
{
	/**
	 * Create a new PHPExcel_Reader_Excel2007 instance
	 */
	public function __construct() {
		$this->_readFilter = new PHPExcel_Reader_DefaultReadFilter();
		$this->_referenceHelper = PHPExcel_ReferenceHelper::getInstance();
	}


	/**
	 * Can the current PHPExcel_Reader_IReader read the file?
	 *
	 * @param 	string 		$pFilename
	 * @return 	boolean
	 * @throws PHPExcel_Reader_Exception
	 */
	public function canRead($pFilename)
	{
		// Check if file exists
		if (!file_exists($pFilename)) {
			throw new PHPExcel_Reader_Exception("Could not open " . $pFilename . " for reading! File does not exist.");
		}

        $zipClass = PHPExcel_Settings::getZipClass();

		// Check if zip class exists
//		if (!class_exists($zipClass, FALSE)) {
//			throw new PHPExcel_Reader_Exception($zipClass . " library is not enabled");
//		}

		$xl = false;
		// Load file
		$zip = new $zipClass;
		if ($zip->open($pFilename) === true) {
			// check if it is an OOXML archive
			$rels = simplexml_load_string($this->_getFromZipArchive($zip, "_rels/.rels"), 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions());
			if ($rels !== false) {
				foreach ($rels->Relationship as $rel) {
					switch ($rel["Type"]) {
						case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
							if (basename($rel["Target"]) == 'workbook.xml') {
								$xl = true;
							}
							break;

					}
				}
			}
			$zip->close();
		}

		return $xl;
	}


	/**
	 * Reads names of the worksheets from a file, without parsing the whole file to a PHPExcel object
	 *
	 * @param 	string 		$pFilename
	 * @throws 	PHPExcel_Reader_Exception
	 */
	public function listWorksheetNames($pFilename)
	{
		// Check if file exists
		if (!file_exists($pFilename)) {
			throw new PHPExcel_Reader_Exception("Could not open " . $pFilename . " for reading! File does not exist.");
		}

		$worksheetNames = array();

        $zipClass = PHPExcel_Settings::getZipClass();

		$zip = new $zipClass;
		$zip->open($pFilename);

		//	The files we're looking at here are small enough that simpleXML is more efficient than XMLReader
		$rels = simplexml_load_string(
		    $this->_getFromZipArchive($zip, "_rels/.rels", 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions())
		); //~ http://schemas.openxmlformats.org/package/2006/relationships");
		foreach ($rels->Relationship as $rel) {
			switch ($rel["Type"]) {
				case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
					$xmlWorkbook = simplexml_load_string(
					    $this->_getFromZipArchive($zip, "{$rel['Target']}", 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions())
					);  //~ http://schemas.openxmlformats.org/spreadsheetml/2006/main");

					if ($xmlWorkbook->sheets) {
						foreach ($xmlWorkbook->sheets->sheet as $eleSheet) {
							// Check if sheet should be skipped
							$worksheetNames[] = (string) $eleSheet["name"];
						}
					}
			}
		}

		$zip->close();

		return $worksheetNames;
	}


	/**
	 * Return worksheet info (Name, Last Column Letter, Last Column Index, Total Rows, Total Columns)
	 *
	 * @param   string     $pFilename
	 * @throws   PHPExcel_Reader_Exception
	 */
	public function listWorksheetInfo($pFilename)
	{
		// Check if file exists
		if (!file_exists($pFilename)) {
			throw new PHPExcel_Reader_Exception("Could not open " . $pFilename . " for reading! File does not exist.");
		}

		$worksheetInfo = array();

        $zipClass = PHPExcel_Settings::getZipClass();

		$zip = new $zipClass;
		$zip->open($pFilename);

		$rels = simplexml_load_string($this->_getFromZipArchive($zip, "_rels/.rels"), 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions()); //~ http://schemas.openxmlformats.org/package/2006/relationships");
		foreach ($rels->Relationship as $rel) {
			if ($rel["Type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument") {
				$dir = dirname($rel["Target"]);
				$relsWorkbook = simplexml_load_string($this->_getFromZipArchive($zip, "$dir/_rels/" . basename($rel["Target"]) . ".rels"), 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions());  //~ http://schemas.openxmlformats.org/package/2006/relationships");
				$relsWorkbook->registerXPathNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");

				$worksheets = array();
				foreach ($relsWorkbook->Relationship as $ele) {
					if ($ele["Type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet") {
						$worksheets[(string) $ele["Id"]] = $ele["Target"];
					}
				}

				$xmlWorkbook = simplexml_load_string($this->_getFromZipArchive($zip, "{$rel['Target']}"), 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions());  //~ http://schemas.openxmlformats.org/spreadsheetml/2006/main");
				if ($xmlWorkbook->sheets) {
					$dir = dirname($rel["Target"]);
					foreach ($xmlWorkbook->sheets->sheet as $eleSheet) {
						$tmpInfo = array(
							'worksheetName' => (string) $eleSheet["name"],
							'lastColumnLetter' => 'A',
							'lastColumnIndex' => 0,
							'totalRows' => 0,
							'totalColumns' => 0,
						);

						$fileWorksheet = $worksheets[(string) self::array_item($eleSheet->attributes("http://schemas.openxmlformats.org/officeDocument/2006/relationships"), "id")];

						$xml = new XMLReader();
						$res = $xml->open('zip://'.PHPExcel_Shared_File::realpath($pFilename).'#'."$dir/$fileWorksheet", null, PHPExcel_Settings::getLibXmlLoaderOptions());
						$xml->setParserProperty(2,true);

						$currCells = 0;
						while ($xml->read()) {
							if ($xml->name == 'row' && $xml->nodeType == XMLReader::ELEMENT) {
								$row = $xml->getAttribute('r');
								$tmpInfo['totalRows'] = $row;
								$tmpInfo['totalColumns'] = max($tmpInfo['totalColumns'],$currCells);
								$currCells = 0;
							} elseif ($xml->name == 'c' && $xml->nodeType == XMLReader::ELEMENT) {
								$currCells++;
							}
						}
						$tmpInfo['totalColumns'] = max($tmpInfo['totalColumns'],$currCells);
						$xml->close();

						$tmpInfo['lastColumnIndex'] = $tmpInfo['totalColumns'] - 1;
						$tmpInfo['lastColumnLetter'] = PHPExcel_Cell::stringFromColumnIndex($tmpInfo['lastColumnIndex']);

						$worksheetInfo[] = $tmpInfo;
					}
				}
			}
		}

		$zip->close();

		return $worksheetInfo;
	}
	

	public function _getFromZipArchive($archive, $fileName = '')
	{
		// Root-relative paths
		if (strpos($fileName, '//') !== false)
		{
			$fileName = substr($fileName, strpos($fileName, '//') + 1);
		}
		$fileName = PHPExcel_Shared_File::realpath($fileName);

		// Apache POI fixes
		$contents = $archive->getFromName($fileName);
		if ($contents === false)
		{
			$contents = $archive->getFromName(substr($fileName, 1));
		}

		return $contents;
	}

	/**
	 * Loads PHPExcel from file
	 *
	 * @param 	string 		$pFilename
     * @return  PHPExcel
	 * @throws 	PHPExcel_Reader_Exception
	 */
	public function load($pFilename)
	{
		// Check if file exists
		if (!file_exists($pFilename)) {
			throw new PHPExcel_Reader_Exception("Could not open " . $pFilename . " for reading! File does not exist.");
		}

		// It's always working readDataOnly
		$this->_readDataOnly = true;

		// Initialisations
		$excel = new PHPExcel;
		$excel->removeSheetByIndex(0);


        $zipClass = PHPExcel_Settings::getZipClass();

		$zip = new $zipClass;
		$zip->open($pFilename);


		$rels = simplexml_load_string($this->_getFromZipArchive($zip, "_rels/.rels"), 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions()); //~ http://schemas.openxmlformats.org/package/2006/relationships");
		foreach ($rels->Relationship as $rel) {
			switch ($rel["Type"]) {
				case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
					$dir = dirname($rel["Target"]);
					$relsWorkbook = simplexml_load_string($this->_getFromZipArchive($zip, "$dir/_rels/" . basename($rel["Target"]) . ".rels"), 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions());  //~ http://schemas.openxmlformats.org/package/2006/relationships");
					$relsWorkbook->registerXPathNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");

					// Load strings
					$sharedStrings = array();
					$xpath = self::array_item($relsWorkbook->xpath("rel:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings']"));
					$xmlStrings = simplexml_load_string($this->_getFromZipArchive($zip, "$dir/$xpath[Target]"), 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions());  //~ http://schemas.openxmlformats.org/spreadsheetml/2006/main");
					if (isset($xmlStrings) && isset($xmlStrings->si)) {
						foreach ($xmlStrings->si as $val) {
							if (isset($val->t)) {
								$sharedStrings[] = PHPExcel_Shared_String::ControlCharacterOOXML2PHP( (string) $val->t );
							} elseif (isset($val->r)) {
								$sharedStrings[] = $this->_parseRichText($val);
							}
						}
					}

					// Load worksheets
					$worksheets = array();
					foreach ($relsWorkbook->Relationship as $ele) {
						switch($ele['Type']){
						case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet":
							$worksheets[(string) $ele["Id"]] = $ele["Target"];
							break;
						}
					}


					$xmlWorkbook = simplexml_load_string($this->_getFromZipArchive($zip, "{$rel['Target']}"), 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions());  //~ http://schemas.openxmlformats.org/spreadsheetml/2006/main");
					// Set base date
					if ($xmlWorkbook->workbookPr) {
						PHPExcel_Shared_Date::setExcelCalendar(PHPExcel_Shared_Date::CALENDAR_WINDOWS_1900);
						if (isset($xmlWorkbook->workbookPr['date1904'])) {
							if (self::boolean((string) $xmlWorkbook->workbookPr['date1904'])) {
								PHPExcel_Shared_Date::setExcelCalendar(PHPExcel_Shared_Date::CALENDAR_MAC_1904);
							}
						}
					}


					$sheetId = 0; // keep track of new sheet id in final workbook
					$oldSheetId = -1; // keep track of old sheet id in final workbook
					$countSkippedSheets = 0; // keep track of number of skipped sheets
					$mapSheetId = array(); // mapping of sheet ids from old to new
					$accessCells = 0; // count total cells for debug

					// Load data from each worksheets
					if ($xmlWorkbook->sheets) {
						foreach ($xmlWorkbook->sheets->sheet as $eleSheet) {
							++$oldSheetId;

							// Check if sheet should be skipped
							if (isset($this->_loadSheetsOnly) && !in_array((string) $eleSheet["name"], $this->_loadSheetsOnly)) {
								++$countSkippedSheets;
								$mapSheetId[$oldSheetId] = null;
								continue;
							}

							// Map old sheet id in original workbook to new sheet id.
							// They will differ if loadSheetsOnly() is being used
							$mapSheetId[$oldSheetId] = $oldSheetId - $countSkippedSheets;

							// Load sheet
							$docSheet = $excel->createSheet();
							//	Use false for $updateFormulaCellReferences to prevent adjustment of worksheet
							//		references in formula cells... during the load, all formulae should be correct,
							//		and we're simply bringing the worksheet name in line with the formula, not the
							//		reverse
							$docSheet->setTitle((string) $eleSheet["name"],false);
							$fileWorksheet = $worksheets[(string) self::array_item($eleSheet->attributes("http://schemas.openxmlformats.org/officeDocument/2006/relationships"), "id")];
							$xmlSheet = simplexml_load_string($this->_getFromZipArchive($zip, "$dir/$fileWorksheet"), 'SimpleXMLElement', PHPExcel_Settings::getLibXmlLoaderOptions());  //~ http://schemas.openxmlformats.org/spreadsheetml/2006/main");


							// Read row
							if ($xmlSheet && $xmlSheet->sheetData && $xmlSheet->sheetData->row) {
								foreach ($xmlSheet->sheetData->row as $row) {

									foreach ($row->c as $c) {
										$r 					= (string) $c["r"];
										$cellDataType 		= (string) $c["t"];
										$value				= null;

										// Empty cell?
										if ('' == $c->v || null == $c->v)
										{
											continue;
										}

										// Read cell?
										if ($this->getReadFilter() !== NULL) {
											$coordinates = PHPExcel_Cell::coordinateFromString($r);

											if (!$this->getReadFilter()->readCell($coordinates[0], $coordinates[1], $docSheet->getTitle())) {
												continue;
											}
										}

	//									echo 'Reading cell ', $coordinates[0], $coordinates[1], PHP_EOL;
	//									print_r($c);
	//									echo PHP_EOL;
	//									echo 'Cell Data Type is ', $cellDataType, ': ';
	//
										++$accessCells;
										// Read cell!
										switch ($cellDataType) {
											case "s":
												$value = (string)$c->v != '' ? $sharedStrings[intval($c->v)] : '';
												break;
											case "inlineStr":
												var_dump($c);
												$value = $this->_parseRichText($c->is);
												break;
											default:
												$value = $c->v;
												break;
										}
	//									echo 'Value is ', $value, PHP_EOL;

										$cell = $docSheet->getCell($r);
										$cell->setValue($value);
									}
								}
							}

							// Next sheet id
							++$sheetId;
						}
					}

					if ((!$this->_readDataOnly) || (!empty($this->_loadSheetsOnly))) {
						// active sheet index
						$activeTab = intval($xmlWorkbook->bookViews->workbookView["activeTab"]); // refers to old sheet index

						// keep active sheet index if sheet is still loaded, else first sheet is set as the active
						if (isset($mapSheetId[$activeTab]) && $mapSheetId[$activeTab] !== null) {
							$excel->setActiveSheetIndex($mapSheetId[$activeTab]);
						} else {
							if ($excel->getSheetCount() == 0) {
								$excel->createSheet();
							}
							$excel->setActiveSheetIndex(0);
						}
					}
				break;
			}

		}

		unset($sharedStrings);
		unset($worksheets);


		$zip->close();

		unset($zip);
	//	echo 'Total access cells: ', $accessCells, '<br/>', PHP_EOL;
		return $excel;
	}


	private function _parseRichText($is = null) {
		return isset($is->t) ? $is->t : '';
	}

	private static function array_item($array, $key = 0) {
		return (isset($array[$key]) ? $array[$key] : null);
	}


	private static function dir_add($base, $add) {
		return preg_replace('~[^/]+/\.\./~', '', dirname($base) . "/$add");
	}

	private static function boolean($value = NULL)
	{
        if (is_object($value)) {
			$value = (string) $value;
        }
		if (is_numeric($value)) {
			return (bool) $value;
        }
		return ($value === 'true' || $value === 'TRUE');
	}
}
