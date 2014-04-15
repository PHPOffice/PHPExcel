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
 * @package	PHPExcel_Shared
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license	http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version	##VERSION##, ##DATE##
 */

if (!defined('DATE_W3C')) {
  define('DATE_W3C', 'Y-m-d\TH:i:sP');
}

if (!defined('DEBUGMODE_ENABLED')) {
  define('DEBUGMODE_ENABLED', false);
}


/**
 * PHPExcel_Shared_XMLWriter
 *
 * @category   PHPExcel
 * @package	PHPExcel_Shared
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Shared_XMLWriter extends XMLWriter {
	const FLUSH_FREQUENCY = 100;

	/**
	 * Temporary filename
	 *
	 * @var string
	 */
	private $_filename = '';

	private $counter;

	/**
	 * Create a new PHPExcel_Shared_XMLWriter instance
	 *
	 * @param string	$filename	Temporary storage filename (empty to use memory)
	 */
	public function __construct($filename = NULL) {
		// Open temporary storage
		if (!$filename) {
			$this->openMemory();
		} else {
			$this->_filename = $filename;

			// Open storage
			if ($this->openUri($this->_filename) === false) {
				// Fallback to memory...
				$this->_filename = '';
				$this->openMemory();
			}
		}

		// Set default values
		if (DEBUGMODE_ENABLED) {
			$this->setIndent(true);
		}
		$this->counter = self::FLUSH_FREQUENCY;
	}

	public function flush($empty=true)
	{
		parent::flush($empty);
		$this->counter = self::FLUSH_FREQUENCY;
	}

	private function flushIfNecessary()
	{
		if ($this->_filename == '')
			return;

		--$this->counter;
		if ($this->counter <= 0)
			$this->flush();
	}

	public function endElement()
	{
		$ret = parent::endElement();
		$this->flushIfNecessary();
		return $ret;
	}

	public function writeElement($name, $content=null)
	{
		$ret = parent::writeElement($name, $content);
		$this->flushIfNecessary();
		return $ret;
	}

	public function writeAttribute($name, $value)
	{
		$ret = parent::writeAttribute($name, $value);
		$this->flushIfNecessary();
		return $ret;
	}

	/**
	 * Get written data
	 *
	 * @return $data
	 */
	public function getData() {
		if ($this->_filename == '') {
			return $this->outputMemory(true);
		} else {
			$this->flush();
			return file_get_contents($this->_filename);
		}
	}

	/**
	 * Get the temporary file name.
	 *
	 * @return filename;
	 */
	public function getFileName() {
		return $this->_filename;
	}

	/**
	 * Fallback method for writeRaw, introduced in PHP 5.2
	 *
	 * @param string $text
	 * @return string
	 */
	public function writeRawData($text)
	{
		if (is_array($text)) {
			$text = implode("\n",$text);
		}

		if (method_exists($this, 'writeRaw')) {
			return $this->writeRaw(htmlspecialchars($text));
		}

		return $this->text($text);
	}
}
