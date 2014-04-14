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
 * @package	PHPExcel_Shared_TempFileTracker
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license	http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version	##VERSION##, ##DATE##
 */

/**
 * PHPExcel_Shared_TempFileTracker
 *
 * @category   PHPExcel
 * @package	PHPExcel_Shared_TempFileTracker
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Shared_TempFileTracker
{
	/**
	 * List of created temporary files
	 *
	 * @var array
	 */
	private $_files;

	function __construct ()
	{
		$this->_files = array ();
	}

	function __destruct ()
	{
		$this->removeAllFiles ();
	}

	/**
	 * Register a new temporary file.
	 *
	 * @param string $filename
	 */
	public function registerFile ($filename)
	{
		$this->_files[] = $filename;
	}

	/**
	 * Generate a new temporary file name, and register it.
	 *
	 * @param string $prefix the new filename prefix.
	 * @return string the new temporary filename.
	 */
	public function getNewFile ($prefix)
	{
		$filename = @tempnam (PHPExcel_Shared_File::sys_get_temp_dir(), $prefix);
		$this->registerFile ($filename);
		return $filename;
	}

	/**
	 * Remove all the registered files.
	 */
	public function removeAllFiles ()
	{
		foreach ($this->_files as $file)
			unlink ($file);
		$this->_files = array ();
	}
}

?>
