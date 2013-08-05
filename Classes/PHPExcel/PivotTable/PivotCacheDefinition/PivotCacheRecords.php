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
 * @category	PHPExcel
 * @package		PHPExcel_Reader_Excel2007
 * @copyright	Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license		http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version		##VERSION##, ##DATE##
 */

/**
 * PHPExcel_PivotTable
 *
 * @category	PHPExcel
 * @package		PHPExcel_Reader_Excel2007
 * @copyright	Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_PivotTable_PivotCacheDefinition_PivotCacheRecords
{
  /**
	 * PivotCacheRecords target
	 *
	 * @var string
	 */
	private $_target = '';
  
  /**
	 * PivotCacheRecords name
	 *
	 * @var string
	 */
	private $_name = '';
  
  /**
	 * PivotCacheRecords id
	 *
	 * @var string
	 */
	private $_id = '';
  
  /**
	 * PivotCacheRecords data
	 *
	 * @var string
	 */
	private $_data = '';

	/**
	 * PivotCacheDefinition
	 *
	 * @var PHPExcel_PivotTable_PivotCacheDefinition
	 */
	private $_pivotCacheDefinition = null;
  
  public function __construct($id,$target,$xmlData)
	{
    $this->setId($id);
    $this->setTarget($target);
    $this->_data = $xmlData;
  }
  
  public function getName(){
    return $this->_name;
  }
  
  public function getTarget(){
    return $this->_target;
  }
  
  public function setTarget($target){
    $this->_target = $target;
    $this->_name = basename($target);
  }
  
  public function getId(){
    return $this->_id;
  }
  
  public function setId($id){       
    $this->_id = preg_replace("/^rId/i","",$id);
  }
  
  public function getXmlData(){
    return $this->_data;
  }   
}