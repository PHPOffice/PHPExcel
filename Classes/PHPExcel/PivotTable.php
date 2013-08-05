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
class PHPExcel_PivotTable
{
  /**
	 * PivotTable target
	 *
	 * @var string
	 */
	private $_target = '';
  
  /**
	 * PivotTable name
	 *
	 * @var string
	 */
	private $_name = '';
  
  /**
	 * PivotTable id
	 *
	 * @var string
	 */
	private $_id = '';
  
  /**
	 * PivotTable data
	 *
	 * @var string
	 */
	private $_data = '';
  
  /**
	 * Collection of PivotCacheDefinition
	 *
	 * @var string
	 */
	private $_pivotCacheDefinitionCollection = array();

	/**
	 * Worksheet
	 *
	 * @var PHPExcel_Worksheet
	 */
	private $_worksheet = null;
  
  /**
	 * Create a new PHPExcel_PivotTable
	 */
  public function __construct($id,$target,$xmlData)
	{
    $this->setId($id);
    $this->setTarget($target);
    // patch to correct  dxfId="XX"
    $this->_data = preg_replace("/ dxfId=\"..\"/i","",$xmlData);
  }
  
  public function setId($id){       
    $this->_id = preg_replace("/^rId/i","",$id);
  }
  
  public function getId(){
    return $this->_id;
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
  
  public function addPivotCacheDefinition(PHPExcel_PivotTable_PivotCacheDefinition $PivotCacheDefinition){
   $this->_pivotCacheDefinitionCollection[] = $PivotCacheDefinition;
  }
  
  public function getXmlData(){
    return $this->_data;
  }
  
  /**
	 * Set Worksheet
	 *
	 * @param	PHPExcel_Worksheet	$pValue
	 * @throws	PHPExcel_PivotTable_Exception
	 * @return PHPExcel_PivotTable
	 */
	public function setWorksheet(PHPExcel_Worksheet $pValue = null) {
		$this->_worksheet = $pValue;

		return $this;
	}
  
  public function getPivotCacheDefinitionCollection(){
    return $this->_pivotCacheDefinitionCollection ;
  }
  
  public static function normalizePath($path)
  {
      $parts = array();// Array to build a new path from the good parts
      $path = str_replace('\\', '/', $path);// Replace backslashes with forwardslashes
      $path = preg_replace('/\/+/', '/', $path);// Combine multiple slashes into a single slash
      $segments = explode('/', $path);// Collect path segments
      $test = '';// Initialize testing variable
      foreach($segments as $segment)
      {
          if($segment != '.')
          {
              $test = array_pop($parts);
              if(is_null($test))
                  $parts[] = $segment;
              else if($segment == '..')
              {
                  if($test == '..')
                      $parts[] = $test;
  
                  if($test == '..' || $test == '')
                      $parts[] = $segment;
              }
              else
              {
                  $parts[] = $test;
                  $parts[] = $segment;
              }
          }
      }
      return implode('/', $parts);
  }
}