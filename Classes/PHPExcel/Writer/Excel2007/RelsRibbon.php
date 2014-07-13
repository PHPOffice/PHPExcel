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
 * @package    PHPExcel_Writer_Excel2007
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version     ##VERSION##, ##DATE##
 */


/**
 * PHPExcel_Writer_Excel2007_RelsRibbon
 *
 * @category   PHPExcel
 * @package    PHPExcel_Writer_Excel2007
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Writer_Excel2007_RelsRibbon extends PHPExcel_Writer_Excel2007_WriterPart
{
	/**
	 * Write relationships for additional objects of custom UI (ribbon)
	 *
	 * @param 	PHPExcel	$pPHPExcel
	 * @param	ZipArchive	$objZip
	 * @param	string		$filename
	 * @throws 	PHPExcel_Writer_Exception
	 */
	public function addRibbonRelationshipsToZip(PHPExcel $pPHPExcel, $objZip, $filename) {
		// Create XML writer
		$objWriter = $this->createXMLWriter();

		// XML header
		$objWriter->startDocument('1.0','UTF-8','yes');

		// Relationships
		$objWriter->startElement('Relationships');
		$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
		$localRels=$pPHPExcel->getRibbonBinObjects('names');
		if(is_array($localRels)){
			foreach($localRels as $aId=>$aTarget){
				$objWriter->startElement('Relationship');
				$objWriter->writeAttribute('Id', $aId);
				$objWriter->writeAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
				$objWriter->writeAttribute('Target', $aTarget);
				$objWriter->endElement();//Relationship
			}
		}
		$objWriter->endElement();//Relationships

		// Add the generated file to the Zip file.
		$this->addXMLToZip($objWriter, $objZip, $filename);
	}

}
