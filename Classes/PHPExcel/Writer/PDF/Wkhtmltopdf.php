<?php
/**
 *  PHPExcel
 *
 *  Copyright (c) 2006 - 2013 PHPExcel
 *
 *  This library is free software; you can redistribute it and/or
 *  modify it under the terms of the GNU Lesser General Public
 *  License as published by the Free Software Foundation; either
 *  version 2.1 of the License, or (at your option) any later version.
 *
 *  This library is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 *  Lesser General Public License for more details.
 *
 *  You should have received a copy of the GNU Lesser General Public
 *  License along with this library; if not, write to the Free Software
 *  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 *  @author      Artur Kozubski <a.kozubski@synerway.pl> - SYNERWAY S.A. (http://synerwaygroup.pl) [Dec 18, 2013 2:13:07 PM]
 *  @category    PHPExcel
 *  @package     PHPExcel_Writer_PDF
 *  @copyright   Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 *  @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 *  @version     ##VERSION##, ##DATE##
 */

/**
 *  PHPExcel_Writer_PDF_Wkhtmltopdf
 *
 *  @author      Artur Kozubski <a.kozubski@synerway.pl> - SYNERWAY S.A. (http://synerwaygroup.pl) [Dec 18, 2013 2:13:07 PM]
 *  @category    PHPExcel
 *  @package     PHPExcel_Writer_PDF
 *  @copyright   Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 *  @todo        Add compatibility with the Windows version of wkhtmltopdf (streams)
 */
class PHPExcel_Writer_PDF_Wkhtmltopdf extends PHPExcel_Writer_PDF_Core implements PHPExcel_Writer_IWriter
{
    /**
     *  Create a new PHPExcel_Writer_PDF
     *
     *  @param PHPExcel $phpExcel PHPExcel object
     */
    public function __construct(PHPExcel $phpExcel)
    {
        parent::__construct($phpExcel);
    }

    /**
     *  Save PHPExcel to file
     *
     *  @param     string     $pFilename   Name of the file to save as
     *  @throws    PHPExcel_Writer_Exception
     */
    public function save($pFilename = NULL)
    {
        $fileHandle = parent::prepareForSave($pFilename);

        // Default PDF paper size
        $paperSize = 'LETTER';    //    Letter    (8.5 in. by 11 in.)

        // Check for paper size and page orientation
        if (is_null($this->getSheetIndex())) {
            $orientation = ($this->_phpExcel->getSheet(0)->getPageSetup()->getOrientation()
                == PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE)
                    ? 'L'
                    : 'P';
            $printPaperSize = $this->_phpExcel->getSheet(0)->getPageSetup()->getPaperSize();
            $printMargins = $this->_phpExcel->getSheet(0)->getPageMargins();
        } else {
            $orientation = ($this->_phpExcel->getSheet($this->getSheetIndex())->getPageSetup()->getOrientation()
                == PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE)
                    ? 'L'
                    : 'P';
            $printPaperSize = $this->_phpExcel->getSheet($this->getSheetIndex())->getPageSetup()->getPaperSize();
            $printMargins = $this->_phpExcel->getSheet($this->getSheetIndex())->getPageMargins();
        }
        $this->setOrientation($orientation);

        // Override Page Orientation
        if (!is_null($this->getOrientation())) {
            $orientation = ($this->getOrientation() == PHPExcel_Worksheet_PageSetup::ORIENTATION_DEFAULT)
                ? PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT
                : $this->getOrientation();
        }
        $orientation = ($orientation == 'L' ? 'Landscape' : 'Portrait');

        // Override Paper Size
        if (!is_null($this->getPaperSize())) {
            $printPaperSize = $this->getPaperSize();
        }

        if (isset(self::$_paperSizes[$printPaperSize])) {
            $paperSize = ucfirst(self::$_paperSizes[$printPaperSize]);
        }

        // Process pipes descriptor
        $descriptorSpec = array(
            0 => array('pipe', 'r'), // stdin - read by process from standard input
            1 => array('pipe', 'w'), // stdout - write by process to the standard output
            2 => array('pipe', 'w')  // stderr - write by process to the standard error
        );

        // Create PDF

        // Orientation and page size
        $options = "-q -O $orientation -s $paperSize ";

        // Document info
        $options .= '--title "' . $this->_phpExcel->getProperties()->getTitle() . '"';
        // unavailable $this->_phpExcel->getProperties()->getCreator();
        // unavailable $this->_phpExcel->getProperties()->getSubject();
        // unavailable $this->_phpExcel->getProperties()->getKeywords();
        // unavailable $this->_phpExcel->getProperties()->getCreator();

        $pipes = null;

        $cmdLine = PHPExcel_Settings::getPdfRendererPath() . " $options - -";
        $process = proc_open($cmdLine, $descriptorSpec, $pipes);

        if (is_resource($process)) {
            // Send HTML contents to stdin
            fwrite($pipes[0], $this->generateHTMLHeader(false) . $this->generateSheetData() . $this->generateHTMLFooter());
            fclose($pipes[0]);

            // Read process output from stdout and stderr
            $stdout = stream_get_contents($pipes[1]);
            $stderr = stream_get_contents($pipes[2]);

            // Close pipes
            fclose($pipes[1]);
            fclose($pipes[2]);

            // Close process
            $returnCode = proc_close($process);

            // Error checking
            if ($returnCode != 0 && $stderr != '') {
                parent::restoreStateAfterSave($fileHandle);
                throw new PHPExcel_Writer_Exception('Conversion error: wkhtmltopdf process returned with error code ' . $returnCode . ' [' . $stderr . ']');
            }

            // Write to file
            fwrite($fileHandle, $stdout);
        } else {
            parent::restoreStateAfterSave($fileHandle);
            throw new PHPExcel_Writer_Exception('Unable to run wkhtmltopdf process');
        }

        parent::restoreStateAfterSave($fileHandle);
    }
}
