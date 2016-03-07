<?php
namespace PhpOffice\PhpExcel\Writer\Excel2007;

/**
 * PHPExcel_Writer_Excel2007_Drawing
 *
 * Copyright (c) 2006 - 2016 PHPExcel
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
 * @copyright  Copyright (c) 2006 - 2016 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Drawing extends WriterPart
{
    /**
     * Write drawings to XML format
     *
     * @param     \PhpOffice\PhpExcel\Worksheet    $pWorksheet
     * @param    int                    &$chartRef        Chart ID
     * @param    boolean                $includeCharts    Flag indicating if we should include drawing details for charts
     * @return string          XML Output
     * @throws     \PhpOffice\PhpExcel\Writer\Exception
     */
    public function writeDrawings(\PhpOffice\PhpExcel\Worksheet $pWorksheet = null, &$chartRef, $includeCharts = false)
    {
        // Create XML writer
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new \PhpOffice\PhpExcel\Shared\XMLWriter(\PhpOffice\PhpExcel\Shared\XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new \PhpOffice\PhpExcel\Shared\XMLWriter(\PhpOffice\PhpExcel\Shared\XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // xdr:wsDr
        $objWriter->startElement('xdr:wsDr');
        $objWriter->writeAttribute('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');

        // Loop through images and write drawings
        $i = 1;
        $iterator = $pWorksheet->getDrawingCollection()->getIterator();
        while ($iterator->valid()) {
            $this->writeDrawing($objWriter, $iterator->current(), $i);

            $iterator->next();
            ++$i;
        }

        if ($includeCharts) {
            $chartCount = $pWorksheet->getChartCount();
            // Loop through charts and write the chart position
            if ($chartCount > 0) {
                for ($c = 0; $c < $chartCount; ++$c) {
                    $this->writeChart($objWriter, $pWorksheet->getChartByIndex($c), $c+$i);
                }
            }
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write drawings to XML format
     *
     * @param     \PhpOffice\PhpExcel\Shared\XMLWriter    $objWriter         XML Writer
     * @param     PHPExcel_Chart                $pChart
     * @param     int                            $pRelationId
     * @throws     \PhpOffice\PhpExcel\Writer\Exception
     */
    public function writeChart(\PhpOffice\PhpExcel\Shared\XMLWriter $objWriter = null, PHPExcel_Chart $pChart = null, $pRelationId = -1)
    {
        $tl = $pChart->getTopLeftPosition();
        $tl['colRow'] = \PhpOffice\PhpExcel\Cell::coordinateFromString($tl['cell']);
        $br = $pChart->getBottomRightPosition();
        $br['colRow'] = \PhpOffice\PhpExcel\Cell::coordinateFromString($br['cell']);

        $objWriter->startElement('xdr:twoCellAnchor');

            $objWriter->startElement('xdr:from');
                $objWriter->writeElement('xdr:col', \PhpOffice\PhpExcel\Cell::columnIndexFromString($tl['colRow'][0]) - 1);
                $objWriter->writeElement('xdr:colOff', \PhpOffice\PhpExcel\Shared\Drawing::pixelsToEMU($tl['xOffset']));
                $objWriter->writeElement('xdr:row', $tl['colRow'][1] - 1);
                $objWriter->writeElement('xdr:rowOff', \PhpOffice\PhpExcel\Shared\Drawing::pixelsToEMU($tl['yOffset']));
            $objWriter->endElement();
            $objWriter->startElement('xdr:to');
                $objWriter->writeElement('xdr:col', \PhpOffice\PhpExcel\Cell::columnIndexFromString($br['colRow'][0]) - 1);
                $objWriter->writeElement('xdr:colOff', \PhpOffice\PhpExcel\Shared\Drawing::pixelsToEMU($br['xOffset']));
                $objWriter->writeElement('xdr:row', $br['colRow'][1] - 1);
                $objWriter->writeElement('xdr:rowOff', \PhpOffice\PhpExcel\Shared\Drawing::pixelsToEMU($br['yOffset']));
            $objWriter->endElement();

            $objWriter->startElement('xdr:graphicFrame');
                $objWriter->writeAttribute('macro', '');
                $objWriter->startElement('xdr:nvGraphicFramePr');
                    $objWriter->startElement('xdr:cNvPr');
                        $objWriter->writeAttribute('name', 'Chart '.$pRelationId);
                        $objWriter->writeAttribute('id', 1025 * $pRelationId);
                    $objWriter->endElement();
                    $objWriter->startElement('xdr:cNvGraphicFramePr');
                        $objWriter->startElement('a:graphicFrameLocks');
                        $objWriter->endElement();
                    $objWriter->endElement();
                $objWriter->endElement();

                $objWriter->startElement('xdr:xfrm');
                    $objWriter->startElement('a:off');
                        $objWriter->writeAttribute('x', '0');
                        $objWriter->writeAttribute('y', '0');
                    $objWriter->endElement();
                    $objWriter->startElement('a:ext');
                        $objWriter->writeAttribute('cx', '0');
                        $objWriter->writeAttribute('cy', '0');
                    $objWriter->endElement();
                $objWriter->endElement();

                $objWriter->startElement('a:graphic');
                    $objWriter->startElement('a:graphicData');
                        $objWriter->writeAttribute('uri', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
                        $objWriter->startElement('c:chart');
                            $objWriter->writeAttribute('xmlns:c', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
                            $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
                            $objWriter->writeAttribute('r:id', 'rId'.$pRelationId);
                        $objWriter->endElement();
                    $objWriter->endElement();
                $objWriter->endElement();
            $objWriter->endElement();

            $objWriter->startElement('xdr:clientData');
            $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write drawings to XML format
     *
     * @param     \PhpOffice\PhpExcel\Shared\XMLWriter            $objWriter         XML Writer
     * @param     \PhpOffice\PhpExcel\Worksheet\BaseDrawing        $pDrawing
     * @param     int                                    $pRelationId
     * @throws     \PhpOffice\PhpExcel\Writer\Exception
     */
    public function writeDrawing(\PhpOffice\PhpExcel\Shared\XMLWriter $objWriter = null, \PhpOffice\PhpExcel\Worksheet\BaseDrawing $pDrawing = null, $pRelationId = -1)
    {
        if ($pRelationId >= 0) {
            // xdr:oneCellAnchor
            $objWriter->startElement('xdr:oneCellAnchor');
            // Image location
            $aCoordinates         = \PhpOffice\PhpExcel\Cell::coordinateFromString($pDrawing->getCoordinates());
            $aCoordinates[0]     = \PhpOffice\PhpExcel\Cell::columnIndexFromString($aCoordinates[0]);

            // xdr:from
            $objWriter->startElement('xdr:from');
            $objWriter->writeElement('xdr:col', $aCoordinates[0] - 1);
            $objWriter->writeElement('xdr:colOff', \PhpOffice\PhpExcel\Shared\Drawing::pixelsToEMU($pDrawing->getOffsetX()));
            $objWriter->writeElement('xdr:row', $aCoordinates[1] - 1);
            $objWriter->writeElement('xdr:rowOff', \PhpOffice\PhpExcel\Shared\Drawing::pixelsToEMU($pDrawing->getOffsetY()));
            $objWriter->endElement();

            // xdr:ext
            $objWriter->startElement('xdr:ext');
            $objWriter->writeAttribute('cx', \PhpOffice\PhpExcel\Shared\Drawing::pixelsToEMU($pDrawing->getWidth()));
            $objWriter->writeAttribute('cy', \PhpOffice\PhpExcel\Shared\Drawing::pixelsToEMU($pDrawing->getHeight()));
            $objWriter->endElement();

            // xdr:pic
            $objWriter->startElement('xdr:pic');

            // xdr:nvPicPr
            $objWriter->startElement('xdr:nvPicPr');

            // xdr:cNvPr
            $objWriter->startElement('xdr:cNvPr');
            $objWriter->writeAttribute('id', $pRelationId);
            $objWriter->writeAttribute('name', $pDrawing->getName());
            $objWriter->writeAttribute('descr', $pDrawing->getDescription());
            $objWriter->endElement();

            // xdr:cNvPicPr
            $objWriter->startElement('xdr:cNvPicPr');

            // a:picLocks
            $objWriter->startElement('a:picLocks');
            $objWriter->writeAttribute('noChangeAspect', '1');
            $objWriter->endElement();

            $objWriter->endElement();

            $objWriter->endElement();

            // xdr:blipFill
            $objWriter->startElement('xdr:blipFill');

            // a:blip
            $objWriter->startElement('a:blip');
            $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
            $objWriter->writeAttribute('r:embed', 'rId' . $pRelationId);
            $objWriter->endElement();

            // a:stretch
            $objWriter->startElement('a:stretch');
                $objWriter->writeElement('a:fillRect', null);
            $objWriter->endElement();

            $objWriter->endElement();

            // xdr:spPr
            $objWriter->startElement('xdr:spPr');

            // a:xfrm
            $objWriter->startElement('a:xfrm');
            $objWriter->writeAttribute('rot', \PhpOffice\PhpExcel\Shared\Drawing::degreesToAngle($pDrawing->getRotation()));
            $objWriter->endElement();

            // a:prstGeom
            $objWriter->startElement('a:prstGeom');
            $objWriter->writeAttribute('prst', 'rect');

            // a:avLst
            $objWriter->writeElement('a:avLst', null);

            $objWriter->endElement();

//                        // a:solidFill
//                        $objWriter->startElement('a:solidFill');

//                            // a:srgbClr
//                            $objWriter->startElement('a:srgbClr');
//                            $objWriter->writeAttribute('val', 'FFFFFF');

///* SHADE
//                                // a:shade
//                                $objWriter->startElement('a:shade');
//                                $objWriter->writeAttribute('val', '85000');
//                                $objWriter->endElement();
//*/

//                            $objWriter->endElement();

//                        $objWriter->endElement();
/*
            // a:ln
            $objWriter->startElement('a:ln');
            $objWriter->writeAttribute('w', '88900');
            $objWriter->writeAttribute('cap', 'sq');

                // a:solidFill
                $objWriter->startElement('a:solidFill');

                    // a:srgbClr
                    $objWriter->startElement('a:srgbClr');
                    $objWriter->writeAttribute('val', 'FFFFFF');
                    $objWriter->endElement();

                $objWriter->endElement();

                // a:miter
                $objWriter->startElement('a:miter');
                $objWriter->writeAttribute('lim', '800000');
                $objWriter->endElement();

            $objWriter->endElement();
*/

            if ($pDrawing->getShadow()->getVisible()) {
                // a:effectLst
                $objWriter->startElement('a:effectLst');

                // a:outerShdw
                $objWriter->startElement('a:outerShdw');
                $objWriter->writeAttribute('blurRad', \PhpOffice\PhpExcel\Shared\Drawing::pixelsToEMU($pDrawing->getShadow()->getBlurRadius()));
                $objWriter->writeAttribute('dist', \PhpOffice\PhpExcel\Shared\Drawing::pixelsToEMU($pDrawing->getShadow()->getDistance()));
                $objWriter->writeAttribute('dir', \PhpOffice\PhpExcel\Shared\Drawing::degreesToAngle($pDrawing->getShadow()->getDirection()));
                $objWriter->writeAttribute('algn', $pDrawing->getShadow()->getAlignment());
                $objWriter->writeAttribute('rotWithShape', '0');

                // a:srgbClr
                $objWriter->startElement('a:srgbClr');
                $objWriter->writeAttribute('val', $pDrawing->getShadow()->getColor()->getRGB());

                // a:alpha
                $objWriter->startElement('a:alpha');
                $objWriter->writeAttribute('val', $pDrawing->getShadow()->getAlpha() * 1000);
                $objWriter->endElement();

                $objWriter->endElement();

                $objWriter->endElement();

                $objWriter->endElement();
            }
/*

                // a:scene3d
                $objWriter->startElement('a:scene3d');

                    // a:camera
                    $objWriter->startElement('a:camera');
                    $objWriter->writeAttribute('prst', 'orthographicFront');
                    $objWriter->endElement();

                    // a:lightRig
                    $objWriter->startElement('a:lightRig');
                    $objWriter->writeAttribute('rig', 'twoPt');
                    $objWriter->writeAttribute('dir', 't');

                        // a:rot
                        $objWriter->startElement('a:rot');
                        $objWriter->writeAttribute('lat', '0');
                        $objWriter->writeAttribute('lon', '0');
                        $objWriter->writeAttribute('rev', '0');
                        $objWriter->endElement();

                    $objWriter->endElement();

                $objWriter->endElement();
*/
/*
                // a:sp3d
                $objWriter->startElement('a:sp3d');

                    // a:bevelT
                    $objWriter->startElement('a:bevelT');
                    $objWriter->writeAttribute('w', '25400');
                    $objWriter->writeAttribute('h', '19050');
                    $objWriter->endElement();

                    // a:contourClr
                    $objWriter->startElement('a:contourClr');

                        // a:srgbClr
                        $objWriter->startElement('a:srgbClr');
                        $objWriter->writeAttribute('val', 'FFFFFF');
                        $objWriter->endElement();

                    $objWriter->endElement();

                $objWriter->endElement();
*/
            $objWriter->endElement();

            $objWriter->endElement();

            // xdr:clientData
            $objWriter->writeElement('xdr:clientData', null);

            $objWriter->endElement();
        } else {
            throw new \PhpOffice\PhpExcel\Writer\Exception("Invalid parameters passed.");
        }
    }

    /**
     * Write VML header/footer images to XML format
     *
     * @param     \PhpOffice\PhpExcel\Worksheet                $pWorksheet
     * @return string                          XML Output
     * @throws     \PhpOffice\PhpExcel\Writer\Exception
     */
    public function writeVMLHeaderFooterImages(\PhpOffice\PhpExcel\Worksheet $pWorksheet = null)
    {
        // Create XML writer
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new \PhpOffice\PhpExcel\Shared\XMLWriter(\PhpOffice\PhpExcel\Shared\XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new \PhpOffice\PhpExcel\Shared\XMLWriter(\PhpOffice\PhpExcel\Shared\XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Header/footer images
        $images = $pWorksheet->getHeaderFooter()->getImages();

        // xml
        $objWriter->startElement('xml');
        $objWriter->writeAttribute('xmlns:v', 'urn:schemas-microsoft-com:vml');
        $objWriter->writeAttribute('xmlns:o', 'urn:schemas-microsoft-com:office:office');
        $objWriter->writeAttribute('xmlns:x', 'urn:schemas-microsoft-com:office:excel');

        // o:shapelayout
        $objWriter->startElement('o:shapelayout');
        $objWriter->writeAttribute('v:ext', 'edit');

        // o:idmap
        $objWriter->startElement('o:idmap');
        $objWriter->writeAttribute('v:ext', 'edit');
        $objWriter->writeAttribute('data', '1');
        $objWriter->endElement();

        $objWriter->endElement();

        // v:shapetype
        $objWriter->startElement('v:shapetype');
        $objWriter->writeAttribute('id', '_x0000_t75');
        $objWriter->writeAttribute('coordsize', '21600,21600');
        $objWriter->writeAttribute('o:spt', '75');
        $objWriter->writeAttribute('o:preferrelative', 't');
        $objWriter->writeAttribute('path', 'm@4@5l@4@11@9@11@9@5xe');
        $objWriter->writeAttribute('filled', 'f');
        $objWriter->writeAttribute('stroked', 'f');

        // v:stroke
        $objWriter->startElement('v:stroke');
        $objWriter->writeAttribute('joinstyle', 'miter');
        $objWriter->endElement();

        // v:formulas
        $objWriter->startElement('v:formulas');

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'if lineDrawn pixelLineWidth 0');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'sum @0 1 0');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'sum 0 0 @1');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'prod @2 1 2');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'prod @3 21600 pixelWidth');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'prod @3 21600 pixelHeight');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'sum @0 0 1');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'prod @6 1 2');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'prod @7 21600 pixelWidth');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'sum @8 21600 0');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'prod @7 21600 pixelHeight');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'sum @10 21600 0');
        $objWriter->endElement();

        $objWriter->endElement();

        // v:path
        $objWriter->startElement('v:path');
        $objWriter->writeAttribute('o:extrusionok', 'f');
        $objWriter->writeAttribute('gradientshapeok', 't');
        $objWriter->writeAttribute('o:connecttype', 'rect');
        $objWriter->endElement();

        // o:lock
        $objWriter->startElement('o:lock');
        $objWriter->writeAttribute('v:ext', 'edit');
        $objWriter->writeAttribute('aspectratio', 't');
        $objWriter->endElement();

        $objWriter->endElement();

        // Loop through images
        foreach ($images as $key => $value) {
            $this->writeVMLHeaderFooterImage($objWriter, $key, $value);
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write VML comment to XML format
     *
     * @param     \PhpOffice\PhpExcel\Shared\XMLWriter        $objWriter             XML Writer
     * @param    string                            $pReference            Reference
     * @param     \PhpOffice\PhpExcel\Worksheet\HeaderFooterDrawing    $pImage        Image
     * @throws     \PhpOffice\PhpExcel\Writer\Exception
     */
    private function writeVMLHeaderFooterImage(\PhpOffice\PhpExcel\Shared\XMLWriter $objWriter = null, $pReference = '', \PhpOffice\PhpExcel\Worksheet\HeaderFooterDrawing $pImage = null)
    {
        // Calculate object id
        preg_match('{(\d+)}', md5($pReference), $m);
        $id = 1500 + (substr($m[1], 0, 2) * 1);

        // Calculate offset
        $width = $pImage->getWidth();
        $height = $pImage->getHeight();
        $marginLeft = $pImage->getOffsetX();
        $marginTop = $pImage->getOffsetY();

        // v:shape
        $objWriter->startElement('v:shape');
        $objWriter->writeAttribute('id', $pReference);
        $objWriter->writeAttribute('o:spid', '_x0000_s' . $id);
        $objWriter->writeAttribute('type', '#_x0000_t75');
        $objWriter->writeAttribute('style', "position:absolute;margin-left:{$marginLeft}px;margin-top:{$marginTop}px;width:{$width}px;height:{$height}px;z-index:1");

        // v:imagedata
        $objWriter->startElement('v:imagedata');
        $objWriter->writeAttribute('o:relid', 'rId' . $pReference);
        $objWriter->writeAttribute('o:title', $pImage->getName());
        $objWriter->endElement();

        // o:lock
        $objWriter->startElement('o:lock');
        $objWriter->writeAttribute('v:ext', 'edit');
        $objWriter->writeAttribute('rotation', 't');
        $objWriter->endElement();

        $objWriter->endElement();
    }


    /**
     * Get an array of all drawings
     *
     * @param     PHPExcel                            $pPHPExcel
     * @return     \PhpOffice\PhpExcel\Worksheet\Drawing[]        All drawings in PHPExcel
     * @throws     \PhpOffice\PhpExcel\Writer\Exception
     */
    public function allDrawings(\PhpOffice\PhpExcel\SpreadSheet $pPHPExcel = null)
    {
        // Get an array of all drawings
        $aDrawings    = array();

        // Loop through PHPExcel
        $sheetCount = $pPHPExcel->getSheetCount();
        for ($i = 0; $i < $sheetCount; ++$i) {
            // Loop through images and add to array
            $iterator = $pPHPExcel->getSheet($i)->getDrawingCollection()->getIterator();
            while ($iterator->valid()) {
                $aDrawings[] = $iterator->current();

                  $iterator->next();
            }
        }

        return $aDrawings;
    }
}
