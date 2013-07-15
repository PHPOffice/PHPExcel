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
 * @category   PHPExcel
 * @package    PHPExcel\Settings
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */


namespace PHPExcel;

class Settings
{
    /**    constants */
    /**    Available Zip library classes */
    const PCLZIP        = 'PHPExcel_Shared_ZipArchive';
    const ZIPARCHIVE    = 'ZipArchive';

    /**    Optional Chart Rendering libraries */
    const CHART_RENDERER_JPGRAPH    = 'jpgraph';

    /**    Optional PDF Rendering libraries */
    const PDF_RENDERER_TCPDF        = 'tcPDF';
    const PDF_RENDERER_DOMPDF        = 'DomPDF';
    const PDF_RENDERER_MPDF         = 'mPDF';


    protected static $_chartRenderers = array(
        self::CHART_RENDERER_JPGRAPH,
    );

    protected static $_pdfRenderers = array(
        self::PDF_RENDERER_TCPDF,
        self::PDF_RENDERER_DOMPDF,
        self::PDF_RENDERER_MPDF,
    );


    /**
     * Name of the class used for Zip file management
     *    e.g.
     *        ZipArchive
     *
     * @var string
     */
    protected static $_zipClass    = self::ZIPARCHIVE;


    /**
     * Name of the external Library used for rendering charts
     *    e.g.
     *        jpgraph
     *
     * @var string
     */
    protected static $_chartRendererName = null;

    /**
     * Directory Path to the external Library used for rendering charts
     *
     * @var string
     */
    protected static $_chartRendererPath = null;


    /**
     * Name of the external Library used for rendering PDF files
     *    e.g.
     *         mPDF
     *
     * @var string
     */
    protected static $_pdfRendererName = null;

    /**
     * Directory Path to the external Library used for rendering PDF files
     *
     * @var string
     */
    protected static $_pdfRendererPath = null;


    /**
     * Set the Zip handler Class that PHPExcel should use for Zip file management (PCLZip or ZipArchive)
     *
     * @param string $zipClass    The Zip handler class that PHPExcel should use for Zip file management
     *      e.g. PHPExcel\Settings::PCLZip or PHPExcel\Settings::ZipArchive
     * @return    boolean    Success or failure
     */
    public static function setZipClass($zipClass)
    {
        if (($zipClass === self::PCLZIP) ||
            ($zipClass === self::ZIPARCHIVE)) {
            self::$_zipClass = $zipClass;
            return true;
        }
        return false;
    }


    /**
     * Return the name of the Zip handler Class that PHPExcel is configured to use (PCLZip or ZipArchive)
     *    or Zip file management
     *
     * @return string Name of the Zip handler Class that PHPExcel is configured to use
     *    for Zip file management
     *    e.g. PHPExcel\Settings::PCLZip or PHPExcel\Settings::ZipArchive
     */
    public static function getZipClass()
    {
        return self::$_zipClass;
    }


    /**
     * Return the name of the method that is currently configured for cell cacheing
     *
     * @return string Name of the cacheing method
     */
    public static function getCacheStorageMethod()
    {
        return CachedObjectStorageFactory::getCacheStorageMethod();
    }


    /**
     * Return the name of the class that is currently being used for cell cacheing
     *
     * @return string Name of the class currently being used for cacheing
     */
    public static function getCacheStorageClass()
    {
        return CachedObjectStorageFactory::getCacheStorageClass();
    }


    /**
     * Set the method that should be used for cell cacheing
     *
     * @param string $method Name of the cacheing method
     * @param array $arguments Optional configuration arguments for the cacheing method
     * @return boolean Success or failure
     */
    public static function setCacheStorageMethod(
        $method = CachedObjectStorageFactory::MEMORY,
        $arguments = array()
    )
    {
        return CachedObjectStorageFactory::initialize($method, $arguments);
    }


    /**
     * Set the locale code to use for formula translations and any special formatting
     *
     * @param string $locale The locale code to use (e.g. "fr" or "pt_br" or "en_uk")
     * @return boolean Success or failure
     */
    public static function setLocale($locale='en_us')
    {
        return Calculation::getInstance()->setLocale($locale);
    }


    /**
     * Set details of the external library that PHPExcel should use for rendering charts
     *
     * @param string $libraryName    Internal reference name of the library
     *    e.g. PHPExcel\Settings::CHART_RENDERER_JPGRAPH
     * @param string $libraryBaseDir Directory path to the library's base folder
     *
     * @return    boolean    Success or failure
     */
    public static function setChartRenderer($libraryName, $libraryBaseDir)
    {
        if (!self::setChartRendererName($libraryName))
            return false;
        return self::setChartRendererPath($libraryBaseDir);
    }


    /**
     * Identify to PHPExcel the external library to use for rendering charts
     *
     * @param string $libraryName    Internal reference name of the library
     *    e.g. PHPExcel\Settings::CHART_RENDERER_JPGRAPH
     *
     * @return    boolean    Success or failure
     */
    public static function setChartRendererName($libraryName)
    {
        if (!in_array($libraryName,self::$_chartRenderers)) {
            return false;
        }

        self::$_chartRendererName = $libraryName;

        return true;
    }


    /**
     * Tell PHPExcel where to find the external library to use for rendering charts
     *
     * @param string $libraryBaseDir    Directory path to the library's base folder
     * @return    boolean    Success or failure
     */
    public static function setChartRendererPath($libraryBaseDir)
    {
        if ((file_exists($libraryBaseDir) === false) || (is_readable($libraryBaseDir) === false)) {
            return false;
        }
        self::$_chartRendererPath = $libraryBaseDir;

        return true;
    }


    /**
     * Return the Chart Rendering Library that PHPExcel is currently configured to use (e.g. jpgraph)
     *
     * @return string|null Internal reference name of the Chart Rendering Library that PHPExcel is
     *    currently configured to use
     *    e.g. PHPExcel\Settings::CHART_RENDERER_JPGRAPH
     */
    public static function getChartRendererName()
    {
        return self::$_chartRendererName;
    }


    /**
     * Return the directory path to the Chart Rendering Library that PHPExcel is currently configured to use
     *
     * @return string|null Directory Path to the Chart Rendering Library that PHPExcel is
     *     currently configured to use
     */
    public static function getChartRendererPath()
    {
        return self::$_chartRendererPath;
    }


    /**
     * Set details of the external library that PHPExcel should use for rendering PDF files
     *
     * @param string $libraryName Internal reference name of the library
     *     e.g. PHPExcel\Settings::PDF_RENDERER_TCPDF,
     *     PHPExcel\Settings::PDF_RENDERER_DOMPDF
     *  or PHPExcel\Settings::PDF_RENDERER_MPDF
     * @param string $libraryBaseDir Directory path to the library's base folder
     *
     * @return boolean Success or failure
     */
    public static function setPdfRenderer($libraryName, $libraryBaseDir)
    {
        if (!self::setPdfRendererName($libraryName))
            return false;
        return self::setPdfRendererPath($libraryBaseDir);
    }


    /**
     * Identify to PHPExcel the external library to use for rendering PDF files
     *
     * @param string $libraryName Internal reference name of the library
     *     e.g. PHPExcel\Settings::PDF_RENDERER_TCPDF,
     *    PHPExcel\Settings::PDF_RENDERER_DOMPDF
     *     or PHPExcel\Settings::PDF_RENDERER_MPDF
     *
     * @return boolean Success or failure
     */
    public static function setPdfRendererName($libraryName)
    {
        if (!in_array($libraryName,self::$_pdfRenderers)) {
            return false;
        }

        self::$_pdfRendererName = $libraryName;

        return true;
    }


    /**
     * Tell PHPExcel where to find the external library to use for rendering PDF files
     *
     * @param string $libraryBaseDir Directory path to the library's base folder
     * @return boolean Success or failure
     */
    public static function setPdfRendererPath($libraryBaseDir)
    {
        if ((file_exists($libraryBaseDir) === false) || (is_readable($libraryBaseDir) === false)) {
            return false;
        }
        self::$_pdfRendererPath = $libraryBaseDir;

        return true;
    }


    /**
     * Return the PDF Rendering Library that PHPExcel is currently configured to use (e.g. dompdf)
     *
     * @return string|null Internal reference name of the PDF Rendering Library that PHPExcel is
     *     currently configured to use
     *  e.g. PHPExcel\Settings::PDF_RENDERER_TCPDF,
     *  PHPExcel\Settings::PDF_RENDERER_DOMPDF
     *  or PHPExcel\Settings::PDF_RENDERER_MPDF
     */
    public static function getPdfRendererName()
    {
        return self::$_pdfRendererName;
    }


    /**
     * Return the directory path to the PDF Rendering Library that PHPExcel is currently configured to use
     *
     * @return string|null Directory Path to the PDF Rendering Library that PHPExcel is
     *        currently configured to use
     */
    public static function getPdfRendererPath()
    {
        return self::$_pdfRendererPath;
    }
}