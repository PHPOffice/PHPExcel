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
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 *
 * Based on:
 * SplClassLoader implementation that implements the technical interoperability
 * standards for PHP 5.3 namespaces and class names.
 *
 * http://groups.google.com/group/php-standards/web/final-proposal
 *
 *     // Example which loads classes for the Doctrine Common package in the
 *     // Doctrine\Common namespace.
 *     $classLoader = new SplClassLoader('Doctrine\Common', '/path/to/doctrine');
 *     $classLoader->register();
 *
 * @author Jonathan H. Wage <jonwage@gmail.com>
 * @author Roman S. Borschel <roman@code-factory.org>
 * @author Matthew Weier O'Phinney <matthew@zend.com>
 * @author Kris Wallsmith <kris.wallsmith@gmail.com>
 * @author Fabien Potencier <fabien.potencier@symfony-project.org>
 */


namespace PHPExcel;

/**
 * PHPExcel\Autoloader
 *
 * @category    PHPExcel
 * @package     PHPExcel
 * @copyright   Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class Autoloader
{
    private $fileExtension = '.php';
    private $namespace;
    private $includePath;
    private $namespaceSeparator = '\\';

    /**
     * Creates a new SplClassLoader that loads classes of the
     * specified namespace.
     *
     * @param string $namespace The namespace to use.
     * @param string $includePath The directory path for that namespace.
     */
    public function __construct($namespace = null, $includePath = null)
    {
        $this->namespace = $namespace;
        $this->includePath = $includePath;
    }

    /**
     * Sets the namespace separator used by classes in the namespace of this class loader.
     *
     * @param string $sep The separator to use.
     */
    public function setNamespaceSeparator($sep)
    {
        $this->namespaceSeparator = $sep;
    }

    /**
     * Gets the namespace seperator used by classes in the namespace of this class loader.
     *
     * @return void
     */
    public function getNamespaceSeparator()
    {
        return $this->namespaceSeparator;
    }

    /**
     * Sets the base include path for all class files in the namespace of this class loader.
     *
     * @param string $includePath
     */
    public function setIncludePath($includePath)
    {
        $this->includePath = $includePath;
    }

    /**
     * Gets the base include path for all class files in the namespace of this class loader.
     *
     * @return string $includePath
     */
    public function getIncludePath()
    {
        return $this->includePath;
    }

    /**
     * Sets the file extension of class files in the namespace of this class loader.
     *
     * @param string $fileExtension
     */
    public function setFileExtension($fileExtension)
    {
        $this->fileExtension = $fileExtension;
    }

    /**
     * Gets the file extension of class files in the namespace of this class loader.
     *
     * @return string $fileExtension
     */
    public function getFileExtension()
    {
        return $this->fileExtension;
    }

    /**
     * Installs this class loader on the SPL autoload stack.
     */
    public function register()
    {
        spl_autoload_register(array($this, 'loadClass'));
    }

    /**
     * Uninstalls this class loader from the SPL autoloader stack.
     */
    public function unregister()
    {
        spl_autoload_unregister(array($this, 'loadClass'));
    }

    /**
     * Loads the given class or interface.
     *
     * @param string $className The name of the class to load.
     * @return void
     */
    public function loadClass($className)
    {
        if (null === $this->namespace ||
            $this->namespace.$this->namespaceSeparator ===
                substr($className, 0, strlen($this->namespace.$this->namespaceSeparator))
            ) {
            //  $fileName = '';
            //  $namespace = '';
            //  if (false !== ($lastNsPos = strripos($className, $this->_namespaceSeparator))) {
            //      $namespace = substr($className, 0, $lastNsPos);
            //      $className = substr($className, $lastNsPos + 1);
            //      $fileName = str_replace(
            //          $this->_namespaceSeparator,
            //          DIRECTORY_SEPARATOR,
            //          $namespace
            //      ) . DIRECTORY_SEPARATOR;
            //  }
            $fileName = str_replace(array('\\','_'), DIRECTORY_SEPARATOR, $className) . $this->fileExtension;
echo 'Filename: '; var_dump($fileName);
$fullFilename = ($this->includePath !== null ? $this->includePath . DIRECTORY_SEPARATOR : '') . $fileName;
echo 'Full Filename: '; var_dump($fullFilename);
if (!file_exists($fullFilename)) {
    var_dump(debug_backtrace());
}
            require ($this->includePath !== null ? $this->includePath . DIRECTORY_SEPARATOR : '') . $fileName;
        }
    }
}
