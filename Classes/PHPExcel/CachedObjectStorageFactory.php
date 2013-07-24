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
 * @package    PHPExcel\CachedObjectStorage
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */


namespace PHPExcel;

/**
 * PHPExcel\CachedObjectStorageFactory
 *
 * @category    PHPExcel
 * @package     PHPExcel\CachedObjectStorage
 * @copyright   Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class CachedObjectStorageFactory
{
    const MEMORY             = 'Memory';
    const MEMORY_GZIP        = 'MemoryGZip';
    const MEMORY_SERIALIZED  = 'MemorySerialized';
    const IGBINARY           = 'Igbinary';
    const DISCISAM           = 'DiscISAM';
    const APC                = 'APC';
    const MEMCACHE           = 'Memcache';
    const PHPTEMP            = 'PHPTemp';
    const WINCACHE           = 'Wincache';
    const SQLITE             = 'SQLite';
    const SQLITE3            = 'SQLite3';


    /**
     * Name of the method used for cell cacheing
     *
     * @var string
     */
    protected static $cacheStorageMethod = null;

    /**
     * Name of the class used for cell cacheing
     *
     * @var string
     */
    protected static $cacheStorageClass = null;


    /**
     * List of all possible cache storage methods
     *
     * @var string[]
     */
    protected static $storageMethods = array(
        self::MEMORY,
        self::MEMORY_GZIP,
        self::MEMORY_SERIALIZED,
        self::IGBINARY,
        self::PHPTEMP,
        self::DISCISAM,
        self::APC,
        self::MEMCACHE,
        self::WINCACHE,
        self::SQLITE,
        self::SQLITE3,
    );


    /**
     * Default arguments for each cache storage method
     *
     * @var array of mixed array
     */
    protected static $storageMethodDefaultParameters = array(
        self::MEMORY => array(
        ),
        self::MEMORY_GZIP => array(
        ),
        self::MEMORY_SERIALIZED => array(
        ),
        self::IGBINARY => array(
        ),
        self::PHPTEMP => array( 'memoryCacheSize' => '1MB'
        ),
        self::DISCISAM => array(
            'dir'             => null
        ),
        self::APC => array(
            'cacheTime' => 600
        ),
        self::MEMCACHE => array(
            'memcacheServer' => 'localhost',
            'memcachePort' => 11211,
            'cacheTime' => 600
        ),
        self::WINCACHE => array(
            'cacheTime' => 600
        ),
        self::SQLITE => array(
        ),
        self::SQLITE3 => array(
        ),
    );


    /**
     * Arguments for the active cache storage method
     *
     * @var array of mixed array
     */
    protected static $storageMethodParameters = array();


    /**
     * Return the current cache storage method
     *
     * @return string|null
     **/
    public static function getCacheStorageMethod()
    {
        return self::$cacheStorageMethod;
    }   //    function getCacheStorageMethod()


    /**
     * Return the current cache storage class
     *
     * @return PHPExcel\CachedObjectStorage_ICache|null
     **/
    public static function getCacheStorageClass()
    {
        return self::$cacheStorageClass;
    }   //    function getCacheStorageClass()


    /**
     * Return the list of all possible cache storage methods
     *
     * @return string[]
     **/
    public static function getAllCacheStorageMethods()
    {
        return self::$storageMethods;
    }   //    function getCacheStorageMethods()


    /**
     * Return the list of all available cache storage methods
     *
     * @return string[]
     **/
    public static function getCacheStorageMethods()
    {
        $activeMethods = array();
        foreach (self::$storageMethods as $storageMethod) {
            $cacheStorageClass = 'PHPExcel\CachedObjectStorage_' . $storageMethod;
            if (call_user_func(array($cacheStorageClass, 'cacheMethodIsAvailable'))) {
                $activeMethods[] = $storageMethod;
            }
        }
        return $activeMethods;
    }   //    function getCacheStorageMethods()


    /**
     * Identify the cache storage method to use
     *
     * @param    string            $method       Name of the method to use for cell cacheing
     * @param    array of mixed    $arguments    Additional arguments to pass to the cell caching class
     *                                               when instantiating
     * @return boolean
     **/
    public static function initialize($method = self::MEMORY, $arguments = array())
    {
        if (!in_array($method, self::$storageMethods)) {
            return false;
        }

        $cacheStorageClass = __NAMESPACE__ . '\CachedObjectStorage_'.$method;
        if (!call_user_func(array($cacheStorageClass, 'cacheMethodIsAvailable'))) {
            return false;
        }

        self::$storageMethodParameters[$method] = self::$storageMethodDefaultParameters[$method];
        foreach($arguments as $k => $v) {
            if (array_key_exists($k, self::$storageMethodParameters[$method])) {
                self::$storageMethodParameters[$method][$k] = $v;
            }
        }

        if (self::$cacheStorageMethod === null) {
            self::$cacheStorageClass = 'CachedObjectStorage_' . $method;
            self::$cacheStorageMethod = $method;
        }
        return true;
    }   //    function initialize()


    /**
     * Initialise the cache storage
     *
     * @param    PHPExcel\Worksheet     $parent        Enable cell caching for this worksheet
     * @return    PHPExcel\CachedObjectStorage_ICache
     **/
    public static function getInstance(Worksheet $parent)
    {
        $cacheMethodIsAvailable = true;
        if (self::$cacheStorageMethod === null) {
            $cacheMethodIsAvailable = self::initialize();
        }

        if ($cacheMethodIsAvailable) {
            $cacheStorageClass = __NAMESPACE__ . '\\' . self::$cacheStorageClass;
            $instance = new $cacheStorageClass(
                $parent,
                self::$storageMethodParameters[self::$cacheStorageMethod]
            );
            if ($instance !== null) {
                return $instance;
            }
        }

        return false;
    }   //    function getInstance()


    /**
     * Clear the cache storage
     *
     **/
    public static function finalize()
    {
        self::$cacheStorageMethod = null;
        self::$cacheStorageClass = null;
        self::$storageMethodParameters = array();
    }
}
