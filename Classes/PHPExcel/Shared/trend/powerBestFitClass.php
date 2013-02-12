<?php
/**
 * PHPExcel
 *
 * Copyright (c) 2006 - 2012 PHPExcel
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
 * @package    PHPExcel_Shared_Trend
 * @copyright  Copyright (c) 2006 - 2012 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */

require_once PHPEXCEL_ROOT . 'PHPExcel/Shared/trend/bestFitClass.php';

/**
 * PHPExcel_Power_Best_Fit
 *
 * @category   PHPExcel
 * @package    PHPExcel_Shared_Trend
 * @copyright  Copyright (c) 2006 - 2012 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Power_Best_Fit extends PHPExcel_Best_Fit
{
    /**
     * Algorithm type to use for best-fit
     * (Name of this trend class)
     *
     * @var	string
     **/
    protected $_bestFitType		= 'power';

    /**
     * Return the Y-Value for a specified value of X
     *
     * @param  float $xValue X-Value
     * @return float Y-Value
     **/
    public function getValueOfYForX($xValue)
    {
        return $this->getIntersect() * pow(($xValue - $this->_Xoffset),$this->getSlope());
    }	//	function getValueOfYForX()

    /**
     * Return the X-Value for a specified value of Y
     *
     * @param  float $yValue Y-Value
     * @return float X-Value
     **/
    public function getValueOfXForY($yValue)
    {
        return pow((($yValue + $this->_Yoffset) / $this->getIntersect()),(1 / $this->getSlope()));
    }	//	function getValueOfXForY()

    /**
     * Return the Equation of the best-fit line
     *
     * @param  int    $dp Number of places of decimal precision to display
     * @return string
     **/
    public function getEquation($dp=0)
    {
        $slope = $this->getSlope($dp);
        $intersect = $this->getIntersect($dp);

        return 'Y = '.$intersect.' * X^'.$slope;
    }	//	function getEquation()

    /**
     * Return the Value of X where it intersects Y = 0
     *
     * @param  int    $dp Number of places of decimal precision to display
     * @return string
     **/
    public function getIntersect($dp=0)
    {
        if ($dp != 0) {
            return round(exp($this->_intersect),$dp);
        }

        return exp($this->_intersect);
    }	//	function getIntersect()

    /**
     * Execute the regression and calculate the goodness of fit for a set of X and Y data values
     *
     * @param float[] $yValues The set of Y-values for this regression
     * @param float[] $xValues The set of X-values for this regression
     * @param boolean $const
     */
    private function _power_regression($yValues, $xValues, $const)
    {
        foreach ($xValues as &$value) {
            if ($value < 0.0) {
                $value = 0 - log(abs($value));
            } elseif ($value > 0.0) {
                $value = log($value);
            }
        }
        unset($value);
        foreach ($yValues as &$value) {
            if ($value < 0.0) {
                $value = 0 - log(abs($value));
            } elseif ($value > 0.0) {
                $value = log($value);
            }
        }
        unset($value);

        $this->_leastSquareFit($yValues, $xValues, $const);
    }	//	function _power_regression()

    /**
     * Define the regression and calculate the goodness of fit for a set of X and Y data values
     *
     * @param float[] $yValues The set of Y-values for this regression
     * @param float[] $xValues The set of X-values for this regression
     * @param boolean $const
     */
    function __construct($yValues, $xValues=array(), $const=True)
    {
        if (parent::__construct($yValues, $xValues) !== False) {
            $this->_power_regression($yValues, $xValues, $const);
        }
    }	//	function __construct()

}	//	class powerBestFit
