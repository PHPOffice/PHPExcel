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

require_once(PHPEXCEL_ROOT . 'PHPExcel/Shared/trend/bestFitClass.php');

/**
 * PHPExcel_Linear_Best_Fit
 *
 * @category   PHPExcel
 * @package    PHPExcel_Shared_Trend
 * @copyright  Copyright (c) 2006 - 2012 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Linear_Best_Fit extends PHPExcel_Best_Fit
{
    /**
     * Algorithm type to use for best-fit
     * (Name of this trend class)
     *
     * @var	string
     **/
    protected $_bestFitType		= 'linear';

    /**
     * Return the Y-Value for a specified value of X
     *
     * @param  float $xValue X-Value
     * @return float Y-Value
     **/
    public function getValueOfYForX($xValue)
    {
        return $this->getIntersect() + $this->getSlope() * $xValue;
    }	//	function getValueOfYForX()

    /**
     * Return the X-Value for a specified value of Y
     *
     * @param  float $yValue Y-Value
     * @return float X-Value
     **/
    public function getValueOfXForY($yValue)
    {
        return ($yValue - $this->getIntersect()) / $this->getSlope();
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

        return 'Y = '.$intersect.' + '.$slope.' * X';
    }	//	function getEquation()

    /**
     * Execute the regression and calculate the goodness of fit for a set of X and Y data values
     *
     * @param float[] $yValues The set of Y-values for this regression
     * @param float[] $xValues The set of X-values for this regression
     * @param boolean $const
     */
    private function _linear_regression($yValues, $xValues, $const)
    {
        $this->_leastSquareFit($yValues, $xValues,$const);
    }	//	function _linear_regression()

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
            $this->_linear_regression($yValues, $xValues, $const);
        }
    }	//	function __construct()

}	//	class linearBestFit
