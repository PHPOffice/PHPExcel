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
 * @category    PHPExcel
 * @package        PHPExcel\Chart
 * @copyright    Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license        http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version        ##VERSION##, ##DATE##
 */


namespace PHPExcel;

/**
 * PHPExcel\Chart
 *
 * @category    PHPExcel
 * @package        PHPExcel\Chart
 * @copyright    Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class Chart
{
    /**
     * Chart Name
     *
     * @var string
     */
    protected $_name = '';

    /**
     * Worksheet
     *
     * @var PHPExcel\Worksheet
     */
    protected $_worksheet = null;

    /**
     * Chart Title
     *
     * @var PHPExcel\Chart_Title
     */
    protected $_title = null;

    /**
     * Chart Legend
     *
     * @var PHPExcel\Chart_Legend
     */
    protected $_legend = null;

    /**
     * X-Axis Label
     *
     * @var PHPExcel\Chart_Title
     */
    protected $_xAxisLabel = null;

    /**
     * Y-Axis Label
     *
     * @var PHPExcel\Chart_Title
     */
    protected $_yAxisLabel = null;

    /**
     * Chart Plot Area
     *
     * @var PHPExcel\Chart_PlotArea
     */
    protected $_plotArea = null;

    /**
     * Plot Visible Only
     *
     * @var boolean
     */
    protected $_plotVisibleOnly = true;

    /**
     * Display Blanks as
     *
     * @var string
     */
    protected $_displayBlanksAs = '0';


    /**
     * Top-Left Cell Position
     *
     * @var string
     */
    protected $_topLeftCellRef = 'A1';


    /**
     * Top-Left X-Offset
     *
     * @var integer
     */
    protected $_topLeftXOffset = 0;


    /**
     * Top-Left Y-Offset
     *
     * @var integer
     */
    protected $_topLeftYOffset = 0;


    /**
     * Bottom-Right Cell Position
     *
     * @var string
     */
    protected $_bottomRightCellRef = 'A1';


    /**
     * Bottom-Right X-Offset
     *
     * @var integer
     */
    protected $_bottomRightXOffset = 10;


    /**
     * Bottom-Right Y-Offset
     *
     * @var integer
     */
    protected $_bottomRightYOffset = 10;


    /**
     * Create a new PHPExcel\Chart
     */
    public function __construct($name, Chart_Title $title = null, Chart_Legend $legend = null, Chart_PlotArea $plotArea = null, $plotVisibleOnly = true, $displayBlanksAs = '0', Chart_Title $xAxisLabel = null, Chart_Title $yAxisLabel = null)
    {
        $this->_name = $name;
        $this->_title = $title;
        $this->_legend = $legend;
        $this->_xAxisLabel = $xAxisLabel;
        $this->_yAxisLabel = $yAxisLabel;
        $this->_plotArea = $plotArea;
        $this->_plotVisibleOnly = $plotVisibleOnly;
        $this->_displayBlanksAs = $displayBlanksAs;
    }

    /**
     * Get Name
     *
     * @return string
     */
    public function getName() {
        return $this->_name;
    }

    /**
     * Get Worksheet
     *
     * @return PHPExcel\Worksheet
     */
    public function getWorksheet() {
        return $this->_worksheet;
    }

    /**
     * Set Worksheet
     *
     * @param    PHPExcel\Worksheet    $pValue
     * @throws    PHPExcel\Chart_Exception
     * @return PHPExcel\Chart
     */
    public function setWorksheet(Worksheet $pValue = null) {
        $this->_worksheet = $pValue;

        return $this;
    }

    /**
     * Get Title
     *
     * @return PHPExcel\Chart_Title
     */
    public function getTitle() {
        return $this->_title;
    }

    /**
     * Set Title
     *
     * @param    PHPExcel\Chart_Title $title
     * @return    PHPExcel\Chart
     */
    public function setTitle(Chart_Title $title) {
        $this->_title = $title;

        return $this;
    }

    /**
     * Get Legend
     *
     * @return PHPExcel\Chart_Legend
     */
    public function getLegend() {
        return $this->_legend;
    }

    /**
     * Set Legend
     *
     * @param    PHPExcel\Chart_Legend $legend
     * @return    PHPExcel\Chart
     */
    public function setLegend(Chart_Legend $legend) {
        $this->_legend = $legend;

        return $this;
    }

    /**
     * Get X-Axis Label
     *
     * @return PHPExcel\Chart_Title
     */
    public function getXAxisLabel() {
        return $this->_xAxisLabel;
    }

    /**
     * Set X-Axis Label
     *
     * @param    PHPExcel\Chart_Title $label
     * @return    PHPExcel\Chart
     */
    public function setXAxisLabel(Chart_Title $label) {
        $this->_xAxisLabel = $label;

        return $this;
    }

    /**
     * Get Y-Axis Label
     *
     * @return PHPExcel\Chart_Title
     */
    public function getYAxisLabel() {
        return $this->_yAxisLabel;
    }

    /**
     * Set Y-Axis Label
     *
     * @param    PHPExcel\Chart_Title $label
     * @return    PHPExcel\Chart
     */
    public function setYAxisLabel(Chart_Title $label) {
        $this->_yAxisLabel = $label;

        return $this;
    }

    /**
     * Get Plot Area
     *
     * @return PHPExcel\Chart_PlotArea
     */
    public function getPlotArea() {
        return $this->_plotArea;
    }

    /**
     * Get Plot Visible Only
     *
     * @return boolean
     */
    public function getPlotVisibleOnly() {
        return $this->_plotVisibleOnly;
    }

    /**
     * Set Plot Visible Only
     *
     * @param boolean $plotVisibleOnly
     * @return PHPExcel\Chart
     */
    public function setPlotVisibleOnly($plotVisibleOnly = true) {
        $this->_plotVisibleOnly = $plotVisibleOnly;

        return $this;
    }

    /**
     * Get Display Blanks as
     *
     * @return string
     */
    public function getDisplayBlanksAs() {
        return $this->_displayBlanksAs;
    }

    /**
     * Set Display Blanks as
     *
     * @param string $displayBlanksAs
     * @return PHPExcel\Chart
     */
    public function setDisplayBlanksAs($displayBlanksAs = '0') {
        $this->_displayBlanksAs = $displayBlanksAs;
    }


    /**
     * Set the Top Left position for the chart
     *
     * @param    string    $cell
     * @param    integer    $xOffset
     * @param    integer    $yOffset
     * @return PHPExcel\Chart
     */
    public function setTopLeftPosition($cell, $xOffset=null, $yOffset=null) {
        $this->_topLeftCellRef = $cell;
        if (!is_null($xOffset))
            $this->setTopLeftXOffset($xOffset);
        if (!is_null($yOffset))
            $this->setTopLeftYOffset($yOffset);

        return $this;
    }

    /**
     * Get the top left position of the chart
     *
     * @return array    an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
     */
    public function getTopLeftPosition() {
        return array( 'cell'    => $this->_topLeftCellRef,
                      'xOffset'    => $this->_topLeftXOffset,
                      'yOffset'    => $this->_topLeftYOffset
                    );
    }

    /**
     * Get the cell address where the top left of the chart is fixed
     *
     * @return string
     */
    public function getTopLeftCell() {
        return $this->_topLeftCellRef;
    }

    /**
     * Set the Top Left cell position for the chart
     *
     * @param    string    $cell
     * @return PHPExcel\Chart
     */
    public function setTopLeftCell($cell) {
        $this->_topLeftCellRef = $cell;

        return $this;
    }

    /**
     * Set the offset position within the Top Left cell for the chart
     *
     * @param    integer    $xOffset
     * @param    integer    $yOffset
     * @return PHPExcel\Chart
     */
    public function setTopLeftOffset($xOffset=null,$yOffset=null) {
        if (!is_null($xOffset))
            $this->setTopLeftXOffset($xOffset);
        if (!is_null($yOffset))
            $this->setTopLeftYOffset($yOffset);

        return $this;
    }

    /**
     * Get the offset position within the Top Left cell for the chart
     *
     * @return integer[]
     */
    public function getTopLeftOffset() {
        return array( 'X' => $this->_topLeftXOffset,
                      'Y' => $this->_topLeftYOffset
                    );
    }

    public function setTopLeftXOffset($xOffset) {
        $this->_topLeftXOffset = $xOffset;

        return $this;
    }

    public function getTopLeftXOffset() {
        return $this->_topLeftXOffset;
    }

    public function setTopLeftYOffset($yOffset) {
        $this->_topLeftYOffset = $yOffset;

        return $this;
    }

    public function getTopLeftYOffset() {
        return $this->_topLeftYOffset;
    }

    /**
     * Set the Bottom Right position of the chart
     *
     * @param    string    $cell
     * @param    integer    $xOffset
     * @param    integer    $yOffset
     * @return PHPExcel\Chart
     */
    public function setBottomRightPosition($cell, $xOffset=null, $yOffset=null) {
        $this->_bottomRightCellRef = $cell;
        if (!is_null($xOffset))
            $this->setBottomRightXOffset($xOffset);
        if (!is_null($yOffset))
            $this->setBottomRightYOffset($yOffset);

        return $this;
    }

    /**
     * Get the bottom right position of the chart
     *
     * @return array    an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
     */
    public function getBottomRightPosition() {
        return array( 'cell'    => $this->_bottomRightCellRef,
                      'xOffset'    => $this->_bottomRightXOffset,
                      'yOffset'    => $this->_bottomRightYOffset
                    );
    }

    public function setBottomRightCell($cell) {
        $this->_bottomRightCellRef = $cell;

        return $this;
    }

    /**
     * Get the cell address where the bottom right of the chart is fixed
     *
     * @return string
     */
    public function getBottomRightCell() {
        return $this->_bottomRightCellRef;
    }

    /**
     * Set the offset position within the Bottom Right cell for the chart
     *
     * @param    integer    $xOffset
     * @param    integer    $yOffset
     * @return PHPExcel\Chart
     */
    public function setBottomRightOffset($xOffset=null,$yOffset=null) {
        if (!is_null($xOffset))
            $this->setBottomRightXOffset($xOffset);
        if (!is_null($yOffset))
            $this->setBottomRightYOffset($yOffset);

        return $this;
    }

    /**
     * Get the offset position within the Bottom Right cell for the chart
     *
     * @return integer[]
     */
    public function getBottomRightOffset() {
        return array( 'X' => $this->_bottomRightXOffset,
                      'Y' => $this->_bottomRightYOffset
                    );
    }

    public function setBottomRightXOffset($xOffset) {
        $this->_bottomRightXOffset = $xOffset;

        return $this;
    }

    public function getBottomRightXOffset() {
        return $this->_bottomRightXOffset;
    }

    public function setBottomRightYOffset($yOffset) {
        $this->_bottomRightYOffset = $yOffset;

        return $this;
    }

    public function getBottomRightYOffset() {
        return $this->_bottomRightYOffset;
    }


    public function refresh() {
        if ($this->_worksheet !== null) {
            $this->_plotArea->refresh($this->_worksheet);
        }
    }

    public function render($outputDestination = null) {
        $libraryName = Settings::getChartRendererName();
        if (is_null($libraryName)) {
            return false;
        }
        //    Ensure that data series values are up-to-date before we render
        $this->refresh();

        $libraryPath = Settings::getChartRendererPath();
        $includePath = str_replace('\\','/',get_include_path());
        $rendererPath = str_replace('\\','/',$libraryPath);
        if (strpos($rendererPath,$includePath) === false) {
            set_include_path(get_include_path() . PATH_SEPARATOR . $libraryPath);
        }

        $rendererName = __NAMESPACE__ . '\Chart_Renderer_'.$libraryName;
        $renderer = new $rendererName($this);

        if ($outputDestination == 'php://output') {
            $outputDestination = null;
        }
        return $renderer->render($outputDestination);
    }
}
