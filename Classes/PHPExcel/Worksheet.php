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
 * @package    PHPExcel\Worksheet
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */


namespace PHPExcel;

/**
 * PHPExcel\Worksheet
 *
 * @category   PHPExcel
 * @package    PHPExcel\Worksheet
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class Worksheet implements IComparable
{
    /* Break types */
    const BREAK_NONE   = 0;
    const BREAK_ROW    = 1;
    const BREAK_COLUMN = 2;

    /* Sheet state */
    const SHEETSTATE_VISIBLE    = 'visible';
    const SHEETSTATE_HIDDEN     = 'hidden';
    const SHEETSTATE_VERYHIDDEN = 'veryHidden';

    /**
     * Invalid characters in sheet title
     *
     * @var array
     */
    protected static $_invalidCharacters = array('*', ':', '/', '\\', '?', '[', ']');

    /**
     * Parent spreadsheet
     *
     * @var PHPExcel
     */
    protected $_parent;

    /**
     * Cacheable collection of cells
     *
     * @var PHPExcel\CachedObjectStorage_xxx
     */
    protected $_cellCollection = null;

    /**
     * Collection of row dimensions
     *
     * @var PHPExcel\Worksheet_RowDimension[]
     */
    protected $_rowDimensions = array();

    /**
     * Default row dimension
     *
     * @var PHPExcel\Worksheet_RowDimension
     */
    protected $_defaultRowDimension = null;

    /**
     * Collection of column dimensions
     *
     * @var PHPExcel\Worksheet_ColumnDimension[]
     */
    protected $_columnDimensions = array();

    /**
     * Default column dimension
     *
     * @var PHPExcel\Worksheet_ColumnDimension
     */
    protected $_defaultColumnDimension = null;

    /**
     * Collection of drawings
     *
     * @var PHPExcel\Worksheet_BaseDrawing[]
     */
    protected $_drawingCollection = null;

    /**
     * Collection of Chart objects
     *
     * @var PHPExcel\Chart[]
     */
    protected $_chartCollection = array();

    /**
     * Worksheet title
     *
     * @var string
     */
    protected $_title;

    /**
     * Sheet state
     *
     * @var string
     */
    protected $_sheetState;

    /**
     * Page setup
     *
     * @var PHPExcel\Worksheet_PageSetup
     */
    protected $_pageSetup;

    /**
     * Page margins
     *
     * @var PHPExcel\Worksheet_PageMargins
     */
    protected $_pageMargins;

    /**
     * Page header/footer
     *
     * @var PHPExcel\Worksheet_HeaderFooter
     */
    protected $_headerFooter;

    /**
     * Sheet view
     *
     * @var PHPExcel\Worksheet_SheetView
     */
    protected $_sheetView;

    /**
     * Protection
     *
     * @var PHPExcel\Worksheet_Protection
     */
    protected $_protection;

    /**
     * Collection of styles
     *
     * @var PHPExcel\Style[]
     */
    protected $_styles = array();

    /**
     * Conditional styles. Indexed by cell coordinate, e.g. 'A1'
     *
     * @var array
     */
    protected $_conditionalStylesCollection = array();

    /**
     * Is the current cell collection sorted already?
     *
     * @var boolean
     */
    protected $_cellCollectionIsSorted = false;

    /**
     * Collection of breaks
     *
     * @var array
     */
    protected $_breaks = array();

    /**
     * Collection of merged cell ranges
     *
     * @var array
     */
    protected $_mergeCells = array();

    /**
     * Collection of protected cell ranges
     *
     * @var array
     */
    protected $_protectedCells = array();

    /**
     * Autofilter Range and selection
     *
     * @var PHPExcel\Worksheet_AutoFilter
     */
    protected $_autoFilter = null;

    /**
     * Freeze pane
     *
     * @var string
     */
    protected $_freezePane = '';

    /**
     * Show gridlines?
     *
     * @var boolean
     */
    protected $_showGridlines = true;

    /**
    * Print gridlines?
    *
    * @var boolean
    */
    protected $_printGridlines = false;

    /**
    * Show row and column headers?
    *
    * @var boolean
    */
    protected $_showRowColHeaders = true;

    /**
     * Show summary below? (Row/Column outline)
     *
     * @var boolean
     */
    protected $_showSummaryBelow = true;

    /**
     * Show summary right? (Row/Column outline)
     *
     * @var boolean
     */
    protected $_showSummaryRight = true;

    /**
     * Collection of comments
     *
     * @var PHPExcel\Comment[]
     */
    protected $_comments = array();

    /**
     * Active cell. (Only one!)
     *
     * @var string
     */
    protected $_activeCell = 'A1';

    /**
     * Selected cells
     *
     * @var string
     */
    protected $_selectedCells = 'A1';

    /**
     * Cached highest column
     *
     * @var string
     */
    protected $_cachedHighestColumn = 'A';

    /**
     * Cached highest row
     *
     * @var int
     */
    protected $_cachedHighestRow = 1;

    /**
     * Right-to-left?
     *
     * @var boolean
     */
    protected $_rightToLeft = false;

    /**
     * Hyperlinks. Indexed by cell coordinate, e.g. 'A1'
     *
     * @var array
     */
    protected $_hyperlinkCollection = array();

    /**
     * Data validation objects. Indexed by cell coordinate, e.g. 'A1'
     *
     * @var array
     */
    protected $_dataValidationCollection = array();

    /**
     * Tab color
     *
     * @var PHPExcel\Style_Color
     */
    protected $_tabColor;

    /**
     * Dirty flag
     *
     * @var boolean
     */
    protected $_dirty    = true;

    /**
     * Hash
     *
     * @var string
     */
    protected $_hash    = null;

    /**
     * Create a new worksheet
     *
     * @param PHPExcel        $pParent
     * @param string        $pTitle
     */
    public function __construct(Workbook $pParent = null, $pTitle = 'Worksheet')
    {
        // Set parent and title
        $this->_parent = $pParent;
        $this->setTitle($pTitle, false);
        $this->setSheetState(Worksheet::SHEETSTATE_VISIBLE);

        $this->_cellCollection        = CachedObjectStorageFactory::getInstance($this);

        // Set page setup
        $this->_pageSetup            = new Worksheet_PageSetup();

        // Set page margins
        $this->_pageMargins         = new Worksheet_PageMargins();

        // Set page header/footer
        $this->_headerFooter        = new Worksheet_HeaderFooter();

        // Set sheet view
        $this->_sheetView            = new Worksheet_SheetView();

        // Drawing collection
        $this->_drawingCollection    = new \ArrayObject();

        // Chart collection
        $this->_chartCollection     = new \ArrayObject();

        // Protection
        $this->_protection            = new Worksheet_Protection();

        // Default row dimension
        $this->_defaultRowDimension = new Worksheet_RowDimension(null);

        // Default column dimension
        $this->_defaultColumnDimension    = new Worksheet_ColumnDimension(null);

        $this->_autoFilter            = new Worksheet_AutoFilter(null, $this);
    }


    /**
     * Disconnect all cells from this PHPExcel\Worksheet object,
     *    typically so that the worksheet object can be unset
     *
     */
    public function disconnectCells() {
        if ( $this->_cellCollection !== null){
            $this->_cellCollection->unsetWorksheetCells();
            $this->_cellCollection = null;
        }
        //    detach ourself from the workbook, so that it can then delete this worksheet successfully
        $this->_parent = null;
    }

    /**
     * Code to execute when this worksheet is unset()
     *
     */
    function __destruct() {
        Calculation::getInstance($this->_parent)
            ->clearCalculationCacheForWorksheet($this->_title);

        $this->disconnectCells();
    }

   /**
     * Return the cache controller for the cell collection
     *
     * @return PHPExcel\CachedObjectStorage_xxx
     */
    public function getCellCacheController() {
        return $this->_cellCollection;
    }    //    function getCellCacheController()


    /**
     * Get array of invalid characters for sheet title
     *
     * @return array
     */
    public static function getInvalidCharacters()
    {
        return self::$_invalidCharacters;
    }

    /**
     * Check sheet title for valid Excel syntax
     *
     * @param string $pValue The string to check
     * @return string The valid string
     * @throws PHPExcel\Exception
     */
    private static function _checkSheetTitle($pValue)
    {
        // Some of the printable ASCII characters are invalid:  * : / \ ? [ ]
        if (str_replace(self::$_invalidCharacters, '', $pValue) !== $pValue) {
            throw new Exception('Invalid character found in sheet title');
        }

        // Maximum 31 characters allowed for sheet title
        if (Shared_String::CountCharacters($pValue) > 31) {
            throw new Exception('Maximum 31 characters allowed in sheet title.');
        }

        return $pValue;
    }

    /**
     * Get collection of cells
     *
     * @param boolean $pSorted Also sort the cell collection?
     * @return PHPExcel\Cell[]
     */
    public function getCellCollection($pSorted = true)
    {
        if ($pSorted) {
            // Re-order cell collection
            return $this->sortCellCollection();
        }
        if ($this->_cellCollection !== null) {
            return $this->_cellCollection->getCellList();
        }
        return array();
    }

    /**
     * Sort collection of cells
     *
     * @return PHPExcel\Worksheet
     */
    public function sortCellCollection()
    {
        if ($this->_cellCollection !== null) {
            return $this->_cellCollection->getSortedCellList();
        }
        return array();
    }

    /**
     * Get collection of row dimensions
     *
     * @return PHPExcel\Worksheet_RowDimension[]
     */
    public function getRowDimensions()
    {
        return $this->_rowDimensions;
    }

    /**
     * Get default row dimension
     *
     * @return PHPExcel\Worksheet_RowDimension
     */
    public function getDefaultRowDimension()
    {
        return $this->_defaultRowDimension;
    }

    /**
     * Get collection of column dimensions
     *
     * @return PHPExcel\Worksheet_ColumnDimension[]
     */
    public function getColumnDimensions()
    {
        return $this->_columnDimensions;
    }

    /**
     * Get default column dimension
     *
     * @return PHPExcel\Worksheet_ColumnDimension
     */
    public function getDefaultColumnDimension()
    {
        return $this->_defaultColumnDimension;
    }

    /**
     * Get collection of drawings
     *
     * @return PHPExcel\Worksheet_BaseDrawing[]
     */
    public function getDrawingCollection()
    {
        return $this->_drawingCollection;
    }

    /**
     * Get collection of charts
     *
     * @return PHPExcel\Chart[]
     */
    public function getChartCollection()
    {
        return $this->_chartCollection;
    }

    /**
     * Add chart
     *
     * @param PHPExcel\Chart $pChart
     * @param int|null $iChartIndex Index where chart should go (0,1,..., or null for last)
     * @return PHPExcel\Chart
     */
    public function addChart(Chart $pChart = null, $iChartIndex = null)
    {
        $pChart->setWorksheet($this);
        if (is_null($iChartIndex)) {
            $this->_chartCollection[] = $pChart;
        } else {
            // Insert the chart at the requested index
            array_splice($this->_chartCollection, $iChartIndex, 0, array($pChart));
        }

        return $pChart;
    }

    /**
     * Return the count of charts on this worksheet
     *
     * @return int        The number of charts
     */
    public function getChartCount()
    {
        return count($this->_chartCollection);
    }

    /**
     * Get a chart by its index position
     *
     * @param string $index Chart index position
     * @return false|PHPExcel\Chart
     * @throws PHPExcel\Exception
     */
    public function getChartByIndex($index = null)
    {
        $chartCount = count($this->_chartCollection);
        if ($chartCount == 0) {
            return false;
        }
        if (is_null($index)) {
            $index = --$chartCount;
        }
        if (!isset($this->_chartCollection[$index])) {
            return false;
        }

        return $this->_chartCollection[$index];
    }

    /**
     * Return an array of the names of charts on this worksheet
     *
     * @return string[] The names of charts
     * @throws PHPExcel\Exception
     */
    public function getChartNames()
    {
        $chartNames = array();
        foreach($this->_chartCollection as $chart) {
            $chartNames[] = $chart->getName();
        }
        return $chartNames;
    }

    /**
     * Get a chart by name
     *
     * @param string $chartName Chart name
     * @return false|PHPExcel\Chart
     * @throws PHPExcel\Exception
     */
    public function getChartByName($chartName = '')
    {
        $chartCount = count($this->_chartCollection);
        if ($chartCount == 0) {
            return false;
        }
        foreach($this->_chartCollection as $index => $chart) {
            if ($chart->getName() == $chartName) {
                return $this->_chartCollection[$index];
            }
        }
        return false;
    }

    /**
     * Refresh column dimensions
     *
     * @return PHPExcel\Worksheet
     */
    public function refreshColumnDimensions()
    {
        $currentColumnDimensions = $this->getColumnDimensions();
        $newColumnDimensions = array();

        foreach ($currentColumnDimensions as $objColumnDimension) {
            $newColumnDimensions[$objColumnDimension->getColumnIndex()] = $objColumnDimension;
        }

        $this->_columnDimensions = $newColumnDimensions;

        return $this;
    }

    /**
     * Refresh row dimensions
     *
     * @return PHPExcel\Worksheet
     */
    public function refreshRowDimensions()
    {
        $currentRowDimensions = $this->getRowDimensions();
        $newRowDimensions = array();

        foreach ($currentRowDimensions as $objRowDimension) {
            $newRowDimensions[$objRowDimension->getRowIndex()] = $objRowDimension;
        }

        $this->_rowDimensions = $newRowDimensions;

        return $this;
    }

    /**
     * Calculate worksheet dimension
     *
     * @return string  String containing the dimension of this worksheet
     */
    public function calculateWorksheetDimension()
    {
        // Return
        return 'A1' . ':' .  $this->getHighestColumn() . $this->getHighestRow();
    }

    /**
     * Calculate worksheet data dimension
     *
     * @return string  String containing the dimension of this worksheet that actually contain data
     */
    public function calculateWorksheetDataDimension()
    {
        // Return
        return 'A1' . ':' .  $this->getHighestDataColumn() . $this->getHighestDataRow();
    }

    /**
     * Calculate widths for auto-size columns
     *
     * @param  boolean  $calculateMergeCells  Calculate merge cell width
     * @return PHPExcel\Worksheet;
     */
    public function calculateColumnWidths($calculateMergeCells = false)
    {
        // initialize $autoSizes array
        $autoSizes = array();
        foreach ($this->getColumnDimensions() as $colDimension) {
            if ($colDimension->getAutoSize()) {
                $autoSizes[$colDimension->getColumnIndex()] = -1;
            }
        }

        // There is only something to do if there are some auto-size columns
        if (!empty($autoSizes)) {

            // build list of cells references that participate in a merge
            $isMergeCell = array();
            foreach ($this->getMergeCells() as $cells) {
                foreach (Cell::extractAllCellReferencesInRange($cells) as $cellReference) {
                    $isMergeCell[$cellReference] = true;
                }
            }

            // loop through all cells in the worksheet
            foreach ($this->getCellCollection(false) as $cellID) {
                $cell = $this->getCell($cellID);
                if (isset($autoSizes[$this->_cellCollection->getCurrentColumn()])) {
                    // Determine width if cell does not participate in a merge
                    if (!isset($isMergeCell[$this->_cellCollection->getCurrentAddress()])) {
                        // Calculated value
                        // To formatted string
                        $cellValue = Style_NumberFormat::toFormattedString(
                            $cell->getCalculatedValue(),
                            $this->getParent()->getCellXfByIndex($cell->getXfIndex())->getNumberFormat()->getFormatCode()
                        );

                        $autoSizes[$this->_cellCollection->getCurrentColumn()] = max(
                            (float) $autoSizes[$this->_cellCollection->getCurrentColumn()],
                            (float) Shared_Font::calculateColumnWidth(
                                $this->getParent()->getCellXfByIndex($cell->getXfIndex())->getFont(),
                                $cellValue,
                                $this->getParent()->getCellXfByIndex($cell->getXfIndex())->getAlignment()->getTextRotation(),
                                $this->getDefaultStyle()->getFont()
                            )
                        );
                    }
                }
            }

            // adjust column widths
            foreach ($autoSizes as $columnIndex => $width) {
                if ($width == -1) $width = $this->getDefaultColumnDimension()->getWidth();
                $this->getColumnDimension($columnIndex)->setWidth($width);
            }
        }

        return $this;
    }

    /**
     * Get parent
     *
     * @return PHPExcel\Workbook
     */
    public function getParent() {
        return $this->_parent;
    }

    /**
     * Re-bind parent
     *
     * @param PHPExcel\Workbook $parent
     * @return PHPExcel\Worksheet
     */
    public function rebindParent(Workbook $parent) {
        $namedRanges = $this->_parent->getNamedRanges();
        foreach ($namedRanges as $namedRange) {
            $parent->addNamedRange($namedRange);
        }

        $this->_parent->removeSheetByIndex(
            $this->_parent->getIndex($this)
        );
        $this->_parent = $parent;

        return $this;
    }

    /**
     * Get title
     *
     * @return string
     */
    public function getTitle()
    {
        return $this->_title;
    }

    /**
     * Set title
     *
     * @param string $pValue String containing the dimension of this worksheet
     * @param string $updateFormulaCellReferences boolean Flag indicating whether cell references in formulae should
     *            be updated to reflect the new sheet name.
     *          This should be left as the default true, unless you are
     *          certain that no formula cells on any worksheet contain
     *          references to this worksheet
     * @return PHPExcel\Worksheet
     */
    public function setTitle($pValue = 'Worksheet', $updateFormulaCellReferences = true)
    {
        // Is this a 'rename' or not?
        if ($this->getTitle() == $pValue) {
            return $this;
        }

        // Syntax check
        self::_checkSheetTitle($pValue);

        // Old title
        $oldTitle = $this->getTitle();

        if ($this->_parent) {
            // Is there already such sheet name?
            if ($this->_parent->sheetNameExists($pValue)) {
                // Use name, but append with lowest possible integer

                if (Shared_String::CountCharacters($pValue) > 29) {
                    $pValue = Shared_String::Substring($pValue,0,29);
                }
                $i = 1;
                while ($this->_parent->sheetNameExists($pValue . ' ' . $i)) {
                    ++$i;
                    if ($i == 10) {
                        if (Shared_String::CountCharacters($pValue) > 28) {
                            $pValue = Shared_String::Substring($pValue,0,28);
                        }
                    } elseif ($i == 100) {
                        if (Shared_String::CountCharacters($pValue) > 27) {
                            $pValue = Shared_String::Substring($pValue,0,27);
                        }
                    }
                }

                $altTitle = $pValue . ' ' . $i;
                return $this->setTitle($altTitle,$updateFormulaCellReferences);
            }
        }

        // Set title
        $this->_title = $pValue;
        $this->_dirty = true;

        if ($this->_parent) {
            // New title
            $newTitle = $this->getTitle();
            Calculation::getInstance($this->_parent)
                ->renameCalculationCacheForWorksheet($oldTitle, $newTitle);
            if ($updateFormulaCellReferences)
                ReferenceHelper::getInstance()->updateNamedFormulas($this->_parent, $oldTitle, $newTitle);
        }

        return $this;
    }

    /**
     * Get sheet state
     *
     * @return string Sheet state (visible, hidden, veryHidden)
     */
    public function getSheetState() {
        return $this->_sheetState;
    }

    /**
     * Set sheet state
     *
     * @param string $value Sheet state (visible, hidden, veryHidden)
     * @return PHPExcel\Worksheet
     */
    public function setSheetState($value = Worksheet::SHEETSTATE_VISIBLE) {
        $this->_sheetState = $value;
        return $this;
    }

    /**
     * Get page setup
     *
     * @return PHPExcel\Worksheet_PageSetup
     */
    public function getPageSetup()
    {
        return $this->_pageSetup;
    }

    /**
     * Set page setup
     *
     * @param PHPExcel\Worksheet_PageSetup    $pValue
     * @return PHPExcel\Worksheet
     */
    public function setPageSetup(Worksheet_PageSetup $pValue)
    {
        $this->_pageSetup = $pValue;
        return $this;
    }

    /**
     * Get page margins
     *
     * @return PHPExcel\Worksheet_PageMargins
     */
    public function getPageMargins()
    {
        return $this->_pageMargins;
    }

    /**
     * Set page margins
     *
     * @param PHPExcel\Worksheet_PageMargins    $pValue
     * @return PHPExcel\Worksheet
     */
    public function setPageMargins(Worksheet_PageMargins $pValue)
    {
        $this->_pageMargins = $pValue;
        return $this;
    }

    /**
     * Get page header/footer
     *
     * @return PHPExcel\Worksheet_HeaderFooter
     */
    public function getHeaderFooter()
    {
        return $this->_headerFooter;
    }

    /**
     * Set page header/footer
     *
     * @param PHPExcel\Worksheet_HeaderFooter    $pValue
     * @return PHPExcel\Worksheet
     */
    public function setHeaderFooter(Worksheet_HeaderFooter $pValue)
    {
        $this->_headerFooter = $pValue;
        return $this;
    }

    /**
     * Get sheet view
     *
     * @return PHPExcel\Worksheet_SheetView
     */
    public function getSheetView()
    {
        return $this->_sheetView;
    }

    /**
     * Set sheet view
     *
     * @param PHPExcel\Worksheet_SheetView    $pValue
     * @return PHPExcel\Worksheet
     */
    public function setSheetView(Worksheet_SheetView $pValue)
    {
        $this->_sheetView = $pValue;
        return $this;
    }

    /**
     * Get Protection
     *
     * @return PHPExcel\Worksheet_Protection
     */
    public function getProtection()
    {
        return $this->_protection;
    }

    /**
     * Set Protection
     *
     * @param PHPExcel\Worksheet_Protection    $pValue
     * @return PHPExcel\Worksheet
     */
    public function setProtection(Worksheet_Protection $pValue)
    {
        $this->_protection = $pValue;
        $this->_dirty = true;

        return $this;
    }

    /**
     * Get highest worksheet column
     *
     * @return string Highest column name
     */
    public function getHighestColumn()
    {
        return $this->_cachedHighestColumn;
    }

    /**
     * Get highest worksheet column that contains data
     *
     * @return string Highest column name that contains data
     */
    public function getHighestDataColumn()
    {
        return $this->_cellCollection->getHighestColumn();
    }

    /**
     * Get highest worksheet row
     *
     * @return int Highest row number
     */
    public function getHighestRow()
    {
        return $this->_cachedHighestRow;
    }

    /**
     * Get highest worksheet row that contains data
     *
     * @return string Highest row number that contains data
     */
    public function getHighestDataRow()
    {
        return $this->_cellCollection->getHighestRow();
    }

    /**
     * Get highest worksheet column and highest row that have cell records
     *
     * @return array Highest column name and highest row number
     */
    public function getHighestRowAndColumn()
    {
        return $this->_cellCollection->getHighestRowAndColumn();
    }

    /**
     * Set a cell value
     *
     * @param string $pCoordinate Coordinate of the cell
     * @param mixed $pValue Value of the cell
     * @param bool $returnCell   Return the worksheet (false, default) or the cell (true)
     * @return PHPExcel\Worksheet|PHPExcel\Cell    Depending on the last parameter being specified
     */
    public function setCellValue($pCoordinate = 'A1', $pValue = null, $returnCell = false)
    {
        $cell = $this->getCell($pCoordinate)->setValue($pValue);
        return ($returnCell) ? $cell : $this;
    }

    /**
     * Set a cell value by using numeric cell coordinates
     *
     * @param string  $pColumn Numeric column coordinate of the cell (A = 0)
     * @param string  $pRow Numeric row coordinate of the cell
     * @param mixed   $pValue Value of the cell
     * @param bool    $returnCell Return the worksheet (false, default) or the cell (true)
     * @return PHPExcel\Worksheet|PHPExcel\Cell    Depending on the last parameter being specified
     */
    public function setCellValueByColumnAndRow($pColumn = 0, $pRow = 1, $pValue = null, $returnCell = false)
    {
        $cell = $this->getCellByColumnAndRow($pColumn, $pRow)->setValue($pValue);
        return ($returnCell) ? $cell : $this;
    }

    /**
     * Set a cell value
     *
     * @param string $pCoordinate Coordinate of the cell
     * @param mixed  $pValue Value of the cell
     * @param string $pDataType Explicit data type
     * @param bool $returnCell Return the worksheet (false, default) or the cell (true)
     * @return PHPExcel\Worksheet|PHPExcel\Cell    Depending on the last parameter being specified
     */
    public function setCellValueExplicit($pCoordinate = 'A1', $pValue = null, $pDataType = Cell_DataType::TYPE_STRING, $returnCell = false)
    {
        // Set value
        $cell = $this->getCell($pCoordinate)->setValueExplicit($pValue, $pDataType);
        return ($returnCell) ? $cell : $this;
    }

    /**
     * Set a cell value by using numeric cell coordinates
     *
     * @param string $pColumn Numeric column coordinate of the cell
     * @param string $pRow Numeric row coordinate of the cell
     * @param mixed $pValue Value of the cell
     * @param string $pDataType Explicit data type
     * @param bool $returnCell Return the worksheet (false, default) or the cell (true)
     * @return PHPExcel\Worksheet|PHPExcel\Cell    Depending on the last parameter being specified
     */
    public function setCellValueExplicitByColumnAndRow($pColumn = 0, $pRow = 1, $pValue = null, $pDataType = Cell_DataType::TYPE_STRING, $returnCell = false)
    {
        $cell = $this->getCellByColumnAndRow($pColumn, $pRow)->setValueExplicit($pValue, $pDataType);
        return ($returnCell) ? $cell : $this;
    }

    /**
     * Get cell at a specific coordinate
     *
     * @param string $pCoordinate    Coordinate of the cell
     * @throws PHPExcel\Exception
     * @return PHPExcel\Cell Cell that was found
     */
    public function getCell($pCoordinate = 'A1')
    {
        // Check cell collection
        if ($this->_cellCollection->isDataSet($pCoordinate)) {
            return $this->_cellCollection->getCacheData($pCoordinate);
        }

        // Worksheet reference?
        if (strpos($pCoordinate, '!') !== false) {
            $worksheetReference = Worksheet::extractSheetTitle($pCoordinate, true);
            return $this->_parent->getSheetByName($worksheetReference[0])->getCell($worksheetReference[1]);
        }

        // Named range?
        if ((!preg_match('/^'.Calculation::CALCULATION_REGEXP_CELLREF.'$/i', $pCoordinate, $matches)) &&
            (preg_match('/^'.Calculation::CALCULATION_REGEXP_NAMEDRANGE.'$/i', $pCoordinate, $matches))) {
            $namedRange = NamedRange::resolveRange($pCoordinate, $this);
            if ($namedRange !== null) {
                $pCoordinate = $namedRange->getRange();
                return $namedRange->getWorksheet()->getCell($pCoordinate);
            }
        }

        // Uppercase coordinate
        $pCoordinate = strtoupper($pCoordinate);

        if (strpos($pCoordinate, ':') !== false || strpos($pCoordinate, ',') !== false) {
            throw new Exception('Cell coordinate can not be a range of cells.');
        } elseif (strpos($pCoordinate, '$') !== false) {
            throw new Exception('Cell coordinate must not be absolute.');
        }

        // Create new cell object
        return $this->_createNewCell($pCoordinate);
    }

    /**
     * Get cell at a specific coordinate by using numeric cell coordinates
     *
     * @param  string $pColumn Numeric column coordinate of the cell
     * @param string $pRow Numeric row coordinate of the cell
     * @return PHPExcel\Cell Cell that was found
     */
    public function getCellByColumnAndRow($pColumn = 0, $pRow = 1)
    {
        $columnLetter = Cell::stringFromColumnIndex($pColumn);
        $coordinate = $columnLetter . $pRow;

        if ($this->_cellCollection->isDataSet($coordinate)) {
            return $this->_cellCollection->getCacheData($coordinate);
        }

        return $this->_createNewCell($coordinate);
    }

    /**
     * Create a new cell at the specified coordinate
     *
     * @param string $pCoordinate    Coordinate of the cell
     * @return PHPExcel\Cell Cell that was created
     */
    private function _createNewCell($pCoordinate)
    {
        $cell = $this->_cellCollection->addCacheData(
            $pCoordinate,
            new Cell(
                null,
                Cell_DataType::TYPE_NULL,
                $this
            )
        );
        $this->_cellCollectionIsSorted = false;

        // Coordinates
        $aCoordinates = Cell::coordinateFromString($pCoordinate);
        if (Cell::columnIndexFromString($this->_cachedHighestColumn) < Cell::columnIndexFromString($aCoordinates[0]))
            $this->_cachedHighestColumn = $aCoordinates[0];
        $this->_cachedHighestRow = max($this->_cachedHighestRow, $aCoordinates[1]);

        // Cell needs appropriate xfIndex from dimensions records
        //    but don't create dimension records if they don't already exist
        $rowDimension    = $this->getRowDimension($aCoordinates[1], false);
        $columnDimension = $this->getColumnDimension($aCoordinates[0], false);

        if ($rowDimension !== null && $rowDimension->getXfIndex() > 0) {
            // then there is a row dimension with explicit style, assign it to the cell
            $cell->setXfIndex($rowDimension->getXfIndex());
        } elseif ($columnDimension !== null && $columnDimension->getXfIndex() > 0) {
            // then there is a column dimension, assign it to the cell
            $cell->setXfIndex($columnDimension->getXfIndex());
        }

        return $cell;
    }

    /**
     * Does the cell at a specific coordinate exist?
     *
     * @param string $pCoordinate  Coordinate of the cell
     * @throws PHPExcel\Exception
     * @return boolean
     */
    public function cellExists($pCoordinate = 'A1')
    {
        // Worksheet reference?
        if (strpos($pCoordinate, '!') !== false) {
            $worksheetReference = self::extractSheetTitle($pCoordinate, true);
            return $this->_parent->getSheetByName($worksheetReference[0])->cellExists($worksheetReference[1]);
        }

        // Named range?
        if ((!preg_match('/^'.Calculation::CALCULATION_REGEXP_CELLREF.'$/i', $pCoordinate, $matches)) &&
            (preg_match('/^'.Calculation::CALCULATION_REGEXP_NAMEDRANGE.'$/i', $pCoordinate, $matches))) {
            $namedRange = NamedRange::resolveRange($pCoordinate, $this);
            if ($namedRange !== null) {
                $pCoordinate = $namedRange->getRange();
                if ($this->getHashCode() != $namedRange->getWorksheet()->getHashCode()) {
                    if (!$namedRange->getLocalOnly()) {
                        return $namedRange->getWorksheet()->cellExists($pCoordinate);
                    } else {
                        throw new Exception('Named range ' . $namedRange->getName() . ' is not accessible from within sheet ' . $this->getTitle());
                    }
                }
            }
            else { return false; }
        }

        // Uppercase coordinate
        $pCoordinate = strtoupper($pCoordinate);

        if (strpos($pCoordinate,':') !== false || strpos($pCoordinate,',') !== false) {
            throw new Exception('Cell coordinate can not be a range of cells.');
        } elseif (strpos($pCoordinate,'$') !== false) {
            throw new Exception('Cell coordinate must not be absolute.');
        } else {
            // Coordinates
            $aCoordinates = Cell::coordinateFromString($pCoordinate);

            // Cell exists?
            return $this->_cellCollection->isDataSet($pCoordinate);
        }
    }

    /**
     * Cell at a specific coordinate by using numeric cell coordinates exists?
     *
     * @param string $pColumn Numeric column coordinate of the cell
     * @param string $pRow Numeric row coordinate of the cell
     * @return boolean
     */
    public function cellExistsByColumnAndRow($pColumn = 0, $pRow = 1)
    {
        return $this->cellExists(Cell::stringFromColumnIndex($pColumn) . $pRow);
    }

    /**
     * Get row dimension at a specific row
     *
     * @param int $pRow Numeric index of the row
     * @return PHPExcel\Worksheet_RowDimension
     */
    public function getRowDimension($pRow = 1, $create = TRUE)
    {
        // Found
        $found = null;

        // Get row dimension
        if (!isset($this->_rowDimensions[$pRow])) {
            if (!$create)
                return null;
            $this->_rowDimensions[$pRow] = new Worksheet_RowDimension($pRow);

            $this->_cachedHighestRow = max($this->_cachedHighestRow,$pRow);
        }
        return $this->_rowDimensions[$pRow];
    }

    /**
     * Get column dimension at a specific column
     *
     * @param string $pColumn String index of the column
     * @return PHPExcel\Worksheet_ColumnDimension
     */
    public function getColumnDimension($pColumn = 'A', $create = TRUE)
    {
        // Uppercase coordinate
        $pColumn = strtoupper($pColumn);

        // Fetch dimensions
        if (!isset($this->_columnDimensions[$pColumn])) {
            if (!$create)
                return null;
            $this->_columnDimensions[$pColumn] = new Worksheet_ColumnDimension($pColumn);

            if (Cell::columnIndexFromString($this->_cachedHighestColumn) < Cell::columnIndexFromString($pColumn))
                $this->_cachedHighestColumn = $pColumn;
        }
        return $this->_columnDimensions[$pColumn];
    }

    /**
     * Get column dimension at a specific column by using numeric cell coordinates
     *
     * @param string $pColumn Numeric column coordinate of the cell
     * @return PHPExcel\Worksheet_ColumnDimension
     */
    public function getColumnDimensionByColumn($pColumn = 0)
    {
        return $this->getColumnDimension(Cell::stringFromColumnIndex($pColumn));
    }

    /**
     * Get styles
     *
     * @return PHPExcel\Style[]
     */
    public function getStyles()
    {
        return $this->_styles;
    }

    /**
     * Get default style of workbook.
     *
     * @deprecated
     * @return PHPExcel\Style
     * @throws PHPExcel\Exception
     */
    public function getDefaultStyle()
    {
        return $this->_parent->getDefaultStyle();
    }

    /**
     * Set default style - should only be used by PHPExcel\IReader implementations!
     *
     * @deprecated
     * @param PHPExcel\Style $pValue
     * @throws PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function setDefaultStyle(Style $pValue)
    {
        $this->_parent->getDefaultStyle()->applyFromArray(array(
            'font' => array(
                'name' => $pValue->getFont()->getName(),
                'size' => $pValue->getFont()->getSize(),
            ),
        ));
        return $this;
    }

    /**
     * Get style for cell
     *
     * @param string $pCellCoordinate Cell coordinate to get style for
     * @return PHPExcel\Style
     * @throws PHPExcel\Exception
     */
    public function getStyle($pCellCoordinate = 'A1')
    {
        // set this sheet as active
        #$this->_parent->setActiveSheetIndex($this->_parent->getIndex($this));
        #will be activated every time, once is enough! 

        // set cell coordinate as active
        $this->setSelectedCells($pCellCoordinate);

        return $this->_parent->getCellXfSupervisor();
    }

    /**
     * Get conditional styles for a cell
     *
     * @param string $pCoordinate
     * @return PHPExcel\Style_Conditional[]
     */
    public function getConditionalStyles($pCoordinate = 'A1')
    {
        if (!isset($this->_conditionalStylesCollection[$pCoordinate])) {
            $this->_conditionalStylesCollection[$pCoordinate] = array();
        }
        return $this->_conditionalStylesCollection[$pCoordinate];
    }

    /**
     * Do conditional styles exist for this cell?
     *
     * @param string $pCoordinate
     * @return boolean
     */
    public function conditionalStylesExists($pCoordinate = 'A1')
    {
        if (isset($this->_conditionalStylesCollection[$pCoordinate])) {
            return true;
        }
        return false;
    }

    /**
     * Removes conditional styles for a cell
     *
     * @param string $pCoordinate
     * @return PHPExcel\Worksheet
     */
    public function removeConditionalStyles($pCoordinate = 'A1')
    {
        unset($this->_conditionalStylesCollection[$pCoordinate]);
        return $this;
    }

    /**
     * Get collection of conditional styles
     *
     * @return array
     */
    public function getConditionalStylesCollection()
    {
        return $this->_conditionalStylesCollection;
    }

    /**
     * Set conditional styles
     *
     * @param $pCoordinate string E.g. 'A1'
     * @param $pValue PHPExcel\Style_Conditional[]
     * @return PHPExcel\Worksheet
     */
    public function setConditionalStyles($pCoordinate = 'A1', $pValue)
    {
        $this->_conditionalStylesCollection[$pCoordinate] = $pValue;
        return $this;
    }

    /**
     * Get style for cell by using numeric cell coordinates
     *
     * @param int $pColumn  Numeric column coordinate of the cell
     * @param int $pRow Numeric row coordinate of the cell
     * @return PHPExcel\Style
     */
    public function getStyleByColumnAndRow($pColumn = 0, $pRow = 1)
    {
        return $this->getStyle(Cell::stringFromColumnIndex($pColumn) . $pRow);
    }

    /**
     * Set shared cell style to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range!
     *
     * @deprecated
     * @param PHPExcel\Style $pSharedCellStyle Cell style to share
     * @param string $pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @throws PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function setSharedStyle(Style $pSharedCellStyle = null, $pRange = '')
    {
        $this->duplicateStyle($pSharedCellStyle, $pRange);
        return $this;
    }

    /**
     * Duplicate cell style to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range!
     *
     * @param PHPExcel\Style $pCellStyle Cell style to duplicate
     * @param string $pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @throws PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function duplicateStyle(Style $pCellStyle = null, $pRange = '')
    {
        // make sure we have a real style and not supervisor
        $style = $pCellStyle->getIsSupervisor() ? $pCellStyle->getSharedComponent() : $pCellStyle;

        // Add the style to the workbook if necessary
        $workbook = $this->_parent;
        if ($existingStyle = $this->_parent->getCellXfByHashCode($pCellStyle->getHashCode())) {
            // there is already such cell Xf in our collection
            $xfIndex = $existingStyle->getIndex();
        } else {
            // we don't have such a cell Xf, need to add
            $workbook->addCellXf($pCellStyle);
            $xfIndex = $pCellStyle->getIndex();
        }

        // Uppercase coordinate
        $pRange = strtoupper($pRange);

        // Is it a cell range or a single cell?
        $rangeA    = '';
        $rangeB    = '';
        if (strpos($pRange, ':') === false) {
            $rangeA = $pRange;
            $rangeB = $pRange;
        } else {
            list($rangeA, $rangeB) = explode(':', $pRange);
        }

        // Calculate range outer borders
        $rangeStart = Cell::coordinateFromString($rangeA);
        $rangeEnd   = Cell::coordinateFromString($rangeB);

        // Translate column into index
        $rangeStart[0]  = Cell::columnIndexFromString($rangeStart[0]) - 1;
        $rangeEnd[0]    = Cell::columnIndexFromString($rangeEnd[0]) - 1;

        // Make sure we can loop upwards on rows and columns
        if ($rangeStart[0] > $rangeEnd[0] && $rangeStart[1] > $rangeEnd[1]) {
            $tmp = $rangeStart;
            $rangeStart = $rangeEnd;
            $rangeEnd = $tmp;
        }

        // Loop through cells and apply styles
        for ($col = $rangeStart[0]; $col <= $rangeEnd[0]; ++$col) {
            for ($row = $rangeStart[1]; $row <= $rangeEnd[1]; ++$row) {
                $this->getCell(Cell::stringFromColumnIndex($col) . $row)->setXfIndex($xfIndex);
            }
        }

        return $this;
    }

    /**
     * Duplicate conditional style to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range!
     *
     * @param    array of PHPExcel\Style_Conditional    $pCellStyle    Cell style to duplicate
     * @param string $pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @throws PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function duplicateConditionalStyle(array $pCellStyle = null, $pRange = '')
    {
        foreach($pCellStyle as $cellStyle) {
            if (!($cellStyle instanceof Style_Conditional)) {
                throw new Exception('Style is not a conditional style');
            }
        }

        // Uppercase coordinate
        $pRange = strtoupper($pRange);

        // Is it a cell range or a single cell?
        $rangeA    = '';
        $rangeB    = '';
        if (strpos($pRange, ':') === false) {
            $rangeA = $pRange;
            $rangeB = $pRange;
        } else {
            list($rangeA, $rangeB) = explode(':', $pRange);
        }

        // Calculate range outer borders
        $rangeStart = Cell::coordinateFromString($rangeA);
        $rangeEnd   = Cell::coordinateFromString($rangeB);

        // Translate column into index
        $rangeStart[0] = Cell::columnIndexFromString($rangeStart[0]) - 1;
        $rangeEnd[0]   = Cell::columnIndexFromString($rangeEnd[0]) - 1;

        // Make sure we can loop upwards on rows and columns
        if ($rangeStart[0] > $rangeEnd[0] && $rangeStart[1] > $rangeEnd[1]) {
            $tmp = $rangeStart;
            $rangeStart = $rangeEnd;
            $rangeEnd = $tmp;
        }

        // Loop through cells and apply styles
        for ($col = $rangeStart[0]; $col <= $rangeEnd[0]; ++$col) {
            for ($row = $rangeStart[1]; $row <= $rangeEnd[1]; ++$row) {
                $this->setConditionalStyles(Cell::stringFromColumnIndex($col) . $row, $pCellStyle);
            }
        }

        return $this;
    }

    /**
     * Duplicate cell style array to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range,
     * if they are in the styles array. For example, if you decide to set a range of
     * cells to font bold, only include font bold in the styles array.
     *
     * @deprecated
     * @param array $pStyles Array containing style information
     * @param string $pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @param boolean $pAdvanced Advanced mode for setting borders.
     * @throws PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function duplicateStyleArray($pStyles = null, $pRange = '', $pAdvanced = true)
    {
        $this->getStyle($pRange)->applyFromArray($pStyles, $pAdvanced);
        return $this;
    }

    /**
     * Set break on a cell
     *
     * @param string $pCell Cell coordinate (e.g. A1)
     * @param int $pBreak Break type (type of PHPExcel\Worksheet::BREAK_*)
     * @throws PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function setBreak($pCell = 'A1', $pBreak = self::BREAK_NONE)
    {
        // Uppercase coordinate
        $pCell = strtoupper($pCell);

        if ($pCell != '') {
            if ($pBreak == self::BREAK_NONE) {
                if (isset($this->_breaks[$pCell])) {
                    unset($this->_breaks[$pCell]);
                }
            } else {
                $this->_breaks[$pCell] = $pBreak;
            }
        } else {
            throw new Exception('No cell coordinate specified.');
        }

        return $this;
    }

    /**
     * Set break on a cell by using numeric cell coordinates
     *
     * @param integer $pColumn Numeric column coordinate of the cell
     * @param integer $pRow Numeric row coordinate of the cell
     * @param  integer $pBreak Break type (type of PHPExcel\Worksheet::BREAK_*)
     * @return PHPExcel\Worksheet
     */
    public function setBreakByColumnAndRow($pColumn = 0, $pRow = 1, $pBreak = Worksheet::BREAK_NONE)
    {
        return $this->setBreak(Cell::stringFromColumnIndex($pColumn) . $pRow, $pBreak);
    }

    /**
     * Get breaks
     *
     * @return array[]
     */
    public function getBreaks()
    {
        return $this->_breaks;
    }

    /**
     * Set merge on a cell range
     *
     * @param string $pRange  Cell range (e.g. A1:E1)
     * @throws PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function mergeCells($pRange = 'A1:A1')
    {
        // Uppercase coordinate
        $pRange = strtoupper($pRange);

        if (strpos($pRange,':') !== false) {
            $this->_mergeCells[$pRange] = $pRange;

            // make sure cells are created

            // get the cells in the range
            $aReferences = Cell::extractAllCellReferencesInRange($pRange);

            // create upper left cell if it does not already exist
            $upperLeft = $aReferences[0];
            if (!$this->cellExists($upperLeft)) {
                $this->getCell($upperLeft)->setValueExplicit(null, Cell_DataType::TYPE_NULL);
            }

            // create or blank out the rest of the cells in the range
            $count = count($aReferences);
            for ($i = 1; $i < $count; $i++) {
                $this->getCell($aReferences[$i])->setValueExplicit(null, Cell_DataType::TYPE_NULL);
            }

        } else {
            throw new Exception('Merge must be set on a range of cells.');
        }

        return $this;
    }

    /**
     * Set merge on a cell range by using numeric cell coordinates
     *
     * @param int $pColumn1    Numeric column coordinate of the first cell
     * @param int $pRow1        Numeric row coordinate of the first cell
     * @param int $pColumn2    Numeric column coordinate of the last cell
     * @param int $pRow2        Numeric row coordinate of the last cell
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function mergeCellsByColumnAndRow($pColumn1 = 0, $pRow1 = 1, $pColumn2 = 0, $pRow2 = 1)
    {
        $cellRange = Cell::stringFromColumnIndex($pColumn1) . $pRow1 . ':' . Cell::stringFromColumnIndex($pColumn2) . $pRow2;
        return $this->mergeCells($cellRange);
    }

    /**
     * Remove merge on a cell range
     *
     * @param    string            $pRange        Cell range (e.g. A1:E1)
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function unmergeCells($pRange = 'A1:A1')
    {
        // Uppercase coordinate
        $pRange = strtoupper($pRange);

        if (strpos($pRange,':') !== false) {
            if (isset($this->_mergeCells[$pRange])) {
                unset($this->_mergeCells[$pRange]);
            } else {
                throw new Exception('Cell range ' . $pRange . ' not known as merged.');
            }
        } else {
            throw new Exception('Merge can only be removed from a range of cells.');
        }

        return $this;
    }

    /**
     * Remove merge on a cell range by using numeric cell coordinates
     *
     * @param int $pColumn1    Numeric column coordinate of the first cell
     * @param int $pRow1        Numeric row coordinate of the first cell
     * @param int $pColumn2    Numeric column coordinate of the last cell
     * @param int $pRow2        Numeric row coordinate of the last cell
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function unmergeCellsByColumnAndRow($pColumn1 = 0, $pRow1 = 1, $pColumn2 = 0, $pRow2 = 1)
    {
        $cellRange = Cell::stringFromColumnIndex($pColumn1) . $pRow1 . ':' . Cell::stringFromColumnIndex($pColumn2) . $pRow2;
        return $this->unmergeCells($cellRange);
    }

    /**
     * Get merge cells array.
     *
     * @return array[]
     */
    public function getMergeCells()
    {
        return $this->_mergeCells;
    }

    /**
     * Set merge cells array for the entire sheet. Use instead mergeCells() to merge
     * a single cell range.
     *
     * @param array
     */
    public function setMergeCells($pValue = array())
    {
        $this->_mergeCells = $pValue;

        return $this;
    }

    /**
     * Set protection on a cell range
     *
     * @param    string            $pRange                Cell (e.g. A1) or cell range (e.g. A1:E1)
     * @param    string            $pPassword            Password to unlock the protection
     * @param    boolean        $pAlreadyHashed    If the password has already been hashed, set this to true
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function protectCells($pRange = 'A1', $pPassword = '', $pAlreadyHashed = false)
    {
        // Uppercase coordinate
        $pRange = strtoupper($pRange);

        if (!$pAlreadyHashed) {
            $pPassword = Shared_PasswordHasher::hashPassword($pPassword);
        }
        $this->_protectedCells[$pRange] = $pPassword;

        return $this;
    }

    /**
     * Set protection on a cell range by using numeric cell coordinates
     *
     * @param int  $pColumn1            Numeric column coordinate of the first cell
     * @param int  $pRow1                Numeric row coordinate of the first cell
     * @param int  $pColumn2            Numeric column coordinate of the last cell
     * @param int  $pRow2                Numeric row coordinate of the last cell
     * @param string $pPassword            Password to unlock the protection
     * @param    boolean $pAlreadyHashed    If the password has already been hashed, set this to true
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function protectCellsByColumnAndRow($pColumn1 = 0, $pRow1 = 1, $pColumn2 = 0, $pRow2 = 1, $pPassword = '', $pAlreadyHashed = false)
    {
        $cellRange = Cell::stringFromColumnIndex($pColumn1) . $pRow1 . ':' . Cell::stringFromColumnIndex($pColumn2) . $pRow2;
        return $this->protectCells($cellRange, $pPassword, $pAlreadyHashed);
    }

    /**
     * Remove protection on a cell range
     *
     * @param    string            $pRange        Cell (e.g. A1) or cell range (e.g. A1:E1)
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function unprotectCells($pRange = 'A1')
    {
        // Uppercase coordinate
        $pRange = strtoupper($pRange);

        if (isset($this->_protectedCells[$pRange])) {
            unset($this->_protectedCells[$pRange]);
        } else {
            throw new Exception('Cell range ' . $pRange . ' not known as protected.');
        }
        return $this;
    }

    /**
     * Remove protection on a cell range by using numeric cell coordinates
     *
     * @param int  $pColumn1            Numeric column coordinate of the first cell
     * @param int  $pRow1                Numeric row coordinate of the first cell
     * @param int  $pColumn2            Numeric column coordinate of the last cell
     * @param int $pRow2                Numeric row coordinate of the last cell
     * @param string $pPassword            Password to unlock the protection
     * @param    boolean $pAlreadyHashed    If the password has already been hashed, set this to true
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function unprotectCellsByColumnAndRow($pColumn1 = 0, $pRow1 = 1, $pColumn2 = 0, $pRow2 = 1, $pPassword = '', $pAlreadyHashed = false)
    {
        $cellRange = Cell::stringFromColumnIndex($pColumn1) . $pRow1 . ':' . Cell::stringFromColumnIndex($pColumn2) . $pRow2;
        return $this->unprotectCells($cellRange, $pPassword, $pAlreadyHashed);
    }

    /**
     * Get protected cells
     *
     * @return array[]
     */
    public function getProtectedCells()
    {
        return $this->_protectedCells;
    }

    /**
     *    Get Autofilter
     *
     *    @return PHPExcel\Worksheet_AutoFilter
     */
    public function getAutoFilter()
    {
        return $this->_autoFilter;
    }

    /**
     *    Set AutoFilter
     *
     *    @param    PHPExcel\Worksheet_AutoFilter|string   $pValue
     *            A simple string containing a Cell range like 'A1:E10' is permitted for backward compatibility
     *    @throws    PHPExcel\Exception
     *    @return PHPExcel\Worksheet
     */
    public function setAutoFilter($pValue)
    {
        if (is_string($pValue)) {
            $this->_autoFilter->setRange($pValue);
        } elseif(is_object($pValue) && ($pValue instanceof Worksheet_AutoFilter)) {
            $this->_autoFilter = $pValue;
        }
        return $this;
    }

    /**
     *    Set Autofilter Range by using numeric cell coordinates
     *
     *    @param  integer  $pColumn1    Numeric column coordinate of the first cell
     *    @param  integer  $pRow1       Numeric row coordinate of the first cell
     *    @param  integer  $pColumn2    Numeric column coordinate of the second cell
     *    @param  integer  $pRow2       Numeric row coordinate of the second cell
     *    @throws    PHPExcel\Exception
     *    @return PHPExcel\Worksheet
     */
    public function setAutoFilterByColumnAndRow($pColumn1 = 0, $pRow1 = 1, $pColumn2 = 0, $pRow2 = 1)
    {
        return $this->setAutoFilter(
            Cell::stringFromColumnIndex($pColumn1) . $pRow1
            . ':' .
            Cell::stringFromColumnIndex($pColumn2) . $pRow2
        );
    }

    /**
     * Remove autofilter
     *
     * @return PHPExcel\Worksheet
     */
    public function removeAutoFilter()
    {
        $this->_autoFilter->setRange(null);
        return $this;
    }

    /**
     * Get Freeze Pane
     *
     * @return string
     */
    public function getFreezePane()
    {
        return $this->_freezePane;
    }

    /**
     * Freeze Pane
     *
     * @param    string        $pCell        Cell (i.e. A2)
     *                                    Examples:
     *                                        A2 will freeze the rows above cell A2 (i.e row 1)
     *                                        B1 will freeze the columns to the left of cell B1 (i.e column A)
     *                                        B2 will freeze the rows above and to the left of cell A2
     *                                            (i.e row 1 and column A)
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function freezePane($pCell = '')
    {
        // Uppercase coordinate
        $pCell = strtoupper($pCell);

        if (strpos($pCell,':') === false && strpos($pCell,',') === false) {
            $this->_freezePane = $pCell;
        } else {
            throw new Exception('Freeze pane can not be set on a range of cells.');
        }
        return $this;
    }

    /**
     * Freeze Pane by using numeric cell coordinates
     *
     * @param int $pColumn    Numeric column coordinate of the cell
     * @param int $pRow        Numeric row coordinate of the cell
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function freezePaneByColumnAndRow($pColumn = 0, $pRow = 1)
    {
        return $this->freezePane(Cell::stringFromColumnIndex($pColumn) . $pRow);
    }

    /**
     * Unfreeze Pane
     *
     * @return PHPExcel\Worksheet
     */
    public function unfreezePane()
    {
        return $this->freezePane('');
    }

    /**
     * Insert a new row, updating all possible related data
     *
     * @param int $pBefore    Insert before this one
     * @param int $pNumRows    Number of rows to insert
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function insertNewRowBefore($pBefore = 1, $pNumRows = 1) {
        if ($pBefore >= 1) {
            $objReferenceHelper = ReferenceHelper::getInstance();
            $objReferenceHelper->insertNewBefore('A' . $pBefore, 0, $pNumRows, $this);
        } else {
            throw new Exception("Rows can only be inserted before at least row 1.");
        }
        return $this;
    }

    /**
     * Insert a new column, updating all possible related data
     *
     * @param int $pBefore    Insert before this one
     * @param int $pNumCols    Number of columns to insert
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function insertNewColumnBefore($pBefore = 'A', $pNumCols = 1) {
        if (!is_numeric($pBefore)) {
            $objReferenceHelper = ReferenceHelper::getInstance();
            $objReferenceHelper->insertNewBefore($pBefore . '1', $pNumCols, 0, $this);
        } else {
            throw new Exception("Column references should not be numeric.");
        }
        return $this;
    }

    /**
     * Insert a new column, updating all possible related data
     *
     * @param int $pBefore    Insert before this one (numeric column coordinate of the cell)
     * @param int $pNumCols    Number of columns to insert
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function insertNewColumnBeforeByIndex($pBefore = 0, $pNumCols = 1) {
        if ($pBefore >= 0) {
            return $this->insertNewColumnBefore(Cell::stringFromColumnIndex($pBefore), $pNumCols);
        } else {
            throw new Exception("Columns can only be inserted before at least column A (0).");
        }
    }

    /**
     * Delete a row, updating all possible related data
     *
     * @param int $pRow        Remove starting with this one
     * @param int $pNumRows    Number of rows to remove
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function removeRow($pRow = 1, $pNumRows = 1) {
        if ($pRow >= 1) {
            $objReferenceHelper = ReferenceHelper::getInstance();
            $objReferenceHelper->insertNewBefore('A' . ($pRow + $pNumRows), 0, -$pNumRows, $this);
        } else {
            throw new Exception("Rows to be deleted should at least start from row 1.");
        }
        return $this;
    }

    /**
     * Remove a column, updating all possible related data
     *
     * @param int $pColumn    Remove starting with this one
     * @param int $pNumCols    Number of columns to remove
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function removeColumn($pColumn = 'A', $pNumCols = 1) {
        if (!is_numeric($pColumn)) {
            $pColumn = Cell::stringFromColumnIndex(Cell::columnIndexFromString($pColumn) - 1 + $pNumCols);
            $objReferenceHelper = ReferenceHelper::getInstance();
            $objReferenceHelper->insertNewBefore($pColumn . '1', -$pNumCols, 0, $this);
        } else {
            throw new Exception("Column references should not be numeric.");
        }
        return $this;
    }

    /**
     * Remove a column, updating all possible related data
     *
     * @param int $pColumn    Remove starting with this one (numeric column coordinate of the cell)
     * @param int $pNumCols    Number of columns to remove
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function removeColumnByIndex($pColumn = 0, $pNumCols = 1) {
        if ($pColumn >= 0) {
            return $this->removeColumn(Cell::stringFromColumnIndex($pColumn), $pNumCols);
        } else {
            throw new Exception("Columns to be deleted should at least start from column 0");
        }
    }

    /**
     * Show gridlines?
     *
     * @return boolean
     */
    public function getShowGridlines() {
        return $this->_showGridlines;
    }

    /**
     * Set show gridlines
     *
     * @param boolean $pValue    Show gridlines (true/false)
     * @return PHPExcel\Worksheet
     */
    public function setShowGridlines($pValue = false) {
        $this->_showGridlines = $pValue;
        return $this;
    }

    /**
    * Print gridlines?
    *
    * @return boolean
    */
    public function getPrintGridlines() {
        return $this->_printGridlines;
    }

    /**
    * Set print gridlines
    *
    * @param boolean $pValue Print gridlines (true/false)
    * @return PHPExcel\Worksheet
    */
    public function setPrintGridlines($pValue = false) {
        $this->_printGridlines = $pValue;
        return $this;
    }

    /**
    * Show row and column headers?
    *
    * @return boolean
    */
    public function getShowRowColHeaders() {
        return $this->_showRowColHeaders;
    }

    /**
    * Set show row and column headers
    *
    * @param boolean $pValue Show row and column headers (true/false)
    * @return PHPExcel\Worksheet
    */
    public function setShowRowColHeaders($pValue = false) {
        $this->_showRowColHeaders = $pValue;
        return $this;
    }

    /**
     * Show summary below? (Row/Column outlining)
     *
     * @return boolean
     */
    public function getShowSummaryBelow() {
        return $this->_showSummaryBelow;
    }

    /**
     * Set show summary below
     *
     * @param boolean $pValue    Show summary below (true/false)
     * @return PHPExcel\Worksheet
     */
    public function setShowSummaryBelow($pValue = true) {
        $this->_showSummaryBelow = $pValue;
        return $this;
    }

    /**
     * Show summary right? (Row/Column outlining)
     *
     * @return boolean
     */
    public function getShowSummaryRight() {
        return $this->_showSummaryRight;
    }

    /**
     * Set show summary right
     *
     * @param boolean $pValue    Show summary right (true/false)
     * @return PHPExcel\Worksheet
     */
    public function setShowSummaryRight($pValue = true) {
        $this->_showSummaryRight = $pValue;
        return $this;
    }

    /**
     * Get comments
     *
     * @return PHPExcel\Comment[]
     */
    public function getComments()
    {
        return $this->_comments;
    }

    /**
     * Set comments array for the entire sheet.
     *
     * @param array of PHPExcel\Comment
     * @return PHPExcel\Worksheet
     */
    public function setComments($pValue = array())
    {
        $this->_comments = $pValue;

        return $this;
    }

    /**
     * Get comment for cell
     *
     * @param string $pCellCoordinate    Cell coordinate to get comment for
     * @return PHPExcel\Comment
     * @throws PHPExcel\Exception
     */
    public function getComment($pCellCoordinate = 'A1')
    {
        // Uppercase coordinate
        $pCellCoordinate = strtoupper($pCellCoordinate);

        if (strpos($pCellCoordinate,':') !== false || strpos($pCellCoordinate,',') !== false) {
            throw new Exception('Cell coordinate string can not be a range of cells.');
        } else if (strpos($pCellCoordinate,'$') !== false) {
            throw new Exception('Cell coordinate string must not be absolute.');
        } else if ($pCellCoordinate == '') {
            throw new Exception('Cell coordinate can not be zero-length string.');
        } else {
            // Check if we already have a comment for this cell.
            // If not, create a new comment.
            if (isset($this->_comments[$pCellCoordinate])) {
                return $this->_comments[$pCellCoordinate];
            } else {
                return $this->_comments[$pCellCoordinate] = new Comment();
            }
        }
    }

    /**
     * Get comment for cell by using numeric cell coordinates
     *
     * @param int $pColumn    Numeric column coordinate of the cell
     * @param int $pRow        Numeric row coordinate of the cell
     * @return PHPExcel\Comment
     */
    public function getCommentByColumnAndRow($pColumn = 0, $pRow = 1)
    {
        return $this->getComment(Cell::stringFromColumnIndex($pColumn) . $pRow);
    }

    /**
     * Get selected cell
     *
     * @deprecated
     * @return string
     */
    public function getSelectedCell()
    {
        return $this->getSelectedCells();
    }

    /**
     * Get active cell
     *
     * @return string Example: 'A1'
     */
    public function getActiveCell()
    {
        return $this->_activeCell;
    }

    /**
     * Get selected cells
     *
     * @return string
     */
    public function getSelectedCells()
    {
        return $this->_selectedCells;
    }

    /**
     * Selected cell
     *
     * @param    string        $pCoordinate    Cell (i.e. A1)
     * @return PHPExcel\Worksheet
     */
    public function setSelectedCell($pCoordinate = 'A1')
    {
        return $this->setSelectedCells($pCoordinate);
    }

    /**
     * Select a range of cells.
     *
     * @param    string        $pCoordinate    Cell range, examples: 'A1', 'B2:G5', 'A:C', '3:6'
     * @throws    PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function setSelectedCells($pCoordinate = 'A1')
    {
        // Uppercase coordinate
        $pCoordinate = strtoupper($pCoordinate);

        // Convert 'A' to 'A:A'
        $pCoordinate = preg_replace('/^([A-Z]+)$/', '${1}:${1}', $pCoordinate);

        // Convert '1' to '1:1'
        $pCoordinate = preg_replace('/^([0-9]+)$/', '${1}:${1}', $pCoordinate);

        // Convert 'A:C' to 'A1:C1048576'
        $pCoordinate = preg_replace('/^([A-Z]+):([A-Z]+)$/', '${1}1:${2}1048576', $pCoordinate);

        // Convert '1:3' to 'A1:XFD3'
        $pCoordinate = preg_replace('/^([0-9]+):([0-9]+)$/', 'A${1}:XFD${2}', $pCoordinate);

        if (strpos($pCoordinate,':') !== false || strpos($pCoordinate,',') !== false) {
            list($first, ) = Cell::splitRange($pCoordinate);
            $this->_activeCell = $first[0];
        } else {
            $this->_activeCell = $pCoordinate;
        }
        $this->_selectedCells = $pCoordinate;
        return $this;
    }

    /**
     * Selected cell by using numeric cell coordinates
     *
     * @param int $pColumn Numeric column coordinate of the cell
     * @param int $pRow Numeric row coordinate of the cell
     * @throws PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function setSelectedCellByColumnAndRow($pColumn = 0, $pRow = 1)
    {
        return $this->setSelectedCells(Cell::stringFromColumnIndex($pColumn) . $pRow);
    }

    /**
     * Get right-to-left
     *
     * @return boolean
     */
    public function getRightToLeft() {
        return $this->_rightToLeft;
    }

    /**
     * Set right-to-left
     *
     * @param boolean $value    Right-to-left true/false
     * @return PHPExcel\Worksheet
     */
    public function setRightToLeft($value = false) {
        $this->_rightToLeft = $value;
        return $this;
    }

    /**
     * Fill worksheet from values in array
     *
     * @param array $source Source array
     * @param mixed $nullValue Value in source array that stands for blank cell
     * @param string $startCell Insert array starting from this cell address as the top left coordinate
     * @param boolean $strictNullComparison Apply strict comparison when testing for null values in the array
     * @throws PHPExcel\Exception
     * @return PHPExcel\Worksheet
     */
    public function fromArray($source = null, $nullValue = null, $startCell = 'A1', $strictNullComparison = false) {
        if (is_array($source)) {
            //    Convert a 1-D array to 2-D (for ease of looping)
            if (!is_array(end($source))) {
                $source = array($source);
            }

            // start coordinate
            list ($startColumn, $startRow) = Cell::coordinateFromString($startCell);

            // Loop through $source
            foreach ($source as $rowData) {
                $currentColumn = $startColumn;
                foreach($rowData as $cellValue) {
                    if ($strictNullComparison) {
                        if ($cellValue !== $nullValue) {
                            // Set cell value
                            $this->getCell($currentColumn . $startRow)->setValue($cellValue);
                        }
                    } else {
                        if ($cellValue != $nullValue) {
                            // Set cell value
                            $this->getCell($currentColumn . $startRow)->setValue($cellValue);
                        }
                    }
                    ++$currentColumn;
                }
                ++$startRow;
            }
        } else {
            throw new Exception("Parameter \$source should be an array.");
        }
        return $this;
    }

    /**
     * Create array from a range of cells
     *
     * @param string $pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @param mixed $nullValue Value returned in the array entry if a cell doesn't exist
     * @param boolean $calculateFormulas Should formulas be calculated?
     * @param boolean $formatData Should formatting be applied to cell values?
     * @param boolean $returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                               True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     */
    public function rangeToArray($pRange = 'A1', $nullValue = null, $calculateFormulas = true, $formatData = true, $returnCellRef = false) {
        // Returnvalue
        $returnValue = array();
        //    Identify the range that we need to extract from the worksheet
        list($rangeStart, $rangeEnd) = Cell::rangeBoundaries($pRange);
        $minCol = Cell::stringFromColumnIndex($rangeStart[0] -1);
        $minRow = $rangeStart[1];
        $maxCol = Cell::stringFromColumnIndex($rangeEnd[0] -1);
        $maxRow = $rangeEnd[1];

        $maxCol++;
        // Loop through rows
        $r = -1;
        for ($row = $minRow; $row <= $maxRow; ++$row) {
            $rRef = ($returnCellRef) ? $row : ++$r;
            $c = -1;
            // Loop through columns in the current row
            for ($col = $minCol; $col != $maxCol; ++$col) {
                $cRef = ($returnCellRef) ? $col : ++$c;
                //    Using getCell() will create a new cell if it doesn't already exist. We don't want that to happen
                //        so we test and retrieve directly against _cellCollection
                if ($this->_cellCollection->isDataSet($col.$row)) {
                    // Cell exists
                    $cell = $this->_cellCollection->getCacheData($col.$row);
                    if ($cell->getValue() !== null) {
                        if ($cell->getValue() instanceof RichText) {
                            $returnValue[$rRef][$cRef] = $cell->getValue()->getPlainText();
                        } else {
                            if ($calculateFormulas) {
                                $returnValue[$rRef][$cRef] = $cell->getCalculatedValue();
                            } else {
                                $returnValue[$rRef][$cRef] = $cell->getValue();
                            }
                        }

                        if ($formatData) {
                            $style = $this->_parent->getCellXfByIndex($cell->getXfIndex());
                            $returnValue[$rRef][$cRef] = Style_NumberFormat::toFormattedString(
                                $returnValue[$rRef][$cRef],
                                ($style->getNumberFormat()) ?
                                    $style->getNumberFormat()->getFormatCode() :
                                    Style_NumberFormat::FORMAT_GENERAL
                            );
                        }
                    } else {
                        // Cell holds a null
                        $returnValue[$rRef][$cRef] = $nullValue;
                    }
                } else {
                    // Cell doesn't exist
                    $returnValue[$rRef][$cRef] = $nullValue;
                }
            }
        }

        // Return
        return $returnValue;
    }


    /**
     * Create array from a range of cells
     *
     * @param  string $pNamedRange Name of the Named Range
     * @param  mixed  $nullValue Value returned in the array entry if a cell doesn't exist
     * @param  boolean $calculateFormulas  Should formulas be calculated?
     * @param  boolean $formatData  Should formatting be applied to cell values?
     * @param  boolean $returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                                True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     * @throws PHPExcel\Exception
     */
    public function namedRangeToArray($pNamedRange = '', $nullValue = null, $calculateFormulas = true, $formatData = true, $returnCellRef = false) {
        $namedRange = NamedRange::resolveRange($pNamedRange, $this);
        if ($namedRange !== null) {
            $pWorkSheet = $namedRange->getWorksheet();
            $pCellRange = $namedRange->getRange();

            return $pWorkSheet->rangeToArray(    $pCellRange,
                                                $nullValue, $calculateFormulas, $formatData, $returnCellRef);
        }

        throw new Exception('Named Range '.$pNamedRange.' does not exist.');
    }


    /**
     * Create array from worksheet
     *
     * @param mixed $nullValue Value returned in the array entry if a cell doesn't exist
     * @param boolean $calculateFormulas Should formulas be calculated?
     * @param boolean $formatData  Should formatting be applied to cell values?
     * @param boolean $returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                               True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     */
    public function toArray($nullValue = null, $calculateFormulas = true, $formatData = true, $returnCellRef = false) {
        // Garbage collect...
        $this->garbageCollect();

        //    Identify the range that we need to extract from the worksheet
        $maxCol = $this->getHighestColumn();
        $maxRow = $this->getHighestRow();
        // Return
        return $this->rangeToArray(    'A1:'.$maxCol.$maxRow,
                                    $nullValue, $calculateFormulas, $formatData, $returnCellRef);
    }

    /**
     * Get row iterator
     *
     * @param  integer                           $startRow    The row number at which to start iterating
     * @return PHPExcel\Worksheet_RowIterator
     */
    public function getRowIterator($startRow = 1) {
        return new Worksheet_RowIterator($this,$startRow);
    }

    /**
     * Run PHPExcel garabage collector.
     *
     * @return PHPExcel\Worksheet
     */
    public function garbageCollect() {
        // Flush cache
        $this->_cellCollection->getCacheData('A1');
        // Build a reference table from images
//        $imageCoordinates = array();
//        $iterator = $this->getDrawingCollection()->getIterator();
//        while ($iterator->valid()) {
//            $imageCoordinates[$iterator->current()->getCoordinates()] = true;
//
//            $iterator->next();
//        }
//
        // Lookup highest column and highest row if cells are cleaned
        $colRow = $this->_cellCollection->getHighestRowAndColumn();
        $highestRow = $colRow['row'];
        $highestColumn = Cell::columnIndexFromString($colRow['column']);

        // Loop through column dimensions
        foreach ($this->_columnDimensions as $dimension) {
            $highestColumn = max($highestColumn,Cell::columnIndexFromString($dimension->getColumnIndex()));
        }

        // Loop through row dimensions
        foreach ($this->_rowDimensions as $dimension) {
            $highestRow = max($highestRow,$dimension->getRowIndex());
        }

        // Cache values
        if ($highestColumn < 0) {
            $this->_cachedHighestColumn = 'A';
        } else {
            $this->_cachedHighestColumn = Cell::stringFromColumnIndex(--$highestColumn);
        }
        $this->_cachedHighestRow = $highestRow;

        // Return
        return $this;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode() {
        if ($this->_dirty) {
            $this->_hash = md5( $this->_title .
                                $this->_autoFilter .
                                ($this->_protection->isProtectionEnabled() ? 't' : 'f') .
                                __CLASS__
                              );
            $this->_dirty = false;
        }
        return $this->_hash;
    }

    /**
     * Extract worksheet title from range.
     *
     * Example: extractSheetTitle("testSheet!A1") ==> 'A1'
     * Example: extractSheetTitle("'testSheet 1'!A1", true) ==> array('testSheet 1', 'A1');
     *
     * @param string $pRange    Range to extract title from
     * @param bool $returnRange    Return range? (see example)
     * @return mixed
     */
    public static function extractSheetTitle($pRange, $returnRange = false) {
        // Sheet title included?
        if (($sep = strpos($pRange, '!')) === false) {
            return '';
        }

        if ($returnRange) {
            return array( trim(substr($pRange, 0, $sep),"'"),
                          substr($pRange, $sep + 1)
                        );
        }

        return substr($pRange, $sep + 1);
    }

    /**
     * Get hyperlink
     *
     * @param string $pCellCoordinate    Cell coordinate to get hyperlink for
     */
    public function getHyperlink($pCellCoordinate = 'A1')
    {
        // return hyperlink if we already have one
        if (isset($this->_hyperlinkCollection[$pCellCoordinate])) {
            return $this->_hyperlinkCollection[$pCellCoordinate];
        }

        // else create hyperlink
        return $this->_hyperlinkCollection[$pCellCoordinate] = new Cell_Hyperlink();
    }

    /**
     * Set hyperlnk
     *
     * @param string $pCellCoordinate    Cell coordinate to insert hyperlink
     * @param    PHPExcel\Cell_Hyperlink    $pHyperlink
     * @return PHPExcel\Worksheet
     */
    public function setHyperlink($pCellCoordinate = 'A1', Cell_Hyperlink $pHyperlink = null)
    {
        if ($pHyperlink === null) {
            unset($this->_hyperlinkCollection[$pCellCoordinate]);
        } else {
            $this->_hyperlinkCollection[$pCellCoordinate] = $pHyperlink;
        }
        return $this;
    }

    /**
     * Hyperlink at a specific coordinate exists?
     *
     * @param string $pCoordinate
     * @return boolean
     */
    public function hyperlinkExists($pCoordinate = 'A1')
    {
        return isset($this->_hyperlinkCollection[$pCoordinate]);
    }

    /**
     * Get collection of hyperlinks
     *
     * @return PHPExcel\Cell_Hyperlink[]
     */
    public function getHyperlinkCollection()
    {
        return $this->_hyperlinkCollection;
    }

    /**
     * Get data validation
     *
     * @param string $pCellCoordinate Cell coordinate to get data validation for
     */
    public function getDataValidation($pCellCoordinate = 'A1')
    {
        // return data validation if we already have one
        if (isset($this->_dataValidationCollection[$pCellCoordinate])) {
            return $this->_dataValidationCollection[$pCellCoordinate];
        }

        // else create data validation
        $this->_dataValidationCollection[$pCellCoordinate] = new Cell_DataValidation();
        return $this->_dataValidationCollection[$pCellCoordinate];
    }

    /**
     * Set data validation
     *
     * @param string $pCellCoordinate    Cell coordinate to insert data validation
     * @param    PHPExcel\Cell_DataValidation    $pDataValidation
     * @return PHPExcel\Worksheet
     */
    public function setDataValidation($pCellCoordinate = 'A1', Cell_DataValidation $pDataValidation = null)
    {
        if ($pDataValidation === null) {
            unset($this->_dataValidationCollection[$pCellCoordinate]);
        } else {
            $this->_dataValidationCollection[$pCellCoordinate] = $pDataValidation;
        }
        return $this;
    }

    /**
     * Data validation at a specific coordinate exists?
     *
     * @param string $pCoordinate
     * @return boolean
     */
    public function dataValidationExists($pCoordinate = 'A1')
    {
        return isset($this->_dataValidationCollection[$pCoordinate]);
    }

    /**
     * Get collection of data validations
     *
     * @return PHPExcel\Cell_DataValidation[]
     */
    public function getDataValidationCollection()
    {
        return $this->_dataValidationCollection;
    }

    /**
     * Accepts a range, returning it as a range that falls within the current highest row and column of the worksheet
     *
     * @param string $range
     * @return string Adjusted range value
     */
    public function shrinkRangeToFit($range) {
        $maxCol = $this->getHighestColumn();
        $maxRow = $this->getHighestRow();
        $maxCol = Cell::columnIndexFromString($maxCol);

        $rangeBlocks = explode(' ',$range);
        foreach ($rangeBlocks as &$rangeSet) {
            $rangeBoundaries = Cell::getRangeBoundaries($rangeSet);

            if (Cell::columnIndexFromString($rangeBoundaries[0][0]) > $maxCol) { $rangeBoundaries[0][0] = Cell::stringFromColumnIndex($maxCol); }
            if ($rangeBoundaries[0][1] > $maxRow) { $rangeBoundaries[0][1] = $maxRow; }
            if (Cell::columnIndexFromString($rangeBoundaries[1][0]) > $maxCol) { $rangeBoundaries[1][0] = Cell::stringFromColumnIndex($maxCol); }
            if ($rangeBoundaries[1][1] > $maxRow) { $rangeBoundaries[1][1] = $maxRow; }
            $rangeSet = $rangeBoundaries[0][0].$rangeBoundaries[0][1].':'.$rangeBoundaries[1][0].$rangeBoundaries[1][1];
        }
        unset($rangeSet);
        $stRange = implode(' ',$rangeBlocks);

        return $stRange;
    }

    /**
     * Get tab color
     *
     * @return PHPExcel\Style_Color
     */
    public function getTabColor()
    {
        if ($this->_tabColor === null)
            $this->_tabColor = new Style_Color();

        return $this->_tabColor;
    }

    /**
     * Reset tab color
     *
     * @return PHPExcel\Worksheet
     */
    public function resetTabColor()
    {
        $this->_tabColor = null;
        unset($this->_tabColor);

        return $this;
    }

    /**
     * Tab color set?
     *
     * @return boolean
     */
    public function isTabColorSet()
    {
        return ($this->_tabColor !== null);
    }

    /**
     * Copy worksheet (!= clone!)
     *
     * @return PHPExcel\Worksheet
     */
    public function copy() {
        $copied = clone $this;

        return $copied;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone() {
        foreach ($this as $key => $val) {
            if ($key == '_parent') {
                continue;
            }

            if (is_object($val) || (is_array($val))) {
                if ($key == '_cellCollection') {
                    $newCollection = clone $this->_cellCollection;
                    $newCollection->copyCellCollection($this);
                    $this->_cellCollection = $newCollection;
                } elseif ($key == '_drawingCollection') {
                    $newCollection = clone $this->_drawingCollection;
                    $this->_drawingCollection = $newCollection;
                } elseif (($key == '_autoFilter') && ($this->_autoFilter instanceof Worksheet_AutoFilter)) {
                    $newAutoFilter = clone $this->_autoFilter;
                    $this->_autoFilter = $newAutoFilter;
                    $this->_autoFilter->setParent($this);
                } else {
                    $this->{$key} = unserialize(serialize($val));
                }
            }
        }
    }
}
