<?php

class PHPExcel_Adapter_Spreadsheet_Excel_Writer_ValueBinder implements PHPExcel_Cell_IValueBinder
{
	public function bindValue(PHPExcel_Cell $cell, $value = null)
	{
		$cell->setValueExplicit($value, $this->dataTypeForValue($value));
		return true;
	}

	public static function dataTypeForValue($pValue = null)
	{
		if ($pValue === null) {
			$dataType = PHPExcel_Cell_DataType::TYPE_NULL;

		} elseif ($pValue === '') {
			$dataType = PHPExcel_Cell_DataType::TYPE_STRING;

		} elseif ($pValue{0} === '0' && strlen($pValue) > 1 && $pValue{1} !== '.') {
			$dataType = PHPExcel_Cell_DataType::TYPE_STRING;

		} elseif (is_numeric($pValue)) {
			$dataType = PHPExcel_Cell_DataType::TYPE_NUMERIC;

		} elseif ($pValue{0} === '=' && strlen($pValue) > 1) {
			$dataType = PHPExcel_Cell_DataType::TYPE_FORMULA;

		} elseif (is_string($pValue)) {
			$dataType = PHPExcel_Cell_DataType::TYPE_STRING;

		} elseif (is_bool($pValue)) {
			$dataType = PHPExcel_Cell_DataType::TYPE_BOOL;

		} elseif ($pValue instanceof PHPExcel_RichText) {
			$dataType = PHPExcel_Cell_DataType::TYPE_STRING;

		} else {
			$dataType = PHPExcel_Cell_DataType::TYPE_STRING;

		}

		/* if (preg_match("/^0\d+$/", $pValue)) { */
		/*	   $dataType = PHPExcel_Cell_DataType::TYPE_STRING; */

		/* } elseif (is_numeric($pValue)) { */
		/*	   $dataType = PHPExcel_Cell_DataType::TYPE_NUMERIC; */

		/* } else { */
		/*	   $dataType = PHPExcel_Cell_DefaultValueBinder::dataTypeForValue($pValue); */

		/* } */

		return $dataType;
	}
}
