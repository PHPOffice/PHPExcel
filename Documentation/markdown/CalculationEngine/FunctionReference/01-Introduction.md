# Calculation Engine - Formula Function Reference

## Frequently asked questions

The up-to-date F.A.Q. page for PHPExcel can be found on [https://github.com/PHPOffice/PHPExcel/Wiki/View.aspx?title=FAQ&referringTitle=Requirements][1].

### Formulas don’t seem to be calculated in Excel2003 using compatibility pack?

This is normal behaviour of the compatibility pack, Excel2007 displays this correctly. Use PHPExcel_Writer_Excel5 if you really need calculated values, or force recalculation in Excel2003.

  [1]: https://github.com/PHPOffice/PHPExcel/Wiki/View.aspx?title=FAQ&referringTitle=Requirements
