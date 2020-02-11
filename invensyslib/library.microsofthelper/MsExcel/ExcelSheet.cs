using library.common;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Runtime.InteropServices;
using DataTable = System.Data.DataTable;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace library.microsofthelper.MsExcel
{
	public class ExcelSheet : ExcelWorkbook
	{
		public Worksheet Worksheet { get; set; }
		public SheetData Data { get; set; }
		public SheetFormatting Formatting { get; set; }
		public SheetDataProcessor DataProcessor { get; set; }

		public ExcelSheet(string fileName, string sheetName) : base(fileName)
		{
			Data = new SheetData(this);
			Formatting = new SheetFormatting(this);
			DataProcessor = new SheetDataProcessor(this);

			if (SheetExists(sheetName))
			{
				Worksheet = Worksheets[sheetName];
				return;
			}
			else if (SheetExists(sheetName.ToUpper()))
			{
				Worksheet = Worksheets[sheetName.ToUpper()];
				return;
			}
			else if (SheetExists(sheetName.ToLower()))
			{
				Worksheet = Worksheets[sheetName.ToLower()];
				return;
			}

			Worksheet = Worksheets.Add();
			Worksheet.Name = sheetName;
		}
		public void Delete() => Worksheet.Delete();

		public Range SetRange(int row1, int column1, int row2, int column2)
		{
			Range cell1 = (Range)Worksheet.Cells[row1, column1];
			Range cell2 = (Range)Worksheet.Cells[row2, column2];
			Range retRange = Worksheet.Range[cell1, cell2];
			Cleanup.ReleaseObject(cell1);
			Cleanup.ReleaseObject(cell2);
			return retRange;
		}
		public Range GetUsedRange()
		{
			int lastRow = 1;
			int lastColInitial = FindLastUsedColumn();
			int lastCol = 1;
			for (int i = 1; i < lastColInitial; i++)
			{
				int tmpRow = FindLastUsedRow(i);
				lastRow = (lastRow < tmpRow) ? tmpRow : lastRow;
			}
			for (int i = 1; i < lastRow; i++)
			{
				int tmpCol = FindLastUsedColumn(i);
				lastCol = (lastCol < tmpCol) ? tmpCol : lastCol;
			}

			return SetRange(1, 1, lastRow, lastCol);
		}
		public int FindLastUsedRow(int searchColumn = 1)
		{
			Range searchRange = SetRange(1, searchColumn, Worksheet.Rows.Count, searchColumn);
			Range lastCell = searchRange.Find(What: "*", After: Worksheet.Cells[1, searchColumn], LookIn: XlFindLookIn.xlFormulas, LookAt: XlLookAt.xlWhole, SearchDirection: XlSearchDirection.xlPrevious, MatchCase: true);
			int retInt = lastCell == null ? 1 : lastCell.Row;
			Cleanup.ReleaseObject(lastCell);
			Cleanup.ReleaseObject(searchRange);
			return retInt;
		}
		public int FindLastUsedColumn(int searchRow = 1)
		{
			Range searchRange = SetRange(searchRow, 1, searchRow, Worksheet.Columns.Count);
			Range lastCell = searchRange.Find(What: "*", After: Worksheet.Cells[searchRow, 1], LookIn: XlFindLookIn.xlFormulas, LookAt: XlLookAt.xlWhole, SearchDirection: XlSearchDirection.xlPrevious, MatchCase: true);
			int retInt = lastCell == null ? 1 : lastCell.Column;
			Cleanup.ReleaseObject(lastCell);
			Cleanup.ReleaseObject(searchRange);
			return retInt;
		}
		public int FindDataColumn(string findString, int searchRow = 1, int startColumn = 1)
		{
			if (string.IsNullOrEmpty(findString))
				return 0;
			Range searchRange = SetRange(searchRow, 1, searchRow, Worksheet.Columns.Count);
			Range foundCell = searchRange.Find(What: findString, After: Worksheet.Cells[searchRow, startColumn], LookIn: XlFindLookIn.xlFormulas, LookAt: XlLookAt.xlWhole, SearchDirection: XlSearchDirection.xlPrevious, MatchCase: true);
			int retInt = foundCell == null ? 0 : foundCell.Column;
			Cleanup.ReleaseObject(foundCell);
			Cleanup.ReleaseObject(searchRange);
			return retInt;
		}
		public int FindDataRow(string findString, int searchColumn = 1, int startRow = 1)
		{
			if (string.IsNullOrEmpty(findString))
				return 0;
			Range searchRange = SetRange(1, searchColumn, Worksheet.Rows.Count, searchColumn);
			Range foundCell = searchRange.Find(What: findString, After: Worksheet.Cells[startRow, searchColumn], LookIn: XlFindLookIn.xlFormulas, LookAt: XlLookAt.xlWhole, SearchDirection: XlSearchDirection.xlPrevious, MatchCase: true);
			int retInt = foundCell == null ? 0 : foundCell.Row;
			Cleanup.ReleaseObject(foundCell);
			Cleanup.ReleaseObject(searchRange);
			return retInt;
		}
		public int[] FindDataRows(string findString, int searchColumn = 1)
		{
			List<int> retList = new List<int>();
			if (string.IsNullOrEmpty(findString))
				return null;

			Range searchRange = SetRange(1, searchColumn, Worksheet.Rows.Count, searchColumn);
			Range currentFind = searchRange.Find(What: findString, LookIn: XlFindLookIn.xlFormulas, LookAt: XlLookAt.xlWhole, SearchDirection: XlSearchDirection.xlNext, MatchCase: true);
			Range firstFind = null;

			while (currentFind != null)
			{
				if (firstFind == null)
					firstFind = currentFind;

				retList.Add(currentFind.Row);
				currentFind = searchRange.FindNext(currentFind);

				if (currentFind.get_Address() == firstFind.get_Address())
					break;
			}
			Cleanup.ReleaseObject(currentFind);
			Cleanup.ReleaseObject(firstFind);
			Cleanup.ReleaseObject(searchRange);

			return retList.ToArray();
		}
		public int[] FindDataColumns(string findString, int searchRow = 1)
		{
			List<int> retList = new List<int>();
			if (string.IsNullOrEmpty(findString))
				return null;

			Range searchRange = SetRange(searchRow, 1, searchRow, Worksheet.Columns.Count);
			Range currentFind = searchRange.Find(What: findString, LookIn: XlFindLookIn.xlFormulas, LookAt: XlLookAt.xlWhole, SearchDirection: XlSearchDirection.xlNext, MatchCase: true);
			Range firstFind = null;

			while (currentFind != null)
			{
				if (firstFind == null)
					firstFind = currentFind;

				retList.Add(currentFind.Column);
				currentFind = searchRange.FindNext(currentFind);

				if (currentFind.get_Address() == firstFind.get_Address())
					break;
			}
			Cleanup.ReleaseObject(currentFind);
			Cleanup.ReleaseObject(firstFind);
			Cleanup.ReleaseObject(searchRange);

			return retList.ToArray();
		}

		public void DeleteColumn(int columnIndex) => Worksheet.Columns[columnIndex].Delete();
		public void DeleteRow(int rowIndex) => Worksheet.Rows[rowIndex].Delete();
		public void InsertColumn(int columnIndex) => Worksheet.Columns[columnIndex].Insert();
		public void InsertRow(int rowIndex) => Worksheet.Rows[rowIndex].Insert();
		#region Private Methods
		public bool SheetExists(string sheetName)
		{
			foreach (Worksheet sht in Worksheets)
			{
				if (sht.Name == sheetName)
				{
					Cleanup.ReleaseObject(sht);
					return true;
				}
				Cleanup.ReleaseObject(sht);
			}
			return false;
		}
		#endregion

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls
		protected override void Dispose(bool disposing)
		{
			if (!disposedValue)
			{
				if (disposing)
				{
					GC.Collect();
					Cleanup.ReleaseObject(Worksheet);
					Cleanup.ReleaseObject(Worksheets);
					if (Workbook != null)
						Workbook.Close(0);
					Cleanup.ReleaseObject(Workbook);
					Cleanup.ReleaseObject(ExcelWorkbooks);
					ExcelApplication.Quit();
					Cleanup.ReleaseObject(ExcelApplication);
					GC.WaitForPendingFinalizers();
				}
				disposedValue = true;
			}
		}
		#endregion
	}

	public class SheetFormatting
	{
		private const int COLOR_CELL_HEADING = 13696965;
		internal SheetFormatting(ExcelSheet excelSheet) => ExcelSheet = excelSheet;
		private ExcelSheet ExcelSheet { get; set; }

		public void UseStandardFormatting(int headingRow = 1, int freezeColumn = 0)
		{
			SetJustificationLeft();
			SetFont();
			SetBorders();
			UseStandardTableHeading(headingRow);
			AutoFitAllColumns();
			ApplyFilters();
			FreezePanes(headingRow + 1, freezeColumn + 1);
		}
		public void SetBorders()
		{
			Range range = ExcelSheet.GetUsedRange();
			range.Borders[XlBordersIndex.xlDiagonalDown].LineStyle = XlLineStyle.xlLineStyleNone;
			range.Borders[XlBordersIndex.xlDiagonalUp].LineStyle = XlLineStyle.xlLineStyleNone;

			range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
			range.Borders[XlBordersIndex.xlEdgeBottom].ThemeColor = 1;
			range.Borders[XlBordersIndex.xlEdgeBottom].TintAndShade = -0.499984740745262;
			range.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;

			range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
			range.Borders[XlBordersIndex.xlEdgeTop].ThemeColor = 1;
			range.Borders[XlBordersIndex.xlEdgeTop].TintAndShade = -0.499984740745262;
			range.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;

			range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
			range.Borders[XlBordersIndex.xlEdgeLeft].ThemeColor = 1;
			range.Borders[XlBordersIndex.xlEdgeLeft].TintAndShade = -0.499984740745262;
			range.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;

			range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
			range.Borders[XlBordersIndex.xlEdgeRight].ThemeColor = 1;
			range.Borders[XlBordersIndex.xlEdgeRight].TintAndShade = -0.499984740745262;
			range.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;

			range.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
			range.Borders[XlBordersIndex.xlInsideHorizontal].ThemeColor = 1;
			range.Borders[XlBordersIndex.xlInsideHorizontal].TintAndShade = -0.499984740745262;
			range.Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;

			range.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
			range.Borders[XlBordersIndex.xlInsideVertical].ThemeColor = 1;
			range.Borders[XlBordersIndex.xlInsideVertical].TintAndShade = -0.499984740745262;
			range.Borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlThin;
			Cleanup.ReleaseObject(range);
		}
		public void FilTableColours(bool hasColour)
		{
			Range range = ExcelSheet.GetUsedRange();
			switch (hasColour)
			{
				case true:
					range.Interior.Pattern = XlPattern.xlPatternSolid;
					range.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
					range.Interior.Color = 15532007;
					range.Interior.TintAndShade = 0;
					range.Interior.PatternTintAndShade = 0;
					break;
				case false:
					range.Interior.Pattern = XlPattern.xlPatternSolid;
					range.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
					range.Interior.ThemeColor = XlThemeColor.xlThemeColorDark1;
					range.Interior.TintAndShade = 0;
					range.Interior.PatternTintAndShade = 0;
					break;
			}
			Cleanup.ReleaseObject(range);
		}
		public void AutoFitAllColumns()
		{
			Range rangeA = ExcelSheet.Worksheet.Cells;
			Range rangeB = rangeA.EntireColumn;
			rangeB.AutoFit();

			for (int i = 1; i <= ExcelSheet.FindLastUsedColumn(); i++)
			{
				Range range = (Range)ExcelSheet.Worksheet.Columns[i];
				range.ColumnWidth += 3;
				Cleanup.ReleaseObject(range);
			}
			Cleanup.ReleaseObject(rangeA);
			Cleanup.ReleaseObject(rangeB);
		}
		public void FreezePanes(int freezeRow = 1, int freezeColumn = 1)
		{
			ExcelSheet.Workbook.Activate();
			ExcelSheet.Worksheet.Activate();
			Range range = (Range)ExcelSheet.Worksheet.Cells[freezeRow, freezeColumn];
			range.Select();
			ExcelSheet.ExcelApplication.ActiveWindow.FreezePanes = true;
			Cleanup.ReleaseObject(range);
		}
		public void ApplyFilters(int filterRow = 1, int filterColumn = 1)
		{
			Range range = ExcelSheet.SetRange(filterRow, filterColumn, filterRow, filterColumn);
			range.AutoFilter(2);
			Cleanup.ReleaseObject(range);
		}
		public void SetFont()
		{
			Range range = ExcelSheet.GetUsedRange();
			range.Font.Name = "Calibri";
			range.Font.Size = 10;
			range.Font.Strikethrough = false;
			range.Font.Superscript = false;
			range.Font.Subscript = false;
			range.Font.OutlineFont = false;
			range.Font.Shadow = false;
			range.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
			range.Font.ThemeColor = XlThemeColor.xlThemeColorLight1;
			range.Font.TintAndShade = 0;
			range.Font.ThemeFont = XlThemeFont.xlThemeFontMinor;
			Cleanup.ReleaseObject(range);
		}
		public void SetJustificationLeft()
		{
			Range range = ExcelSheet.GetUsedRange();
			range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			range.VerticalAlignment = XlVAlign.xlVAlignCenter;
			Cleanup.ReleaseObject(range);
		}
		public void UseStandardTableHeading(int headingRow = 1)
		{
			Range range = ExcelSheet.SetRange(headingRow, 1, headingRow, ExcelSheet.FindLastUsedColumn(headingRow));
			range.Interior.Pattern = XlPattern.xlPatternSolid;
			range.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
			range.Interior.Color = COLOR_CELL_HEADING;
			range.Interior.TintAndShade = 0;
			range.Interior.PatternTintAndShade = 0;
			range.Font.Bold = true;
			range.WrapText = true;
			Cleanup.ReleaseObject(range);
		}
		public void MergeAndCentre(int row1, int column1, int row2, int column2)
		{
			Range rng = ExcelSheet.SetRange(row1, column1, row2, column2);
			rng.Merge();
			rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			rng.VerticalAlignment = XlVAlign.xlVAlignCenter;
			Cleanup.ReleaseObject(rng);
		}
		public void SetNumberFormatAccounting(int EmployeeCodeColumn, int decimalPlaces = 2, Range range = null)
		{
			if (range == null)
				range = ExcelSheet.GetUsedRange();

			if (range == null)
				return;

			foreach (Range cell in range.Cells)
			{
				if (cell.Column == EmployeeCodeColumn)
					continue;

				object val = cell.Value;
				if (val == null)
					continue;

				string typeName = val.GetType().ToString();
				switch (typeName)
				{
					case "System.Double":
						string zeros = GetNumZeros(decimalPlaces);
						cell.NumberFormat = "_(* #,##0." + zeros + "_);_(* -#,##0." + zeros + ";_(* \" - \"??_);_ @_ ";
						break;
					default:
						cell.NumberFormat = "@";
						break;
				}

				cell.Value = cell.Value;
			}
		}
		private static string GetNumZeros(int decimalPlaces)
		{
			string zeros = "";
			for (int i = 0; i < decimalPlaces; i++)
			{
				zeros += "0";
			}
			if (zeros == "")
				zeros = "00";
			return zeros;
		}
		public void SetNumberFormatColumn(int columnToFormat, string format)
		{
			Range x = ExcelSheet.SetRange(1, columnToFormat, ExcelSheet.FindLastUsedRow(columnToFormat), columnToFormat);
			x.NumberFormat = format;
			foreach (Range cell in x)
			{
				cell.Value = cell.Value;
			}
		}
	}

	public class SheetData
	{
		internal SheetData(ExcelSheet excelSheet) => ExcelSheet = excelSheet;
		private ExcelSheet ExcelSheet { get; set; }

		/// <summary>
		/// Convert Excel data sheet into DataTable
		/// Sheet Must have clearly defined Column names, which translate to datatable columns
		/// </summary>
		/// <returns>DataTable</returns>
		public DataTable WriteExcelSheetToDataTable()
		{
			Range range = ExcelSheet.GetUsedRange();
			Range tempRange = null;
			DataTable dt = new DataTable();
			//Add Headers
			for (int i = 1; i <= range.Columns.Count; i++)
			{
				tempRange = ((Range)range.Cells[1, i]);
				dynamic colname = tempRange.Value;
				Cleanup.ReleaseObject(tempRange);
				dt.Columns.Add(colname);
			}
			//Body
			for (int j = 2; j <= range.Rows.Count; j++)
			{
				DataRow dr = dt.NewRow();
				for (int i = 1; i <= range.Columns.Count; i++)
				{
					tempRange = ((Range)range.Cells[j, i]);
					dr[i - 1] = tempRange.Value;
					Cleanup.ReleaseObject(tempRange);
				}
				dt.Rows.Add(dr);
			}
			Cleanup.ReleaseObject(range);
			return dt;
		}

		/// <summary>
		/// Writes DataTable to Excel sheet, at specified Cell location
		/// </summary>
		/// <returns>void</returns>
		public void WriteDatatableToRange(int row, int col, DataTable dataTable) => WriteArrayToExcelSheet(row, col, dataTable.ToArray());

		/// <summary>
		/// Writes @D Array of objects to Excel sheet, at specified Cell location
		/// </summary>
		/// <returns>void</returns>
		public void WriteArrayToExcelSheet(int row, int col, object[,] dataArr)
		{
			Range range = ExcelSheet.SetRange(row, col, dataArr.GetLength(0) + row - 1, dataArr.GetLength(1) + col - 1);
			range.set_Value(Type.Missing, dataArr);
			Cleanup.ReleaseObject(range);
		}
	}

	[ComVisible(true)]
	[ClassInterface(ClassInterfaceType.None)]
	public class SheetDataProcessor
	{
		internal SheetDataProcessor(ExcelSheet excelSheet) => ExcelSheet = excelSheet;
		private ExcelSheet ExcelSheet { get; set; }
		public void DeleteZeroSumColumns(int startColumn, int endColumn, int startRow = 1, int endRow = 0)
		{
			for (int i = endColumn; i >= startColumn; i--)
			{
				if (endRow == 0)
					endRow = ExcelSheet.FindLastUsedRow(i);

				if (!AnyValueExistsInRange(ExcelSheet.SetRange(startRow, i, endRow, i)))
					ExcelSheet.DeleteColumn(i);
			}
		}
		public void DeleteZeroSumRows(int startRow, int endRow, int startColumn = 1, int endColumn = 0)
		{
			if (endRow == 1)
				return;

			for (int i = endRow; i >= startRow; i--)
			{
				if (endColumn == 0)
					endColumn = ExcelSheet.FindLastUsedColumn(i);

				if (!AnyValueExistsInRange(ExcelSheet.SetRange(i, startColumn, i, endColumn)))
					ExcelSheet.DeleteRow(i);
			}
		}
		public void GetConsecutiveColumnsAsArray(int fromColumn, int toColumn, ref int[] columnNumbers)
		{
			int arrSize = toColumn - fromColumn + 1;
			int[] tempArray = new int[arrSize];
			for (int i = 0; i < arrSize; i++)
			{
				tempArray[i] = i + fromColumn;
			}

			columnNumbers = tempArray;
		}
		public void SortSheet(ref int[] columnNumbers)
		{
			int lRow = ExcelSheet.FindLastUsedRow(columnNumbers[0]);

			ExcelSheet.Worksheet.Sort.SortFields.Clear();
			for (int i = 0; i < columnNumbers.GetLength(0); i++)
			{
				ExcelSheet.Worksheet.Sort.SortFields.Add(ExcelSheet.SetRange(1, columnNumbers[i], lRow, columnNumbers[i]), XlSortOn.xlSortOnValues, XlSortOrder.xlAscending, XlSortDataOption.xlSortTextAsNumbers);
			}
			ExcelSheet.Worksheet.Sort.SetRange(ExcelSheet.SetRange(1, 1, lRow, ExcelSheet.FindLastUsedColumn()));
			ExcelSheet.Worksheet.Sort.Header = XlYesNoGuess.xlYes;
			ExcelSheet.Worksheet.Sort.MatchCase = false;
			ExcelSheet.Worksheet.Sort.SortMethod = XlSortMethod.xlPinYin;
			ExcelSheet.Worksheet.Sort.Apply();
		}
		public void SubTotalSheet(int columnToGroupBy, ref int[] columnsToApplyTotalsTO) => ExcelSheet.Worksheet.UsedRange.Subtotal(columnToGroupBy, XlConsolidationFunction.xlSum, columnsToApplyTotalsTO, false, false, XlSummaryRow.xlSummaryBelow);
		public void SubTotalThenRemoveDetails(int columnWithSubtotal, int fromColumnToAddTotalsTo, int toColumnToAddTotalsTo)
		{
			int[] columnsWithTotals = new int[1];
			GetConsecutiveColumnsAsArray(fromColumnToAddTotalsTo, toColumnToAddTotalsTo, ref columnsWithTotals);

			int lastRow = ExcelSheet.FindLastUsedRow(columnWithSubtotal);
			int delRow = lastRow + 10;
			int valFromCol = columnsWithTotals[0];
			int valToCol = columnsWithTotals[columnsWithTotals.Length - 1];

			for (int i = lastRow; i >= 2; i--)
			{
				if (GetCellValueAsString(i, 1) != "" && GetCellValueAsString(i, 1) != null)
				{
					Range xRng = ExcelSheet.SetRange(i + 1, valFromCol, i + 1, valToCol);
					dynamic val = xRng.Value;
					while (GetCellValueAsString(i, 1) != "" && GetCellValueAsString(i, 1) != null)
					{
						if (i == 1)
							break;

						i--;
					}

					ExcelSheet.Worksheet.Rows[delRow + ":" + (i + 2)].Delete(XlDeleteShiftDirection.xlShiftUp);
					ExcelSheet.SetRange(i + 1, valFromCol, i + 1, valToCol).Value = val;
					delRow = i;
				}
			}
		}
		public void SortThenSubtotalSheet(ref int[] columnNumbersToSortOn, ref int[] columnNumbersToGroupSubtotalBy, int fromColumnToAddTotalsTo, int toColumnToAddTotalsTo)
		{
			SortSheet(ref columnNumbersToSortOn);
			int[] columnNumbers = new int[1];
			GetConsecutiveColumnsAsArray(fromColumnToAddTotalsTo, toColumnToAddTotalsTo, ref columnNumbers);
			for (int i = 0; i < columnNumbersToGroupSubtotalBy.Length; i++)
			{
				SubTotalSheet(columnNumbersToGroupSubtotalBy[i], ref columnNumbers);
			}
		}
		public double SumRange(Range rangeToSum)
		{
			double sum = 0;
			foreach (Range cell in rangeToSum.Cells)
			{
				try
				{
					sum += cell.Value2;
				}
				catch
				{
					sum += 0;
				}
			}

			return sum;
		}
		public void SumColumn(int startColumn, int endcolumn = 0)
		{
			int tR = ExcelSheet.FindLastUsedRow();
			endcolumn = endcolumn != 0 ? endcolumn : ExcelSheet.FindLastUsedColumn();
			for (int i = startColumn; i <= endcolumn; i++)
			{
				ExcelSheet.Worksheet.Cells[tR + 1, i].Value = SumRange(ExcelSheet.SetRange(1, i, tR, i));
			}
		}
		public object GetLookupValueFromTable(ExcelSheet lookupSheet, string lookupTableName, string lookupValue, int lookupIndex) => ExcelSheet.Worksheet.Application.WorksheetFunction.VLookup(lookupValue, lookupSheet.Worksheet.Range[lookupTableName], lookupIndex, false);
		public string GetLookupValueFromTableUsingValidations(Range Target, ExcelSheet lookupSheet, object[] lookupTableNames, int lookupIndex)
		{
			string retval = "";

			if (lookupTableNames == null || Target == null || Target.Rows.Count != 1)
				return retval;

			string lookupValue = Target.Value;
			string validationFormula = Target.Validation.Formula1.ToString().Replace("=", "");

			foreach (string table in lookupTableNames)
			{
				if (!validationFormula.Contains(table))
					continue;

				retval = GetLookupValueFromTable(lookupSheet, table, lookupValue, lookupIndex).ToString();
			}

			return retval;
		}
		public void ReplaceLookupCodesVertical(ExcelSheet lookupSheet, string lookupAttribute, double checkColumnIndex = 1)
		{
			int col = ExcelSheet.FindDataColumn(lookupAttribute);
			int dtCol = lookupSheet.FindDataColumn(lookupAttribute);
			int lRow = ExcelSheet.FindLastUsedRow();

			if (col == 0 || dtCol == 0)
				return;

			for (int j = 2; j <= lRow; j++)
			{
				dynamic chkVal = ExcelSheet.Worksheet.Cells[j, col].Value.ToString();
				if (chkVal == " " || chkVal == "")
				{
					ExcelSheet.Worksheet.Cells[j, col].Value = "Undefined";
					continue;
				}

				dynamic DTRow = lookupSheet.FindDataRow(chkVal, dtCol);
				if (DTRow == 0)
					continue;

				ExcelSheet.Worksheet.Cells[j, col].Value = lookupSheet.Worksheet.Cells[DTRow, dtCol + checkColumnIndex].Value;
			}
		}
		public string ReplaceLookupCodesHorisontal(ExcelSheet lookupSheet, string lookupAttribute, double checkRowIndex = 1) => lookupSheet.Worksheet.Cells[checkRowIndex + 1, lookupSheet.FindDataColumn(lookupAttribute)].Value;

		public void ReplaceLookupCodesMultiplePerColumn(ExcelSheet lookupSheet, int columnOfAttributesToLookup, int columnToChangeCodeToLookupValue)
		{
			int LRow = ExcelSheet.FindLastUsedRow(columnOfAttributesToLookup);

			for (int i = 2; i <= LRow; i++)
			{
				dynamic searchVal = ExcelSheet.Worksheet.Cells[i, columnOfAttributesToLookup].value;
				dynamic searchLookCol = lookupSheet.FindDataColumn(searchVal);
				if (searchLookCol != 0)
				{
					try
					{
						dynamic loookupVal = ExcelSheet.Worksheet.Cells[i, columnToChangeCodeToLookupValue].Value;
						dynamic searchLookRow = lookupSheet.FindDataRow(loookupVal.ToString(), searchLookCol);
						dynamic tVal = lookupSheet.Worksheet.Cells[searchLookRow, searchLookCol + 1].Value;
						ExcelSheet.Worksheet.Cells[i, columnToChangeCodeToLookupValue].Value = loookupVal + " - " + tVal.ToString();
					}
					catch { }
				}
			}
		}
		public void GenerateHeadingsFromLookupSheet(ExcelSheet lookupSheet, int headingColumn, bool hasHeading = false)
		{
			int ToRow = lookupSheet.FindLastUsedRow(headingColumn);
			int h = hasHeading ? 2 : 1;
			for (int i = h; i <= ToRow; i++)
			{
				ExcelSheet.Worksheet.Cells[1, i - h + 1].Value = lookupSheet.Worksheet.Cells[i, headingColumn].Value;
			}
		}

		private string GetCellValueAsString(int y, int x)
		{
			string testVal;
			try
			{
				double d = ExcelSheet.Worksheet.Cells[y, x].Value;
				testVal = d.ToString();
			}
			catch
			{
				testVal = ExcelSheet.Worksheet.Cells[y, x].Value;
			}

			return testVal;
		}
		private bool AnyValueExistsInRange(Range range)
		{
			foreach (Range cell in range.Cells)
			{
				if (cell.Value != 0)
					return true;
			}
			return false;
		}
	}
}