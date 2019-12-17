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
	public class ExcelSheet: ExcelWorkbook
	{
		public Worksheet Worksheet { get; set; }

		public ExcelSheet(string fileName, string sheetName): base(fileName)
		{
			if (SheetExists(sheetName))
			{
				Worksheet = Worksheets[sheetName];
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

		public DataTable WriteExcelSheetToDataTable()
		{
			Range range = GetUsedRange();
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
		public void WriteDatatableToRange(int row, int col, DataTable dataTable) => WriteArrayToExcelSheet(row, col, dataTable.ToArray());
		public void WriteArrayToExcelSheet(int row, int col, object[,] dataArr)
		{
			Range range = SetRange(row, col, dataArr.GetLength(0) + row - 1, dataArr.GetLength(1) + col - 1);
			range.set_Value(Type.Missing, dataArr);
			Cleanup.ReleaseObject(range);
		}

		#region Private Methods
		private bool SheetExists(string sheetName)
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
}
