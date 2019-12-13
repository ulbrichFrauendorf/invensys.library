using library.common;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using DataTable = System.Data.DataTable;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace library.microsofthelper.MsExcel
{
	public class ExcelSheet : IDisposable
	{
		public Worksheet Worksheet { get; set; }
		public ExcelSheet(Workbook workbook, string sheetName)
		{
			if (SheetExists(workbook, sheetName))
			{
				Worksheet = workbook.Worksheets[sheetName];
				return;
			}

			Sheets xlSheets = workbook.Worksheets;
			Worksheet = xlSheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
			Worksheet.Name = sheetName;
		}
		public void Delete() => Worksheet.Delete();

		public Range SetRange(int row1, int column1, int row2, int column2) => Worksheet.Range[(Range)Worksheet.Cells[row1, column1], (Range)Worksheet.Cells[row2, column2]];
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
			return lastCell == null ? 1 : lastCell.Row;
		}
		public int FindLastUsedColumn(int searchRow = 1)
		{
			Range searchRange = SetRange(searchRow, 1, searchRow, Worksheet.Columns.Count);
			Range lastCell = searchRange.Find(What: "*", After: Worksheet.Cells[searchRow, 1], LookIn: XlFindLookIn.xlFormulas, LookAt: XlLookAt.xlWhole, SearchDirection: XlSearchDirection.xlPrevious, MatchCase: true);
			return lastCell == null ? 1 : lastCell.Column;
		}
		public int FindDataColumn(string findString, int searchRow = 1, int startColumn = 1)
		{
			if (findString == "")
				return 0;
			Range searchRange = SetRange(searchRow, 1, searchRow, Worksheet.Columns.Count);
			Range foundCell = searchRange.Find(What: findString, After: Worksheet.Cells[searchRow, startColumn], LookIn: XlFindLookIn.xlFormulas, LookAt: XlLookAt.xlWhole, SearchDirection: XlSearchDirection.xlPrevious, MatchCase: true);
			return foundCell == null ? 0 : foundCell.Column;
		}
		public int FindDataRow(string findString, int searchColumn = 1, int startRow = 1)
		{
			if (findString == "")
				return 0;
			Range searchRange = SetRange(1, searchColumn, Worksheet.Rows.Count, searchColumn);
			Range foundCell = searchRange.Find(What: findString, After: Worksheet.Cells[startRow, searchColumn], LookIn: XlFindLookIn.xlFormulas, LookAt: XlLookAt.xlWhole, SearchDirection: XlSearchDirection.xlPrevious, MatchCase: true);
			return foundCell == null ? 0 : foundCell.Row;
		}
		public int[] FindDataRows(string findString, int searchColumn = 1)
		{
			List<int> retList = new List<int>();
			if (findString == "")
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
			return retList.ToArray();
		}
		public int[] FindDataColumns(string findString, int searchRow = 1)
		{
			List<int> retList = new List<int>();
			if (findString == "")
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
			return retList.ToArray();
		}

		public DataTable WriteExcelSheetToDataTable()
		{
			Range range = GetUsedRange();
			DataTable dt = new DataTable();
			//Add Headers
			for (int i = 1; i <= range.Columns.Count; i++)
			{
				dynamic colname = ((Range)range.Cells[1, i]).Value;
				dt.Columns.Add(colname);
			}
			//Body
			for (int j = 2; j <= range.Rows.Count; j++)
			{
				DataRow dr = dt.NewRow();
				for (int i = 1; i <= range.Columns.Count; i++)
				{
					dr[i - 1] = ((Range)range.Cells[j, i]).Value;
				}
				dt.Rows.Add(dr);
			}
			Marshal.ReleaseComObject(range);
			range = null;
			return dt;
		}
		public void WriteDatatableToRange(int row, int col, DataTable dataTable) => WriteArrayToRange(row, col, dataTable.ToArray());
		public void WriteArrayToRange(int row, int col, object[,] dataArr) => SetRange(row, col, dataArr.GetLength(0) + row - 1, dataArr.GetLength(1) + col - 1).set_Value(Type.Missing, dataArr);

		//public string SaveAsWorkbook(string saveName, bool savePopupFlag = true, string password = "")
		//{
		//	Workbook wb = Worksheet.Application.Workbooks.Add();
		//	wb.Application.DisplayAlerts = false;
		//	Worksheet.Copy(Before: wb.Sheets[1]);
		//	((Worksheet)wb.Sheets[2]).Delete();
		//	wb.Application.DisplayAlerts = true;
		//	string fileExt = Path.GetExtension(saveName);
		//	var xSaveName = wb.Application.GetSaveAsFilename(Path.GetFileNameWithoutExtension(saveName), "Excel Workbook (*" + fileExt + "), *" + fileExt);
		//	wb.Application.DisplayAlerts = false;
		//	XlFileFormat fileFormat = fileExt switch
		//	{
		//		".xlsx" => XlFileFormat.xlOpenXMLWorkbook,
		//		".xlsm" => XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
		//		".xls" => XlFileFormat.xlExcel8,
		//		".csv" => XlFileFormat.xlCSV,
		//		".txt" => XlFileFormat.xlTextWindows,
		//		_ => XlFileFormat.xlWorkbookDefault,
		//	};
		//	wb.Application.DisplayAlerts = true;
		//	wb.SaveAs(Filename: xSaveName, FileFormat: fileFormat, CreateBackup: false, Password: password, ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges);

		//	wb.Close();
		//	Marshal.ReleaseComObject(wb);
		//	return xSaveName;
		//}

		#region Private Methods
		private static bool SheetExists(Workbook workbook, string sheetName)
		{
			foreach (Worksheet sht in workbook.Worksheets)
			{
				if (sht.Name == sheetName)
					return true;
			}
			return false;
		}
		#endregion

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls
		protected virtual void Dispose(bool disposing)
		{
			if (!disposedValue)
			{
				if (disposing)
				{
					Marshal.ReleaseComObject(Worksheet);
					Worksheet = null;
					GC.Collect();
					GC.WaitForPendingFinalizers();
				}
				disposedValue = true;
			}
		}
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion
	}
}
