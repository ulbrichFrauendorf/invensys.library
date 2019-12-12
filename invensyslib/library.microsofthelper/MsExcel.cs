using library.common;
using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.Runtime.InteropServices;
using DataTable = System.Data.DataTable;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace library.microsofthelper
{
	public class MsExcel : IDisposable
	{
		public Application ExcelApplication { get; private set; }
		public Workbook Workbook { get; private set; }

		public MsExcel(string filename, string password = "")
		{
			ExcelApplication = new Application
			{
				DisplayAlerts = false
			};
			try
			{
				Workbook = ExcelApplication.Application.Workbooks.Open(Filename: filename, ReadOnly: false, Password: password);
			}
			catch (Exception ex)
			{
				throw new LocalSystemException("Could not open locked Excel file : " + filename, ex);
			}
		}
		public DataTable ConvertExcelSheetToDataTable(string sheetName)
		{
			using ExcelSheet excelSheet = new ExcelSheet(Workbook, sheetName);
			Range range = excelSheet.GetUsedRange();
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

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls
		protected virtual void Dispose(bool disposing)
		{
			if (!disposedValue)
			{
				if (disposing)
				{
					Workbook.Close();
					ExcelApplication.Quit();
					Marshal.ReleaseComObject(Workbook);
					Marshal.ReleaseComObject(ExcelApplication);
					Workbook = null;
					ExcelApplication = null;
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
	public class ExcelSheet: IDisposable
	{
		public Worksheet Worksheet { get; set; }
		public ExcelSheet(Workbook workbook, string sheetName) => Worksheet = (Worksheet)workbook.Worksheets[sheetName];
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