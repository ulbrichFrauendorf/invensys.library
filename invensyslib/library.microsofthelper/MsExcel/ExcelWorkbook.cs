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
	public class ExcelWorkbook : IDisposable
	{
		public Application ExcelApplication { get; private set; }
		public Workbook Workbook { get; private set; }

		public ExcelWorkbook(string filename, string password = "")
		{
			ExcelApplication = new Application
			{
				DisplayAlerts = false
			};

			if (File.Exists(filename))
			{
				try
				{
					Workbook = ExcelApplication.Application.Workbooks.Open(Filename: filename, ReadOnly: false, Password: password);
				}
				catch (Exception ex)
				{
					throw new LocalSystemException("Could not open locked Excel file : " + filename, ex);
				}
			}
			else
			{
				try
				{
					Workbook = ExcelApplication.Application.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
					Save(filename, true, password);
				}
				catch (Exception ex)
				{
					throw new LocalSystemException("Could not create new Excel file : " + filename, ex);
				}
			}

		}
		public string Save(string filename, bool savePopupFlag = true, string password = "")
		{
			string fileExt = Path.GetExtension(filename);
			string xSaveName;
			if (savePopupFlag)
			{
				Workbook.Application.DisplayAlerts = true;
				xSaveName = Workbook.Application.GetSaveAsFilename(Path.GetFileNameWithoutExtension(filename), "Excel Workbook (*" + fileExt + "), *" + fileExt);
				Workbook.Application.DisplayAlerts = false;
			}
			else
			{
				if (!File.Exists(filename))
					xSaveName = filename;
				else
					xSaveName = Path.Combine(Path.GetDirectoryName(filename), Path.GetFileNameWithoutExtension(filename) + Path.GetExtension(filename));
			}
			XlFileFormat fileFormat = fileExt switch
			{
				".xlsx" => XlFileFormat.xlOpenXMLWorkbook,
				".xlsm" => XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
				".xls" => XlFileFormat.xlExcel8,
				".csv" => XlFileFormat.xlCSV,
				".txt" => XlFileFormat.xlTextWindows,
				_ => XlFileFormat.xlWorkbookDefault,
			};
			Workbook.Application.DisplayAlerts = false;
			Workbook.SaveAs(Filename: xSaveName, FileFormat: fileFormat, CreateBackup: false, Password: password, ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges);
			return xSaveName;
		}

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls
		protected virtual void Dispose(bool disposing)
		{
			if (!disposedValue)
			{
				if (disposing)
				{
					Workbook.Close(0);
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
}