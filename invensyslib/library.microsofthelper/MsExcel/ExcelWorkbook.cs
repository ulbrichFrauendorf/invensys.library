using library.common;
using Microsoft.Office.Interop.Excel;
using System;
using System.IO;

namespace library.microsofthelper.MsExcel
{
	public class ExcelWorkbook : IDisposable
	{
		public Application ExcelApplication { get; private set; }
		public Workbook Workbook { get; private set; }
		internal Workbooks ExcelWorkbooks { get; set; }
		public Sheets Worksheets { get; set; }
		private string xSaveName;


		public ExcelWorkbook(string filename, string password = "")
		{
			ExcelApplication = new Application
			{
				DisplayAlerts = false
			};

			ExcelWorkbooks = ExcelApplication.Workbooks;
			if (File.Exists(filename))
			{
				try
				{
					Workbook = ExcelWorkbooks.Open(Filename: filename, UpdateLinks: false, ReadOnly: false, Password: password);
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
					Workbook = ExcelWorkbooks.Add(XlWBATemplate.xlWBATWorksheet);
					Save(filename, true, password);
					Workbook.Close();
					Cleanup.ReleaseObject(Workbook);
					Workbook = ExcelWorkbooks.Open(Filename: xSaveName, UpdateLinks: false, ReadOnly: false, Password: password);
				}
				catch (Exception ex)
				{
					throw new LocalSystemException("Could not create new Excel file : " + filename, ex);
				}
			}

			Worksheets = Workbook.Worksheets;
		}
		public string Save(string filename, bool savePopupFlag = true, string password = "")
		{
			string fileExt = Path.GetExtension(filename);
			if (savePopupFlag)
			{
				ExcelApplication.DisplayAlerts = true;
				xSaveName = ExcelApplication.GetSaveAsFilename(Path.GetFileNameWithoutExtension(filename), "Excel Workbook (*" + fileExt + "), *" + fileExt);
				ExcelApplication.DisplayAlerts = false;
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
			ExcelApplication.DisplayAlerts = false;
			Workbook.SaveAs(Filename: xSaveName, FileFormat: fileFormat, CreateBackup: false, Password: password, ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges);
			return xSaveName;
		}
		public void Save()
		{
			Workbook.Save();
		}

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls
		protected virtual void Dispose(bool disposing)
		{
			if (!disposedValue)
			{
				if (disposing)
				{
					Cleanup.ReleaseObject(Worksheets);
					Workbook.Close(0);
					Cleanup.ReleaseObject(Workbook);
					Cleanup.ReleaseObject(ExcelWorkbooks);
					ExcelApplication.Quit();
					Cleanup.ReleaseObject(ExcelApplication);
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