using library.microsofthelper;
using System;

namespace testactions
{
	internal class Program
	{
		private static void Main(string[] args)
		{
			using MsExcel msx = new MsExcel(@"C:\Citrix\Book.xlsx");
			System.Data.DataTable tble = msx.ConvertExcelSheetToDataTable("Sheet1");
			
			Console.WriteLine("Hello World!");
		}
	}
}
