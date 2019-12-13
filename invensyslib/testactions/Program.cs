using library.microsofthelper;
using library.microsofthelper.MsExcel;
using System;
using System.Data;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace testactions
{
	internal class Program
	{
		private static void Main(string[] args)
		{
			GetData();
		}

		private static void GetData()
		{
			using ExcelWorkbook msx = new ExcelWorkbook(@"C:\Citrix\Sage Monthly Process Checklist_Final.xlsx");
			using ExcelSheet wss = new ExcelSheet(msx.Workbook, "Sheet1");
			var id = 1;
			DataSet feedbackTables;
			feedbackTables = new DataSet("Personnel Feedback");
			GetReportInfo(wss, id, feedbackTables);
			GetPAInformation(wss, id, feedbackTables);
			GetAMInformation(wss, id, feedbackTables);
		}
		private static void WriteData(DataSet feedbackTables)
		{
			feedbackTables.Tables["Info"].Merge(feedbackTables.Tables["FeedBackAM"]);
			feedbackTables.Tables["Info"].Merge(feedbackTables.Tables["FeedBackPA"]);
			using ExcelWorkbook sms = new ExcelWorkbook("FeedBackReport.xlsx", "");
			using ExcelSheet sw = new ExcelSheet(sms.Workbook, "FeedBackReport");
			sw.WriteDatatableToRange(1, 1, feedbackTables.Tables["Info"]);
			sms.Workbook.Save();
		}
		//Report Gather info
		private static void GetAMInformation(ExcelSheet wss, int id, DataSet feedbackTables)
		{
			Microsoft.Office.Interop.Excel.Worksheet ws = wss.Worksheet; 
			var amRows = wss.FindDataRows("PA (To rate AM) ", 2);
			DataTable AccountManagers = new DataTable("FeedBackAM");
			AccountManagers.Columns.Add("Id", typeof(int));
			AccountManagers.Columns.Add("Input Late");
			AccountManagers.Columns.Add("Input Manual");
			AccountManagers.Columns.Add("Input Comment");
			AccountManagers.Columns.Add("MonthEnd Late");
			AccountManagers.Columns.Add("MonthEnd Non-Auto Reports");
			AccountManagers.Columns.Add("MonthEnd Comments");

			DataRow row = AccountManagers.NewRow();
			row["Id"] = id;
			row["Input Late"] = ((Range)ws.Cells[amRows[0], 4]).Value;
			row["Input Manual"] = ((Range)ws.Cells[amRows[0], 6]).Value;
			row["Input Comment"] = ((Range)ws.Cells[amRows[0], 8]).Value;
			row["MonthEnd Late"] = ((Range)ws.Cells[amRows[1], 4]).Value;
			row["MonthEnd Non-Auto Reports"] = ((Range)ws.Cells[amRows[1], 6]).Value;
			row["MonthEnd Comments"] = ((Range)ws.Cells[amRows[1], 8]).Value;

			AccountManagers.Rows.Add(row);
			feedbackTables.Tables.Add(AccountManagers);
		}
		private static void GetPAInformation(ExcelSheet wss, int id, DataSet feedbackTables)
		{
			Microsoft.Office.Interop.Excel.Worksheet ws = wss.Worksheet;
			var paRows = wss.FindDataRows("AM (To rate PA) ", 2);
			DataTable administrators = new DataTable("FeedBackPA");
			administrators.Columns.Add("Id", typeof(int));
			administrators.Columns.Add("Internal Error Free");
			administrators.Columns.Add("Internal Service Rating");
			administrators.Columns.Add("Internal Comment");
			administrators.Columns.Add("MonthEnd Error Free");
			administrators.Columns.Add("MonthEnd Service Rating");
			administrators.Columns.Add("MonthEnd Comment");
			administrators.Columns.Add("RollOver Checklist");
			administrators.Columns.Add("RollOver Error Free");
			administrators.Columns.Add("RollOver Comment");

			DataRow row = administrators.NewRow();
			row["Id"] = id;
			row["Internal Error Free"] = ((Range)ws.Cells[paRows[0], 4]).Value;
			row["Internal Service Rating"] = ((Range)ws.Cells[paRows[0], 6]).Value;
			row["Internal Comment"] = ((Range)ws.Cells[paRows[0], 8]).Value;
			row["MonthEnd Error Free"] = ((Range)ws.Cells[paRows[1], 4]).Value;
			row["MonthEnd Service Rating"] = ((Range)ws.Cells[paRows[1], 6]).Value;
			row["MonthEnd Comment"] = ((Range)ws.Cells[paRows[1], 8]).Value;
			row["RollOver Checklist"] = ((Range)ws.Cells[paRows[2], 4]).Value;
			row["RollOver Error Free"] = ((Range)ws.Cells[paRows[2], 6]).Value;
			row["RollOver Comment"] = ((Range)ws.Cells[paRows[2], 8]).Value;

			administrators.Rows.Add(row);
			feedbackTables.Tables.Add(administrators);
		}
		private static void GetReportInfo(ExcelSheet wss, int id, DataSet feedbackTables)
		{
			Microsoft.Office.Interop.Excel.Worksheet ws = wss.Worksheet;

			DataTable infoTable = new DataTable("Info");
			infoTable.Columns.Add("Id", typeof(int));
			infoTable.Columns.Add("Name");
			infoTable.Columns.Add("Month");
			infoTable.Columns.Add("Company Number");
			infoTable.Columns.Add("Environment");
			infoTable.Columns.Add("Directory");
			infoTable.Columns.Add("Pay Frequency");
			infoTable.Columns.Add("Site Code");
			infoTable.Columns.Add("CUID");
			infoTable.Columns.Add("Account Manager");
			infoTable.Columns.Add("Payroll Administrator");
			infoTable.Columns.Add("Rollover Administrator");
			infoTable.PrimaryKey = new DataColumn[] { infoTable.Columns["Id"] };

			DataRow row = infoTable.NewRow();
			row["Id"] = id;
			string name = wss.SetRange(1, 2, 1, 2).Value.ToString();
			row["Name"] = name.Replace("Monthly Process ", "").Replace(" ", "");
			row["Month"] = ((Range)ws.Cells[2, 3]).Value;
			row["Company Number"] = ((Range)ws.Cells[4, 3]).Value;
			row["Environment"] = ((Range)ws.Cells[5, 3]).Value;
			row["Directory"] = ((Range)ws.Cells[6, 3]).Value;
			row["Pay Frequency"] = ((Range)ws.Cells[7, 3]).Value;
			row["Site Code"] = ((Range)ws.Cells[2, 9]).Value;
			row["CUID"] = ((Range)ws.Cells[3, 9]).Value;
			row["Account Manager"] = ((Range)ws.Cells[6, 9]).Value;
			row["Payroll Administrator"] = ((Range)ws.Cells[7, 9]).Value;
			row["Rollover Administrator"] = ((Range)ws.Cells[8, 9]).Value;

			infoTable.Rows.Add(row);
			feedbackTables.Tables.Add(infoTable);
		}
	}
}
