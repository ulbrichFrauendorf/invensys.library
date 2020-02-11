using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace library.microsofthelper.MsExcel
{
	public static class ExcelReportFormat
	{
		private const int COLOR_CELL_HEADING = 13696965;
		public static void FormatReportStandard(this ExcelSheet sheet, int headingRow = 1, int freezeColumn = 0)
		{
			//sheet.FormatSetJustificationLeft();
			sheet.FormatSetFont();
			//sheet.FormatBorders();
			//sheet.FormatHeadingStandard(headingRow);
			//sheet.FormatAutoFitAllColumns();
			//sheet.FormatApplyFilters();
			//sheet.FormatFreezePanes(headingRow + 1, freezeColumn + 1);
		}
		public static void FormatBorders(this ExcelSheet sheet, Range range = null)
		{
			if (range == null)
				range = sheet.GetUsedRange();

			Borders borders = range.Borders;
			borders[XlBordersIndex.xlDiagonalDown].LineStyle = XlLineStyle.xlLineStyleNone;
			borders[XlBordersIndex.xlDiagonalUp].LineStyle = XlLineStyle.xlLineStyleNone;

			borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
			borders[XlBordersIndex.xlEdgeBottom].ThemeColor = 1;
			borders[XlBordersIndex.xlEdgeBottom].TintAndShade = -0.499984740745262;
			borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;

			borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
			borders[XlBordersIndex.xlEdgeTop].ThemeColor = 1;
			borders[XlBordersIndex.xlEdgeTop].TintAndShade = -0.499984740745262;
			borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;

			borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
			borders[XlBordersIndex.xlEdgeLeft].ThemeColor = 1;
			borders[XlBordersIndex.xlEdgeLeft].TintAndShade = -0.499984740745262;
			borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;

			borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
			borders[XlBordersIndex.xlEdgeRight].ThemeColor = 1;
			borders[XlBordersIndex.xlEdgeRight].TintAndShade = -0.499984740745262;
			borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;

			borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
			borders[XlBordersIndex.xlInsideHorizontal].ThemeColor = 1;
			borders[XlBordersIndex.xlInsideHorizontal].TintAndShade = -0.499984740745262;
			borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;

			borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
			borders[XlBordersIndex.xlInsideVertical].ThemeColor = 1;
			borders[XlBordersIndex.xlInsideVertical].TintAndShade = -0.499984740745262;
			borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlThin;
			Cleanup.ReleaseObject(range);
		}
		public static void FormatFilTableColours(this ExcelSheet sheet, bool hasColour)
		{
			Range range = sheet.GetUsedRange();
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
		public static void FormatAutoFitAllColumns(this ExcelSheet sheet)
		{
			Range rangeA = (Range)sheet.Worksheet.Cells;
			rangeA = rangeA.EntireColumn;
			rangeA.AutoFit();

			for (int i = 1; i <= sheet.FindLastUsedColumn(); i++)
			{
				Range range = (Range)sheet.Worksheet.Columns[i];
				range.ColumnWidth += 3;
				Cleanup.ReleaseObject(range);
			}
			Cleanup.ReleaseObject(rangeA);
		}
		public static void FormatFreezePanes(this ExcelSheet sheet, int freezeRow = 1, int freezeColumn = 1)
		{
			sheet.Worksheet.Activate();
			Range range = sheet.Worksheet.Cells[freezeRow, freezeColumn];
			range.Select();
			sheet.ExcelApplication.ActiveWindow.FreezePanes = true;
			Cleanup.ReleaseObject(range);
		}
		public static void FormatApplyFilters(this ExcelSheet sheet, int filterRow = 1, int filterColumn = 1)
		{
			Range range = sheet.SetRange(filterRow, filterColumn, filterRow, filterColumn);
			range.AutoFilter(2);
			Cleanup.ReleaseObject(range);
		}
		public static void FormatSetFont(this ExcelSheet sheet)
		{
			Range range = sheet.GetUsedRange();
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
		public static void FormatSetJustificationLeft(this ExcelSheet sheet)
		{
			Range range = sheet.GetUsedRange();
			range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			range.VerticalAlignment = XlVAlign.xlVAlignCenter;
			Cleanup.ReleaseObject(range);
		}
		public static void FormatHeadingStandard(this ExcelSheet sheet, int headingRow = 1)
		{
			Range range = sheet.SetRange(headingRow, 1, headingRow, sheet.FindLastUsedColumn(headingRow));
			range.Interior.Pattern = XlPattern.xlPatternSolid;
			range.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
			range.Interior.Color = COLOR_CELL_HEADING;
			range.Interior.TintAndShade = 0;
			range.Interior.PatternTintAndShade = 0;
			range.Font.Bold = true;
			range.WrapText = true;
			Cleanup.ReleaseObject(range);
		}
		public static void FormatMergeAndCentre(this ExcelSheet sheet, int row1, int column1, int row2, int column2)
		{
			Range rng = sheet.SetRange(row1, column1, row2, column2);
			rng.Merge();
			rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			rng.VerticalAlignment = XlVAlign.xlVAlignCenter;
			Cleanup.ReleaseObject(rng);
		}
	}
}
