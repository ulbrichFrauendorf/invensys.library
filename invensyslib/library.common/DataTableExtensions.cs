using System;
using System.Data;
using System.Linq;

namespace library.common
{
	public static class DataTableExtension
	{
		public static object[,] ToArray(this DataTable dt)
		{
			DataRowCollection rows = dt.Rows;
			int rowCount = rows.Count;
			int colCount = dt.Columns.Count;
			object[,] result = new object[rowCount + 1, colCount];

			string[] columnNames = dt.Columns.Cast<DataColumn>()
																 .Select(x => x.ColumnName)
																 .ToArray();

			for (int i = 0; i < colCount; i++)
			{
				result[0, i] = columnNames[i];
			}

			for (int i = 0; i < rowCount; i++)
			{
				DataRow row = rows[i];
				for (int j = 0; j < colCount; j++)
				{
					result[i + 1, j] = row[j];
				}
			}

			return result;
		}
		public static bool IsDataColumnNumericType(this DataColumn o)
		{
			Type x = o.DataType;

			switch (x.Name)
			{
				case "Double":
				case "Decimal":
					return true;
				default:
					return false;
			}
		}
	}
}
