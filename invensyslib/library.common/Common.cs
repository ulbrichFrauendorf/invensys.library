using System;
using System.IO;
using System.Linq;

namespace library.common
{
	public static class Common
	{
		public static bool CreateDirectoryIfNotExist(string directory)
		{
			try
			{
				if (!Directory.Exists(directory))
					Directory.CreateDirectory(directory);
			}
			catch
			{
				return false;
			}
			return true;
		}

		public static string MakeValidFileName(string name)
		{
			string invalidChars = System.Text.RegularExpressions.Regex.Escape(new string(System.IO.Path.GetInvalidFileNameChars()));
			string invalidRegStr = $@"([{invalidChars}]*\.+$)|([{invalidChars}]+)";
			return System.Text.RegularExpressions.Regex.Replace(name, invalidRegStr, "");
		}

		public static string GetDateTimeStamp(DateTime date) => $"{date.Year.ToString().Substring(0, 4)}{date.Month.ToString("00")}{date.Day.ToString("00")}_{date.Hour}{date.Minute}{date.Second}";

		#region Extension methods

		public static bool TryCast<T>(object obj, out T result)
		{
			if (obj is T)
			{
				result = (T)obj;
				return true;
			}

			result = default(T);
			return false;
		}

		public static bool IsIn<T>(this T obj, params T[] collection) => collection.Contains(obj);

		public static string Truncate(this string value, int maxLength)
		{
			if (string.IsNullOrEmpty(value)) return value;
			return value.Length <= maxLength ? value : value.Substring(0, maxLength);
		}

		#endregion Extension methods
	}
}