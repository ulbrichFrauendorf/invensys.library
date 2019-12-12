using library.common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace WindowsLib
{
	public class WindowInfo
	{
		private static int GetProcess(IntPtr handle)
		{
			try
			{
				uint p = User32.GetWindowThreadProcessId(handle, out int pId);
				return pId;
			}
			catch (Exception ex)
			{
				throw new LocalSystemException("Could not hook process.", ex);
			}
		}

		public static string GetWindowFileName(IntPtr handle)
		{
			int ProcNo = GetProcess(handle);
			if (ProcNo != 0)
				return Process.GetProcessById(ProcNo).MainModule.FileName;

			return null;
		}

		public static string GetWindowTitle(IntPtr handle)
		{
			StringBuilder buffer = new StringBuilder(255);
			int nLengt = User32.GetWindowText(handle, buffer, buffer.Capacity + 1);
			if (nLengt > 0)
				return buffer.ToString();

			return null;
		}

		private static int GetActiveProcess() => GetProcess(User32.GetForegroundWindow());

		public static string GetActiveWindowFileName()
		{
			int ProcNo = GetActiveProcess();
			if (ProcNo != 0)
				return Process.GetProcessById(ProcNo).MainModule.FileName;

			return null;
		}

		public static string GetActiveWindowTitle()
		{
			StringBuilder buffer = new StringBuilder(255);
			IntPtr handle = User32.GetForegroundWindow();
			int nLengt = User32.GetWindowText(handle, buffer, buffer.Capacity + 1);
			if (nLengt > 0)
				return buffer.ToString();

			return null;
		}

		public static IntPtr GetActiveWindowHandle()
		{
			IntPtr handle = User32.GetForegroundWindow();
			return handle;
		}

		private static List<string> GetAllOpenWindows()
		{
			List<string> collection = new List<string>();
			User32.EnumDelegate filter = delegate (IntPtr hWnd, int lParam)
			{
				StringBuilder strbTitle = new StringBuilder(255);
				int nLength = User32.GetWindowText(hWnd, strbTitle, strbTitle.Capacity + 1);
				string strTitle = strbTitle.ToString();

				if (User32.IsWindowVisible(hWnd) && string.IsNullOrEmpty(strTitle) == false)
				{
					collection.Add(strTitle);
				}
				return true; //TODO
			};

			if (User32.EnumDesktopWindows(IntPtr.Zero, filter, IntPtr.Zero))
				return collection;

			return null;
		}
	}
}