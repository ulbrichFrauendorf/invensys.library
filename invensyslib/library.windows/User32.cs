using PInvoke;
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace WindowsLib
{
	internal static class User32
	{
		[DllImport("user32.dll")]
		internal static extern IntPtr GetForegroundWindow();

		[DllImport("user32")]
		internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

		/// <summary>
		/// Check if windows visible
		/// </summary>
		/// <param name="hWnd"></param>
		/// <returns>Boolean</returns>
		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool IsWindowVisible(IntPtr hWnd);

		/// <summary>
		/// Find the window Title
		/// </summary>
		/// <param name="hWnd"></param>
		/// <param name="lpWindowText"></param>
		/// <param name="nMaxCount"></param>
		/// <returns>Title Text</returns>
		[DllImport("user32.dll", EntryPoint = nameof(GetWindowText), ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
		internal static extern int GetWindowText(IntPtr hWnd, StringBuilder lpWindowText, int nMaxCount);

		//Enumerate Windows
		[DllImport("user32.dll", EntryPoint = nameof(EnumDesktopWindows), ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
		internal static extern bool EnumDesktopWindows(IntPtr hDesktop, EnumDelegate lpEnumCallbackFunction, IntPtr lParam);

		internal delegate bool EnumDelegate(IntPtr hWnd, int lParam);

		[DllImport("user32.Dll")]
		internal static extern bool EnumChildWindows(IntPtr hWndParent, PChildCallBack lpEnumFunc, int lParam);

		internal delegate bool PChildCallBack(IntPtr hWnd, int lParam);

		//
		[DllImport("user32.dll")]
		internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

		[DllImport("User32.Dll")]
		internal static extern void GetClassName(IntPtr hWnd, StringBuilder s, int nMaxCount);

		[DllImport("user32.dll", SetLastError = true)]
		internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

		[DllImport("kernel32.dll")]
		internal static extern uint GetCurrentThreadId();

		[DllImport("user32.dll")]
		internal static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

		[DllImport("user32.dll")]
		internal static extern IntPtr GetFocus();

		//Send Message Overloads
		[DllImport("user32.dll")]
		internal static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, StringBuilder lParam);

		[DllImport("user32.dll")]
		internal static extern int SendMessage(IntPtr hWnd, uint Msg, out int wParam, out int lParam);

		//[DllImport("User32.dll")]
		//internal static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, StringBuilder lParam);
		//[DllImport("User32.dll")]
		//internal static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

		[DllImport("user32.dll", SetLastError = true)]
		internal static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

		[DllImport("user32.dll", SetLastError = true)]
		internal static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string className, string lpszWindow);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
		internal static extern int SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, string lParam);

		[DllImport("user32.dll")]
		internal static extern IntPtr GetDlgItem(IntPtr hWnd, int nIDDlgItem);

		internal const uint WM_GETTEXT = 0x0D;
		internal const uint WM_GETTEXTLENGTH = 0x0E;
		internal const uint EM_GETSEL = 0xB0;
		internal const int WM_SETTEXT = 12;

		[DllImport("user32.dll", SetLastError = true)]
		internal static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

		[DllImport("user32.dll")]
		internal static extern IntPtr MonitorFromWindow(IntPtr hWnd, uint dwFlags);

		internal const int MONITOR_DEFAULTONNULL = 0;
		internal const int MONITOR_DEFAULTONPRIMARY = 1;
		internal const int MONITOR_DEFAULTONNEAREST = 2;
	}
}