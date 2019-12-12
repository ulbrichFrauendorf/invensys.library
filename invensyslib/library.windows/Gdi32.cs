using System;
using System.Runtime.InteropServices;

namespace WindowsLib
{
	internal static class Gdi32
	{
		[DllImport("gdi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		internal static extern IntPtr CreateDC(string lpszDriver, string lpszDeviceName, string lpszOutput, IntPtr devMode);

		[DllImport("gdi32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
		internal static extern int GetDeviceCaps(IntPtr hDC, int nIndex);

		[DllImport("user32.dll")]
		internal static extern IntPtr GetDC(IntPtr hWnd);

		[DllImport("user32.dll")]
		internal static extern int ReleaseDC(IntPtr hWnd, IntPtr hDc);

		internal const int HORZRES = 8;
		internal const int VERTRES = 10;
		internal const int DESKTOPVERTRES = 117;
		internal const int LOGPIXELSX = 88;
		internal const int LOGPIXELSY = 90;
		internal const int DESKTOPHORZRES = 118;
	}
}