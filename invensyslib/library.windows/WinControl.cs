using PInvoke;
using System;
using System.Drawing;
using System.Text;

namespace WindowsLib
{
	public class WinControl
	{
		public IntPtr ControlPtr { get; private set; }

		public WinControl() => ControlPtr = IntPtr.Zero;

		public WinControl(IntPtr handle) => ControlPtr = handle;

		public WinControl(string className, string windowTitle) => ControlPtr = User32.FindWindow(className, windowTitle);

		public WinControl(IntPtr parentControl, string className) => ControlPtr = User32.FindWindowEx(parentControl, IntPtr.Zero, className, null);

		public WinControl(IntPtr parentControl, int controlId) => ControlPtr = User32.GetDlgItem(parentControl, controlId);

		public string Class
		{
			get
			{
				StringBuilder sb = new StringBuilder(256);
				User32.GetClassName(ControlPtr, sb, sb.Capacity);
				return sb.ToString();
			}
		}

		public string Text
		{
			get
			{
				int len = User32.SendMessage(ControlPtr, User32.WM_GETTEXTLENGTH, 0, null);
				StringBuilder sb = new StringBuilder(len);
				int numChars = User32.SendMessage(ControlPtr, User32.WM_GETTEXT, len + 1, sb);
				return sb.ToString();
			}
		}

		public IntPtr ActiveScreenHandle => User32.MonitorFromWindow(ControlPtr, User32.MONITOR_DEFAULTONNEAREST);
		public Rectangle WindowRectangle => GetRectangle(ControlPtr);

		private Rectangle GetRectangle(IntPtr hWind)
		{
			User32.GetWindowRect(hWind, out RECT outRect);
			Point topLeft = new Point(outRect.left, outRect.top);
			Size recSize = new Size(outRect.right - outRect.left, outRect.bottom - outRect.top);
			return new Rectangle(topLeft, recSize);
		}

		public WinControl GetChildIstance(int instance)
		{
			int i = 0;
			WinControl ctrl = new WinControl();
			User32.PChildCallBack filter = delegate (IntPtr hWnd, int lParam)
			{
				i++;
				if (i == instance)
				{
					ctrl.ControlPtr = hWnd;
				}

				return true;
			};

			_ = User32.EnumChildWindows(ControlPtr, filter, 0);
			return ctrl;
		}
	}
}