using System;
using System.Text.RegularExpressions;
using WindowsLib;

public class WindowFinder
{
	private event FoundWindowCallback foundWindow;

	public delegate bool FoundWindowCallback(IntPtr hWnd);

	// Members that'll hold the search criterias while searching.
	private IntPtr parentHandle;

	private Regex className;
	private Regex windowText;
	private Regex process;

	// The main search function of the WindowFinder class. The parentHandle parameter is optional, taking in a zero if omitted.
	// The className can be null as well, in this case the class name will not be searched. For the window text we can input
	// a Regex object that will be matched to the window text, unless it's null. The process parameter can be null as well,
	// otherwise it'll match on the process name (Internet Explorer = "iexplore"). Finally we take the FoundWindowCallback
	// function that'll be called each time a suitable window has been found.
	public void FindWindows(IntPtr parentHandle, Regex className, Regex windowText, Regex process, FoundWindowCallback fwc)
	{
		this.parentHandle = parentHandle;
		this.className = className;
		this.windowText = windowText;
		this.process = process;

		// Add the FounWindowCallback to the foundWindow event.
		foundWindow = fwc;

		// Invoke the EnumChildWindows function.
		User32.EnumChildWindows(parentHandle, new User32.PChildCallBack(enumChildWindowsCallback), 0);
	}

	// This function gets called each time a window is found by the EnumChildWindows function. The foun windows here
	// are NOT the final found windows as the only filtering done by EnumChildWindows is on the parent window handle.
	private bool enumChildWindowsCallback(IntPtr handle, int lParam) =>
		// If a class name was provided, check to see if it matches the window.
		//if (className != null)
		//{
		//	StringBuilder sbClass = new StringBuilder(256);
		//	User32.GetClassName(handle, sbClass, sbClass.Capacity);

		//	// If it does not match, return true so we can continue on with the next window.
		//	if (!className.IsMatch(sbClass.ToString()))
		//		return true;
		//}

		//// If a window text was provided, check to see if it matches the window.
		//if (windowText != null)
		//{
		//	int txtLength = User32.SendMessage(handle, User32.WM_GETTEXTLENGTH, 0, null);
		//	StringBuilder sbText = new StringBuilder(txtLength + 1);
		//	User32.SendMessage(handle, User32.WM_GETTEXT, sbText.Capacity, sbText);

		//	// If it does not match, return true so we can continue on with the next window.
		//	if (!windowText.IsMatch(sbText.ToString()))
		//		return true;
		//}

		//// If a process name was provided, check to see if it matches the window.
		//if (process != null)
		//{
		//	int processID;
		//	User32.GetWindowThreadProcessId(handle, out processID);

		//	// Now that we have the process ID, we can use the built in .NET function to obtain a process object.
		//	Process p = Process.GetProcessById(processID);

		//	// If it does not match, return true so we can continue on with the next window.
		//	if (!process.IsMatch(p.ProcessName))
		//		return true;
		//}

		// If we get to this point, the window is a match. Now invoke the foundWindow event and based upon
		// the return value, whether we should continue to search for windows.
		foundWindow(handle);
}