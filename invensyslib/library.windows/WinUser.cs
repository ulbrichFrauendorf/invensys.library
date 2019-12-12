using System;
using System.Security.Principal;

namespace WindowsLib
{
	public class WinUser
	{
		public static bool IsRunAsAdmin()
		{
			try
			{
				WindowsIdentity id = WindowsIdentity.GetCurrent();
				WindowsPrincipal principal = new WindowsPrincipal(id);
				return principal.IsInRole(WindowsBuiltInRole.Administrator);
			}
			catch (Exception)
			{
				return false;
			}
		}
	}
}