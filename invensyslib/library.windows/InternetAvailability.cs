using System.Runtime.InteropServices;

public class InternetAvailability
{
	[DllImport("wininet.dll")]
	private static extern bool InternetGetConnectedState(out int description, int reservedValue);

	public static bool IsInternetAvailable() => InternetGetConnectedState(out int description, 0);
}