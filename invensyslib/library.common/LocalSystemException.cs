using System;

namespace library.common
{
	/// <summary>
	/// Custom Exception handler, writes to windows log
	/// </summary>
	public class LocalSystemException : Exception
	{
		public LocalSystemException(string message)
				: base(message)
		{
			WinLogger log = new WinLogger("GeneralErrors", "LocalExceptions");
			log.FireWindowsLog(WinLogger.ApplicationEventType.Error, message);
		}

		public LocalSystemException(string message, Exception inner)
				: base(message, inner)
		{
			WinLogger log = new WinLogger("GeneralErrors", "LocalExceptions");
			log.FireWindowsLog(WinLogger.ApplicationEventType.Error, message + " -> " + inner);
		}

		public LocalSystemException(WinLogger log, string message)
		: base(message) => log.FireWindowsLog(WinLogger.ApplicationEventType.Error, message);

		public LocalSystemException(WinLogger log, string message, Exception inner)
				: base(message, inner) => log.FireWindowsLog(WinLogger.ApplicationEventType.Error, message + " -> " + inner);
	}
}