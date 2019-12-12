using System;
using System.Diagnostics;
using System.IO;

namespace library.common
{
	/// <summary>
	/// Adds a log file to the windows event logger
	/// </summary>
	public class WinLogger
	{
		public const int STOPPED = 1;
		public const int STARTED = 2;
		public const int RUNNING = 3;
		public const int ERRORED = 69;

		private readonly EventLog outsEventLog;

		public WinLogger(string Source, string LogName)
		{
			outsEventLog = new EventLog();
			if (!EventLog.SourceExists(Source))
			{
				EventLog.CreateEventSource(
						Source, LogName);
			}
			outsEventLog.Source = Source;
			outsEventLog.Log = LogName;
		}

		public void FireWindowsLog(ApplicationEventType eventType, string customMessage = "")

		{
			switch (eventType)
			{
				case ApplicationEventType.Started:
					outsEventLog.WriteEntry("Application started - " + customMessage, EventLogEntryType.Information, STARTED);
					break;

				case ApplicationEventType.Running:
					outsEventLog.WriteEntry("Application running - " + customMessage, EventLogEntryType.Information, RUNNING);
					break;

				case ApplicationEventType.Stopped:
					outsEventLog.WriteEntry("Application stopped - " + customMessage, EventLogEntryType.Information, STOPPED);
					break;

				case ApplicationEventType.Error:
					outsEventLog.WriteEntry("Application ERROR   - " + customMessage, EventLogEntryType.Error, ERRORED);
					break;

				default:
					throw new Exception("Unexpected Case");
			}
		}

		public enum ApplicationEventType { Started, Running, Stopped, Error }
	}

	/// <summary>
	/// Creates a text log file
	/// </summary>
	public class TextLogger
	{
		private readonly string outsArchiveLogFile;

		public TextLogger(string logFile) => outsArchiveLogFile = logFile;

		public void FireTextLog(InternalProcessEventType eventType, string customMessage = "")
		{
			switch (eventType)
			{
				case InternalProcessEventType.Success:
					WriteTextEntry("Success -> " + customMessage + Environment.NewLine);
					break;

				case InternalProcessEventType.Failure:
					WriteTextEntry("Failure -> " + customMessage + Environment.NewLine);
					break;

				default:
					throw new Exception("Unexpected Case");
			}
		}

		private void WriteTextEntry(string v) => File.AppendAllText(outsArchiveLogFile, v);

		public enum InternalProcessEventType { Success, Failure, }
	}
}