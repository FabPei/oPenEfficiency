using System;
using System.Diagnostics;
using System.IO;

namespace oPenEfficiency.Utils
{
    /// <summary>
    /// Central exception logging utility for the add-in.
    /// Logs exceptions to a file in the user's temp folder for debugging.
    /// </summary>
    public static class ExceptionLogger
    {
        private static readonly string LogFilePath = Path.Combine(
            Path.GetTempPath(),
            "oPenEfficiency",
            $"error_{DateTime.Now:yyyyMMdd}.log");

        private static readonly object LockObj = new object();

        /// <summary>
        /// Logs an exception to the error log file and debug output.
        /// </summary>
        public static void Log(Exception ex, string context = null)
        {
            var message = FormatException(ex, context);

            Debug.WriteLine(message);

            try
            {
                lock (LockObj)
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath));
                    File.AppendAllText(LogFilePath, message + Environment.NewLine);
                }
            }
            catch (Exception logEx)
            {
                // If logging fails, at least we have Debug output
                Debug.WriteLine($"Failed to write to log file: {logEx.Message}");
            }
        }

        /// <summary>
        /// Logs a simple error message without exception details.
        /// </summary>
        public static void Log(string message, string context = null)
        {
            var fullMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}]";
            if (!string.IsNullOrEmpty(context))
                fullMessage += $" [{context}]";
            fullMessage += $" ERROR: {message}";

            Debug.WriteLine(fullMessage);

            try
            {
                lock (LockObj)
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath));
                    File.AppendAllText(LogFilePath, fullMessage + Environment.NewLine);
                }
            }
            catch (Exception logEx)
            {
                Debug.WriteLine($"Failed to write to log file: {logEx.Message}");
            }
        }

        private static string FormatException(Exception ex, string context)
        {
            var message = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}]";
            if (!string.IsNullOrEmpty(context))
                message += $" [{context}]";
            message += $" EXCEPTION: {ex.GetType().Name}: {ex.Message}";
            message += Environment.NewLine + "Stack Trace: " + ex.StackTrace;

            if (ex.InnerException != null)
            {
                message += Environment.NewLine + "Inner Exception: " + ex.InnerException.Message;
            }

            return message;
        }

        /// <summary>
        /// Gets the path to the current log file for debugging.
        /// </summary>
        public static string GetLogFilePath() => LogFilePath;
    }
}
