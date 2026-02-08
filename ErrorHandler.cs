using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace NM.Core
{
    /// <summary>Centralized error handling and logging for NM SolidWorks Automator.</summary>
    public static class ErrorHandler
    {
        private const string ErrorSeparator = "------------------------------------------------";
        private const string DateFormat = "yyyy-MM-dd HH:mm:ss";

        private static readonly List<string> _callStack = new List<string>();
        private static readonly object _lockObject = new object();

        /// <summary>
        /// Additional debug log file path (used by QA runner for session-specific logging).
        /// Set to a file path to capture all DebugLog calls to that file as well.
        /// </summary>
        public static string AdditionalDebugLogPath { get; set; }

        /// <summary>Structured log level values for compile-time safety and easy mapping to sinks.</summary>
        public enum LogLevel
        {
            Info = 20,
            Warning = 30,
            Error = 40,
            Critical = 50
        }

        /// <summary>Push the current procedure name onto the call stack.</summary>
        public static void PushCallStack(string procedureName)
        {
            if (string.IsNullOrWhiteSpace(procedureName)) return;
            lock (_lockObject)
            {
                _callStack.Add(procedureName);
            }
        }

        /// <summary>Pop the most recent procedure name from the call stack.</summary>
        public static void PopCallStack()
        {
            lock (_lockObject)
            {
                if (_callStack.Count > 0)
                {
                    _callStack.RemoveAt(_callStack.Count - 1);
                }
            }
        }

        /// <summary>Write a diagnostic message when debug mode is enabled; optionally log to file.</summary>
        public static void DebugLog(string message, bool includeTimestamp = true)
        {
            if (!Configuration.Logging.EnableDebugMode) return;

            var line = includeTimestamp
                ? $"{DateTime.Now.ToString(DateFormat)}: {message}"
                : message;

            Debug.WriteLine(line);

            if (Configuration.Logging.LogEnabled)
            {
                TryWriteToLog(line);
            }

            // Also write to additional debug log if set (used by QA runner)
            if (!string.IsNullOrEmpty(AdditionalDebugLogPath))
            {
                TryWriteToAdditionalLog(line);
            }
        }

        /// <summary>
        /// Always-on informational log for operational events (performance mode changes,
        /// scope entry/exit, state transitions). Unlike DebugLog, this writes in ALL modes
        /// including Normal and Production when file logging is enabled. Cost: one boolean
        /// check + timestamp format + file append (~0.1ms). Use for events that an AI
        /// debugging agent needs to see even when verbose debug tracing is off.
        /// </summary>
        public static void LogInfo(string message)
        {
            if (!Configuration.Logging.LogEnabled) return;

            var line = $"{DateTime.Now.ToString(DateFormat)}: {message}";

            // Always write to file log when logging is enabled (independent of debug mode)
            TryWriteToLog(line);

            // Also write to additional debug log if set (QA runner captures these)
            if (!string.IsNullOrEmpty(AdditionalDebugLogPath))
            {
                TryWriteToAdditionalLog(line);
            }

            // Debug output only when debug mode is on (avoids noise in VS output window during production)
            if (Configuration.Logging.EnableDebugMode)
            {
                Debug.WriteLine(line);
            }
        }

        private static void TryWriteToAdditionalLog(string message)
        {
            try
            {
                lock (_lockObject)
                {
                    File.AppendAllText(AdditionalDebugLogPath, message + Environment.NewLine);
                }
            }
            catch { /* ignore logging failures */ }
        }

        /// <summary>New preferred overload using LogLevel. Centralized error handler to build and write a formatted log entry.</summary>
        public static void HandleError(
            string procedureName = "Unknown",
            string errorDesc = "",
            Exception ex = null,
            LogLevel level = LogLevel.Error,
            string context = "")
        {
            try
            {
                var time = DateTime.Now.ToString(DateFormat);
                var callStack = GetCallStackString();

                // Derive standard fields from exception (if any)
                var errNum = ex?.HResult ?? 0;
                var src = ex?.Source ?? string.Empty;
                var desc = !string.IsNullOrWhiteSpace(errorDesc) ? errorDesc : (ex?.Message ?? string.Empty);
                var stack = ex?.StackTrace ?? string.Empty;

                var msg = new StringBuilder();
                msg.AppendLine(ErrorSeparator);
                msg.AppendLine($"Time: {time}");
                msg.AppendLine($"Level: {level} ({(int)level})");
                msg.AppendLine($"Call Stack: {callStack}");
                msg.AppendLine($"Proc: {procedureName}");
                msg.AppendLine($"Err#: {errNum}");
                if (!string.IsNullOrEmpty(src)) msg.AppendLine($"Src: {src}");
                if (!string.IsNullOrEmpty(desc)) msg.AppendLine($"Desc: {desc}");
                if (!string.IsNullOrEmpty(context)) msg.AppendLine($"Context: {context}");
                if (!string.IsNullOrEmpty(stack)) msg.AppendLine($"StackTrace: {stack}");
                // NOTE: System info (e.g., SW version) intentionally omitted to keep NM.Core free of COM types
                // TODO(vNext): Allow NM.SwAddin to inject environment/system info into messages when needed.
                msg.AppendLine($"Log Path: {Configuration.Logging.ErrorLogPath}");
                msg.AppendLine(ErrorSeparator);

                var finalMessage = msg.ToString();

                // Console/Debug output when debugging is enabled
                if (Configuration.Logging.EnableDebugMode)
                {
                    Debug.WriteLine(finalMessage);
                    Console.WriteLine(finalMessage);
                }

                // File logging: errors/warnings ALWAYS write when logging is enabled,
                // regardless of debug mode. An AI debugging agent needs the error trail.
                if (Configuration.Logging.LogEnabled)
                {
                    TryWriteToLog(finalMessage);
                }

                // QA runner additional log: always capture errors
                if (!string.IsNullOrEmpty(AdditionalDebugLogPath))
                {
                    TryWriteToAdditionalLog(finalMessage);
                }

                // TODO(vNext): UI notifications (MessageBox/TaskPane) controlled by Logging.ShowWarnings & level
            }
            catch (Exception logEx)
            {
                // Last-resort fallback: write to Debug to avoid recursive failures
                Debug.WriteLine($"Error while handling error: {logEx}");
            }
        }

        /// <summary>Returns the current call stack depth for performance tracking.</summary>
        public static int CallStackDepth
        {
            get
            {
                lock (_lockObject)
                {
                    return _callStack.Count;
                }
            }
        }

        private static string GetCallStackString()
        {
            lock (_lockObject)
            {
                if (_callStack.Count == 0) return "Empty";
                return string.Join(" -> ", _callStack);
            }
        }

        private static void TryWriteToLog(string message)
        {
            try
            {
                var path = Configuration.Logging.ErrorLogPath;
                if (string.IsNullOrWhiteSpace(path)) return;

                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                File.AppendAllText(path, message + Environment.NewLine);
            }
            catch
            {
                // Fallback to temp path if the configured path fails
                try
                {
                    var temp = Path.Combine(Path.GetTempPath(), "SolidWorksMacroErrorLog.txt");
                    File.AppendAllText(temp, message + Environment.NewLine);
                }
                catch
                {
                    // swallow - nowhere else to write
                }
            }
        }
    }
}