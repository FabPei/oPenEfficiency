using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;

namespace oPenEfficiency.Utils
{
    /// <summary>
    /// Writes exactly one line per add-in start that names the build the Office host actually
    /// loaded: the loaded assembly's file path, its file write time and its versions.
    ///
    /// Why it exists: Office reads an add-in manifest only when the host process starts, so a
    /// running Excel or PowerPoint session can keep serving an older assembly than the one that
    /// was just built. This line answers "which build is loaded?" for a human and for a log
    /// reader without a debugger.
    ///
    /// Host-neutral: lives in oPenEfficiency.Core next to <see cref="ExceptionLogger"/> and
    /// reuses its log directory (%TEMP%\oPenEfficiency).
    /// </summary>
    public static class StartupBuildLog
    {
        /// <summary>Field placeholder used when a value cannot be read from disk.</summary>
        public const string Unavailable = "<unavailable>";

        private static readonly string LogDirectory = Path.Combine(
            Path.GetTempPath(),
            "oPenEfficiency");

        private static readonly object LockObj = new object();

        /// <summary>
        /// Startup log of the current day: %TEMP%\oPenEfficiency\startup_YYYYMMDD.log.
        /// </summary>
        public static string GetLogFilePath() =>
            Path.Combine(LogDirectory, $"startup_{DateTime.Now:yyyyMMdd}.log");

        /// <summary>
        /// Appends the startup line for <paramref name="addInAssembly"/> and returns it.
        /// Never throws: a diagnostic marker must not be able to break add-in startup.
        /// </summary>
        /// <param name="addInAssembly">The add-in's own assembly, so the line names what Office loaded.</param>
        /// <param name="addInName">Stable add-in identity (the ProgId), for example "oPenEfficiency.Excel".</param>
        public static string Record(Assembly addInAssembly, string addInName) =>
            Record(addInAssembly, addInName, GetLogFilePath());

        /// <summary>
        /// Test seam: <see cref="Record(Assembly, string)"/> against an explicit log file.
        /// </summary>
        public static string Record(Assembly addInAssembly, string addInName, string logFilePath)
        {
            if (addInAssembly == null || string.IsNullOrEmpty(logFilePath))
            {
                return null;
            }

            try
            {
                string line = FormatLine(addInAssembly, addInName, DateTime.Now);

                lock (LockObj)
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(logFilePath));
                    File.AppendAllText(logFilePath, line + Environment.NewLine);
                }
                return line;
            }
            catch (Exception ex)
            {
                // Diagnostic only: never let the startup marker fail the add-in start.
                Debug.WriteLine($"Startup build line not written: {SecretRedactor.Redact(ex.Message)}");
                return null;
            }
        }

        /// <summary>
        /// Formats the single startup line that is appended per start.
        /// The field order is pinned by oPenEfficiency.Core.Tests/StartupBuildLogTests.cs;
        /// keep the two in sync when the format changes.
        /// </summary>
        public static string FormatLine(Assembly addInAssembly, string addInName, DateTime nowLocal)
        {
            string location = GetAssemblyLocation(addInAssembly);

            return $"[{nowLocal:yyyy-MM-dd HH:mm:ss}] BUILD"
                + $" addin={addInName}"
                + $" assembly={Path.GetFileName(location)}"
                + $" assemblyVersion={GetAssemblyVersion(addInAssembly)}"
                + $" fileVersion={GetFileVersion(addInAssembly, location)}"
                + $" builtUtc={GetFileWriteTimeUtc(location)}"
                + $" path={location}";
        }

        /// <summary>
        /// Compact "built ..." hint for the add-in ribbons, so the loaded build is visible
        /// without opening the log. Returns an empty string when the build cannot be read.
        /// </summary>
        public static string FormatBuildHint(Assembly addInAssembly)
        {
            try
            {
                string location = addInAssembly?.Location;
                if (string.IsNullOrEmpty(location) || !File.Exists(location))
                {
                    return string.Empty;
                }

                DateTime stamped = ReadStampedBuildTime(addInAssembly, location);

                // Prefer the stamped version over the file's write time: Office loads the add-in from
                // a shadow copy, so that copy's write time is the copy time and looks older than the
                // build (e.g. "built 18:20" for a build of 18:40).
                return stamped == DateTime.MinValue
                    ? $"built {File.GetLastWriteTime(location):yyyy-MM-dd HH:mm}"
                    : $"built {stamped:yyyy-MM-dd HH:mm}";
            }
            catch (Exception ex)
            {
                // The ribbon hint is diagnostic only: a missing hint must not fail the ribbon.
                Debug.WriteLine($"Build hint not available: {SecretRedactor.Redact(ex.Message)}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Build time encoded in the stamped assembly version (yyyy.M.d.HHmm, see
        /// Directory.Build.props). Returns <see cref="DateTime.MinValue"/> when the version does not
        /// carry a stamp, so callers can fall back to the file write time.
        /// </summary>
        private static DateTime ReadStampedBuildTime(Assembly addInAssembly, string location)
        {
            DateTime stamped;
            string version = GetFileVersion(addInAssembly, location);

            return DateTime.TryParseExact(
                version,
                "yyyy.M.d.HHmm",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out stamped)
                ? stamped
                : DateTime.MinValue;
        }

        private static string GetAssemblyLocation(Assembly addInAssembly)
        {
            try
            {
                string location = addInAssembly?.Location;
                return string.IsNullOrEmpty(location) ? Unavailable : location;
            }
            catch (NotSupportedException)
            {
                // Assemblies loaded from memory expose no location by design.
                return Unavailable;
            }
        }

        /// <summary>
        /// One-line build description for the About section: the versions, build time and file
        /// name of the assembly the running Office host actually loaded. Answers "which build am
        /// I looking at?" without opening a log. Format pinned by StartupBuildLogTests.
        /// </summary>
        public static string FormatAboutLine(Assembly addInAssembly)
        {
            string location = GetAssemblyLocation(addInAssembly);

            return $"Version {GetAssemblyVersion(addInAssembly)} | file {GetFileVersion(addInAssembly, location)}"
                + $" | {FormatBuildHint(addInAssembly)} | {Path.GetFileName(location)}";
        }

        private static string GetFileWriteTimeUtc(string location)
        {
            try
            {
                return File.Exists(location)
                    ? File.GetLastWriteTimeUtc(location)
                        .ToString("yyyy-MM-ddTHH:mm:ss.fff", CultureInfo.InvariantCulture) + "Z"
                    : Unavailable;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Build write time not available: {SecretRedactor.Redact(ex.Message)}");
                return Unavailable;
            }
        }

        private static string GetAssemblyVersion(Assembly addInAssembly)
        {
            return addInAssembly?.GetName()?.Version?.ToString() ?? Unavailable;
        }

        private static string GetFileVersion(Assembly addInAssembly, string location)
        {
            string fileVersion = null;
            try
            {
                if (File.Exists(location))
                {
                    fileVersion = FileVersionInfo.GetVersionInfo(location).FileVersion;
                }
            }
            catch (Exception ex)
            {
                // Fall back to the assembly attribute below.
                Debug.WriteLine($"File version not readable: {SecretRedactor.Redact(ex.Message)}");
            }

            if (string.IsNullOrEmpty(fileVersion))
            {
                fileVersion = addInAssembly.GetCustomAttribute<AssemblyFileVersionAttribute>()?.Version;
            }

            return string.IsNullOrEmpty(fileVersion) ? GetAssemblyVersion(addInAssembly) : fileVersion;
        }
    }
}
