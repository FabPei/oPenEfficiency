using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// Pins the startup build marker: support reads this line to decide which assembly an
    /// Office session actually loaded, so the format may not drift silently.
    /// </summary>
    [TestClass]
    public sealed class StartupBuildLogTests
    {
        /// <summary>Pinned startup line format (field names, order and separators).</summary>
        private static readonly Regex LinePattern = new Regex(
            @"^\[(?<localTime>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\] BUILD"
            + @" addin=(?<addin>\S+)"
            + @" assembly=(?<assembly>\S+)"
            + @" assemblyVersion=(?<assemblyVersion>\S+)"
            + @" fileVersion=(?<fileVersion>\S+)"
            + @" builtUtc=(?<builtUtc>\S+)"
            + @" path=(?<path>.+)$",
            RegexOptions.CultureInvariant);

        private const string BuiltUtcFormat = "yyyy-MM-ddTHH:mm:ss.fff'Z'";

        private string _logPath;

        [TestInitialize]
        public void SetUp()
        {
            _logPath = Path.Combine(
                Path.GetTempPath(),
                "oPenEfficiency.Core.Tests",
                "startup_" + Guid.NewGuid().ToString("N") + ".log");
        }

        [TestCleanup]
        public void TearDown()
        {
            try
            {
                if (File.Exists(_logPath))
                {
                    File.Delete(_logPath);
                }
            }
            catch (IOException ex)
            {
                // A leftover test log must not fail the suite.
                Debug.WriteLine(ex.Message);
            }
        }

        [TestMethod]
        public void StartupLineIdentifiesTheLoadedBuildAndMatchesThePinnedFormat()
        {
            Assembly addInAssembly = Assembly.GetExecutingAssembly();

            string line = StartupBuildLog.Record(addInAssembly, "oPenEfficiency.Excel", _logPath);

            Assert.IsNotNull(line, "The startup line was not written.");
            string[] lines = File.ReadAllLines(_logPath);
            Assert.AreEqual(1, lines.Length, "A single start must write exactly one line.");
            Assert.AreEqual(line, lines[0]);

            Match match = LinePattern.Match(lines[0]);
            Assert.IsTrue(match.Success, $"The startup line no longer matches the pinned format: {lines[0]}");

            Assert.AreEqual("oPenEfficiency.Excel", match.Groups["addin"].Value);
            Assert.AreEqual(Path.GetFileName(addInAssembly.Location), match.Groups["assembly"].Value);
            Assert.AreEqual(addInAssembly.Location, match.Groups["path"].Value);
            Assert.AreEqual(
                addInAssembly.GetName().Version.ToString(),
                match.Groups["assemblyVersion"].Value);

            string fileVersion = match.Groups["fileVersion"].Value;
            Assert.IsTrue(
                Regex.IsMatch(fileVersion, @"^\d+(\.\d+){1,3}$", RegexOptions.CultureInvariant),
                $"fileVersion is not a version shape: {fileVersion}");
            string fileVersionOnDisk = FileVersionInfo.GetVersionInfo(addInAssembly.Location).FileVersion;
            if (!string.IsNullOrEmpty(fileVersionOnDisk))
            {
                Assert.AreEqual(fileVersionOnDisk, fileVersion);
            }

            Assert.AreEqual(
                File.GetLastWriteTimeUtc(addInAssembly.Location)
                    .ToString("yyyy-MM-ddTHH:mm:ss.fff", CultureInfo.InvariantCulture) + "Z",
                match.Groups["builtUtc"].Value,
                "builtUtc must be the write time of the loaded assembly file.");

            Assert.IsTrue(
                DateTime.TryParseExact(
                    match.Groups["builtUtc"].Value,
                    BuiltUtcFormat,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
                    out _),
                "builtUtc must be a parseable UTC timestamp.");

            Assert.IsTrue(
                DateTime.TryParseExact(
                    match.Groups["localTime"].Value,
                    "yyyy-MM-dd HH:mm:ss",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out _),
                "The line prefix must be a parseable local timestamp.");
        }

        [TestMethod]
        public void EachStartAppendsExactlyOneLine()
        {
            Assembly addInAssembly = Assembly.GetExecutingAssembly();

            Assert.IsNotNull(StartupBuildLog.Record(addInAssembly, "oPenEfficiency", _logPath));
            Assert.IsNotNull(StartupBuildLog.Record(addInAssembly, "oPenEfficiency", _logPath));

            string[] lines = File.ReadAllLines(_logPath);
            Assert.AreEqual(2, lines.Length);
            foreach (string line in lines)
            {
                Assert.IsTrue(LinePattern.IsMatch(line), $"Unparsable startup line: {line}");
            }
        }

        [TestMethod]
        public void LoggingFailureNeverThrows()
        {
            Assembly addInAssembly = Assembly.GetExecutingAssembly();

            Assert.IsNull(StartupBuildLog.Record(addInAssembly, "oPenEfficiency", "C:\\invalid\u0000path\\startup.log"));
            Assert.IsNull(StartupBuildLog.Record(null, "oPenEfficiency", _logPath));
            Assert.IsFalse(File.Exists(_logPath), "A rejected record must not create a log file.");
        }

        /// <summary>
        /// Pins the About line shown in the settings dialog: version, file version, build time and
        /// the loaded assembly file name.
        /// </summary>
        [TestMethod]
        public void AboutLineNamesVersionFileVersionBuildTimeAndFile()
        {
            string line = StartupBuildLog.FormatAboutLine(typeof(StartupBuildLog).Assembly);

            StringAssert.Matches(
                line,
                new Regex(
                    @"^Version \d+(\.\d+){3} \| file \S+ \| built \d{4}-\d{2}-\d{2} \d{2}:\d{2} \| oPenEfficiency\.Core\.dll$",
                    RegexOptions.CultureInvariant),
                $"Unexpected About line: {line}");
        }

        [TestMethod]
        public void BuildHintNamesTheBuildTime()
        {
            string hint = StartupBuildLog.FormatBuildHint(Assembly.GetExecutingAssembly());

            StringAssert.StartsWith(hint, "built ");
            Assert.IsTrue(
                DateTime.TryParseExact(
                    hint.Substring("built ".Length),
                    "yyyy-MM-dd HH:mm",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out _),
                $"The ribbon build hint is not parseable: {hint}");

            Assert.AreEqual(string.Empty, StartupBuildLog.FormatBuildHint(null));
        }
    }
}
