using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.Command.Tests
{
    [TestClass]
    public class DiagnosticBundleServiceTests
    {
        private string _testDirectory;

        [TestInitialize]
        public void Initialize()
        {
            _testDirectory = Path.Combine(Path.GetTempPath(), "oPenEfficiency-diagnostics-tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_testDirectory);
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (Directory.Exists(_testDirectory)) Directory.Delete(_testDirectory, recursive: true);
        }

        [TestMethod]
        public void Create_ExportsUsefulMetadataWithoutSecretsOrFullPaths()
        {
            const string secret = "diagnostic-test-secret-938475";
            string logPath = Path.Combine(_testDirectory, "error.log");
            string sensitiveDocumentPath = @"C:\Users\Example\Documents\Board Review.pptx";
            File.WriteAllText(
                logPath,
                "[2026-09-03 10:00:00] [BtnAlignLeft] ERROR: Failed at " + sensitiveDocumentPath +
                Environment.NewLine + "Authorization: Bearer " + secret +
                Environment.NewLine + "Host " + Environment.MachineName + " user " + Environment.UserName,
                Encoding.UTF8);

            var config = new SidebarConfig
            {
                AppTheme = "Dark",
                SelectedTranslationTool = "OpenAI",
                OpenAIApiKey = secret,
                LibraryPaths = new List<string> { @"C:\Sensitive\Assets" },
                Sections = new List<SidebarSection>
                {
                    new SidebarSection
                    {
                        Name = "Client Name",
                        Features = new List<SidebarFeature>
                        {
                            new SidebarFeature { Id = "BtnAlignLeft", Name = "Align Left" }
                        }
                    }
                }
            };
            var runtime = new DiagnosticRuntimeInfo
            {
                OpenEfficiencyVersion = "1.2.3.4",
                OfficeVersion = "16.0",
                OfficeBitness = "x64",
                WindowsVersion = "Windows test",
                DotNetRuntimeVersion = "4.8-test",
                VstoRuntimeVersion = "10.0-test",
                AddInLoaded = true
            };
            string bundlePath = Path.Combine(_testDirectory, "diagnostics.zip");

            new DiagnosticBundleService().Create(
                bundlePath,
                config,
                runtime,
                logPath,
                new DateTimeOffset(2026, 9, 3, 10, 30, 0, TimeSpan.Zero));

            string manifest = ReadEntry(bundlePath, "diagnostics.json");
            string log = ReadEntry(bundlePath, "logs/recent-errors.log");
            string allContent = manifest + log + ReadEntry(bundlePath, "README.txt");

            StringAssert.Contains(manifest, "BtnAlignLeft");
            StringAssert.Contains(manifest, "\"HasOpenAICredential\":true");
            StringAssert.Contains(manifest, "\"LibraryPathCount\":1");
            StringAssert.Contains(manifest, "\"OfficeBitness\":\"x64\"");
            StringAssert.Contains(log, "[REDACTED]");
            StringAssert.Contains(log, "[path]");
            Assert.IsFalse(allContent.Contains(secret), "Credential value leaked into diagnostic bundle.");
            Assert.IsFalse(allContent.Contains(sensitiveDocumentPath), "Full document path leaked into diagnostic bundle.");
            Assert.IsFalse(allContent.Contains(@"C:\Sensitive\Assets"), "Configured library path leaked into diagnostic bundle.");
            Assert.IsFalse(allContent.Contains("Client Name"), "User-defined section name leaked into diagnostic bundle.");
            if (Environment.MachineName.Length >= 3)
                Assert.IsFalse(allContent.IndexOf(Environment.MachineName, StringComparison.OrdinalIgnoreCase) >= 0, "Machine identity leaked into diagnostic bundle.");
            if (Environment.UserName.Length >= 3)
                Assert.IsFalse(allContent.IndexOf(Environment.UserName, StringComparison.OrdinalIgnoreCase) >= 0, "User identity leaked into diagnostic bundle.");
        }

        [TestMethod]
        public void Create_TruncatesLogsToBoundedRecentTail()
        {
            string logPath = Path.Combine(_testDirectory, "large.log");
            string oldMarker = "OLD-SENSITIVE-CONTEXT";
            string recentMarker = "RECENT-DIAGNOSTIC-CONTEXT";
            File.WriteAllText(
                logPath,
                oldMarker + Environment.NewLine + new string('x', DiagnosticBundleService.MaximumLogBytes + 1024) +
                Environment.NewLine + recentMarker,
                Encoding.UTF8);
            string bundlePath = Path.Combine(_testDirectory, "diagnostics.zip");

            new DiagnosticBundleService().Create(
                bundlePath,
                new SidebarConfig(),
                new DiagnosticRuntimeInfo(),
                logPath);

            string log = ReadEntry(bundlePath, "logs/recent-errors.log");
            Assert.IsFalse(log.Contains(oldMarker));
            StringAssert.Contains(log, recentMarker);
            Assert.IsTrue(Encoding.UTF8.GetByteCount(log) <= DiagnosticBundleService.MaximumLogBytes);
        }

        [TestMethod]
        public void Create_OmitsLogEntryWhenNoLogExists()
        {
            string bundlePath = Path.Combine(_testDirectory, "diagnostics.zip");

            new DiagnosticBundleService().Create(
                bundlePath,
                null,
                null,
                Path.Combine(_testDirectory, "missing.log"));

            using (var archive = ZipFile.OpenRead(bundlePath))
            {
                Assert.IsNotNull(archive.GetEntry("diagnostics.json"));
                Assert.IsNotNull(archive.GetEntry("README.txt"));
                Assert.IsNull(archive.GetEntry("logs/recent-errors.log"));
            }
        }

        [TestMethod]
        public void Create_ReplacesConfirmedDestinationWithoutLeavingTemporaryFiles()
        {
            string bundlePath = Path.Combine(_testDirectory, "diagnostics.zip");
            File.WriteAllText(bundlePath, "previous bundle");

            new DiagnosticBundleService().Create(
                bundlePath,
                new SidebarConfig(),
                new DiagnosticRuntimeInfo { OfficeBitness = "x86" });

            StringAssert.Contains(ReadEntry(bundlePath, "diagnostics.json"), "\"OfficeBitness\":\"x86\"");
            Assert.AreEqual(1, Directory.GetFiles(_testDirectory).Length, "Temporary bundle file was not cleaned up.");
        }

        private static string ReadEntry(string archivePath, string entryName)
        {
            using (var archive = ZipFile.OpenRead(archivePath))
            {
                var entry = archive.GetEntry(entryName);
                Assert.IsNotNull(entry, "Expected ZIP entry " + entryName);
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                    return reader.ReadToEnd();
            }
        }
    }
}
