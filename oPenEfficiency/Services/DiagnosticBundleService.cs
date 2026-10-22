using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using oPenEfficiency.Models;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Creates an intentionally small, local diagnostic package without presentation
    /// content, machine/user identity, configuration paths, or credentials.
    /// </summary>
    internal sealed class DiagnosticBundleService
    {
        internal const int MaximumLogBytes = 512 * 1024;

        private static readonly Regex WindowsPathRegex = new Regex(
            @"(?im)(?:\b[A-Z]:\\|\\\\)[^\r\n]+",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex FileUriRegex = new Regex(
            @"(?im)\bfile:/+[^\r\n]+",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex FeatureIdRegex = new Regex(
            @"^[A-Za-z][A-Za-z0-9._-]{0,99}$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        public void Create(
            string destinationPath,
            SidebarConfig config,
            DiagnosticRuntimeInfo runtime,
            string errorLogPath = null,
            DateTimeOffset? generatedAtUtc = null)
        {
            if (string.IsNullOrWhiteSpace(destinationPath))
                throw new ArgumentException("A destination path is required.", nameof(destinationPath));

            string fullDestinationPath = Path.GetFullPath(destinationPath);
            string destinationDirectory = Path.GetDirectoryName(fullDestinationPath);
            if (string.IsNullOrEmpty(destinationDirectory) || !Directory.Exists(destinationDirectory))
                throw new DirectoryNotFoundException("The destination folder does not exist.");

            config = config ?? new SidebarConfig();
            runtime = runtime ?? new DiagnosticRuntimeInfo();
            RegisterConfigurationSecrets(config);

            string temporaryPath = Path.Combine(
                destinationDirectory,
                "." + Path.GetFileName(fullDestinationPath) + "." + Guid.NewGuid().ToString("N") + ".tmp");

            try
            {
                using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: false))
                {
                    var manifest = BuildManifest(config, runtime, generatedAtUtc ?? DateTimeOffset.UtcNow);
                    WriteJsonEntry(archive, "diagnostics.json", manifest);
                    WriteTextEntry(
                        archive,
                        "README.txt",
                        "This bundle was created locally by openEfficiency. It excludes presentation content, " +
                        "credential values, full configuration paths, and machine/user identity. Review it before sharing.\r\n");

                    string sanitizedLog = ReadSanitizedLog(errorLogPath);
                    if (!string.IsNullOrWhiteSpace(sanitizedLog))
                        WriteTextEntry(archive, "logs/recent-errors.log", sanitizedLog);
                }

                if (File.Exists(fullDestinationPath))
                    File.Replace(temporaryPath, fullDestinationPath, null);
                else
                    File.Move(temporaryPath, fullDestinationPath);
            }
            finally
            {
                try
                {
                    if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "DiagnosticBundle.CleanupTemporaryFile");
                }
            }
        }

        private static DiagnosticBundleManifest BuildManifest(
            SidebarConfig config,
            DiagnosticRuntimeInfo runtime,
            DateTimeOffset generatedAtUtc)
        {
            var featureIds = (config.Sections ?? new List<SidebarSection>())
                .Where(section => section != null)
                .SelectMany(section => section.Features ?? new List<SidebarFeature>())
                .Where(feature => feature != null && FeatureIdRegex.IsMatch(feature.Id ?? string.Empty))
                .Select(feature => Sanitize(feature.Id))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(id => id, StringComparer.Ordinal)
                .ToList();

            return new DiagnosticBundleManifest
            {
                SchemaVersion = 1,
                GeneratedAtUtc = generatedAtUtc.UtcDateTime.ToString("O"),
                Runtime = new DiagnosticRuntimeInfo
                {
                    OpenEfficiencyVersion = Sanitize(runtime.OpenEfficiencyVersion),
                    OfficeVersion = Sanitize(runtime.OfficeVersion),
                    OfficeBitness = Sanitize(runtime.OfficeBitness),
                    WindowsVersion = Sanitize(runtime.WindowsVersion),
                    DotNetRuntimeVersion = Sanitize(runtime.DotNetRuntimeVersion),
                    VstoRuntimeVersion = Sanitize(runtime.VstoRuntimeVersion),
                    AddInLoaded = runtime.AddInLoaded
                },
                Configuration = new DiagnosticConfigurationInfo
                {
                    Theme = NormalizeKnownValue(config.AppTheme, new[] { "Dark", "Light" }, "Custom"),
                    TranslationProvider = NormalizeKnownValue(
                        config.SelectedTranslationTool,
                        new[] { "DeepL", "OpenAI", "Claude", "LibreTranslate" },
                        "Other"),
                    SectionCount = (config.Sections ?? new List<SidebarSection>()).Count,
                    ConfiguredFeatureCount = featureIds.Count,
                    LibraryPathCount = (config.LibraryPaths ?? new List<string>()).Count,
                    HasDeepLCredential = !string.IsNullOrWhiteSpace(config.DeepLApiKey),
                    HasOpenAICredential = !string.IsNullOrWhiteSpace(config.OpenAIApiKey),
                    HasClaudeCredential = !string.IsNullOrWhiteSpace(config.ClaudeApiKey),
                    HasLibreTranslateCredential = !string.IsNullOrWhiteSpace(config.LibreTranslateApiKey)
                },
                ConfiguredFeatureIds = featureIds
            };
        }

        private static string ReadSanitizedLog(string errorLogPath)
        {
            if (string.IsNullOrWhiteSpace(errorLogPath) || !File.Exists(errorLogPath)) return null;

            byte[] buffer;
            using (var stream = new FileStream(errorLogPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                int bytesToRead = (int)Math.Min(stream.Length, MaximumLogBytes);
                buffer = new byte[bytesToRead];
                stream.Seek(-bytesToRead, SeekOrigin.End);

                int offset = 0;
                while (offset < bytesToRead)
                {
                    int read = stream.Read(buffer, offset, bytesToRead - offset);
                    if (read == 0) break;
                    offset += read;
                }

                if (offset != buffer.Length) Array.Resize(ref buffer, offset);
            }

            string text = Encoding.UTF8.GetString(buffer).TrimStart('\uFEFF');
            if (text.Length > 0 && text[0] != '[' && text[0] != '\r' && text[0] != '\n')
            {
                int firstLineBreak = text.IndexOf('\n');
                if (firstLineBreak >= 0) text = text.Substring(firstLineBreak + 1);
            }

            return Sanitize(text);
        }

        private static string Sanitize(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;
            string redacted = SecretRedactor.Redact(value);
            redacted = FileUriRegex.Replace(redacted, "[path]");
            redacted = WindowsPathRegex.Replace(redacted, "[path]");
            redacted = ReplaceIdentity(redacted, Environment.UserName);
            redacted = ReplaceIdentity(redacted, Environment.MachineName);
            return ReplaceIdentity(redacted, Environment.UserDomainName);
        }

        private static string ReplaceIdentity(string value, string identity)
        {
            if (string.IsNullOrWhiteSpace(identity) || identity.Length < 3) return value;
            return Regex.Replace(
                value,
                Regex.Escape(identity),
                "[identity]",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static string NormalizeKnownValue(string value, IEnumerable<string> allowedValues, string fallback)
        {
            string match = allowedValues.FirstOrDefault(
                allowed => string.Equals(allowed, value, StringComparison.OrdinalIgnoreCase));
            return match ?? fallback;
        }

        private static void RegisterConfigurationSecrets(SidebarConfig config)
        {
            SecretRedactor.RegisterSecrets(new[]
            {
                config.DeepLApiKey,
                config.OpenAIApiKey,
                config.ClaudeApiKey,
                config.LibreTranslateApiKey
            });
        }

        private static void WriteJsonEntry(ZipArchive archive, string entryName, object value)
        {
            var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using (var stream = entry.Open())
            {
                var serializer = new DataContractJsonSerializer(value.GetType());
                serializer.WriteObject(stream, value);
            }
        }

        private static void WriteTextEntry(ZipArchive archive, string entryName, string value)
        {
            var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using (var writer = new StreamWriter(entry.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)))
                writer.Write(value);
        }

        internal static DiagnosticRuntimeInfo CaptureRuntime(string officeVersion, bool addInLoaded)
        {
            Assembly addInAssembly = typeof(DiagnosticBundleService).Assembly;
            Assembly vstoAssembly = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(
                assembly => assembly.GetName().Name.Equals(
                    "Microsoft.VisualStudio.Tools.Applications.Runtime",
                    StringComparison.OrdinalIgnoreCase));

            return new DiagnosticRuntimeInfo
            {
                OpenEfficiencyVersion = addInAssembly.GetName().Version?.ToString() ?? "Unavailable",
                OfficeVersion = string.IsNullOrWhiteSpace(officeVersion) ? "Unavailable" : officeVersion,
                OfficeBitness = Environment.Is64BitProcess ? "x64" : "x86",
                WindowsVersion = Environment.OSVersion.VersionString,
                DotNetRuntimeVersion = Environment.Version.ToString(),
                VstoRuntimeVersion = vstoAssembly?.GetName().Version?.ToString() ?? "Unavailable",
                AddInLoaded = addInLoaded
            };
        }
    }

    [DataContract]
    internal sealed class DiagnosticBundleManifest
    {
        [DataMember(Order = 1)] public int SchemaVersion { get; set; }
        [DataMember(Order = 2)] public string GeneratedAtUtc { get; set; }
        [DataMember(Order = 3)] public DiagnosticRuntimeInfo Runtime { get; set; }
        [DataMember(Order = 4)] public DiagnosticConfigurationInfo Configuration { get; set; }
        [DataMember(Order = 5)] public List<string> ConfiguredFeatureIds { get; set; }
    }

    [DataContract]
    internal sealed class DiagnosticRuntimeInfo
    {
        [DataMember(Order = 1)] public string OpenEfficiencyVersion { get; set; }
        [DataMember(Order = 2)] public string OfficeVersion { get; set; }
        [DataMember(Order = 3)] public string OfficeBitness { get; set; }
        [DataMember(Order = 4)] public string WindowsVersion { get; set; }
        [DataMember(Order = 5)] public string DotNetRuntimeVersion { get; set; }
        [DataMember(Order = 6)] public string VstoRuntimeVersion { get; set; }
        [DataMember(Order = 7)] public bool AddInLoaded { get; set; }
    }

    [DataContract]
    internal sealed class DiagnosticConfigurationInfo
    {
        [DataMember(Order = 1)] public string Theme { get; set; }
        [DataMember(Order = 2)] public string TranslationProvider { get; set; }
        [DataMember(Order = 3)] public int SectionCount { get; set; }
        [DataMember(Order = 4)] public int ConfiguredFeatureCount { get; set; }
        [DataMember(Order = 5)] public int LibraryPathCount { get; set; }
        [DataMember(Order = 6)] public bool HasDeepLCredential { get; set; }
        [DataMember(Order = 7)] public bool HasOpenAICredential { get; set; }
        [DataMember(Order = 8)] public bool HasClaudeCredential { get; set; }
        [DataMember(Order = 9)] public bool HasLibreTranslateCredential { get; set; }
    }
}
