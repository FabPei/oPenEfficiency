using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.Services
{
    /// <summary>One imported file as recorded in the user library's license trail.</summary>
    [DataContract]
    public class IllustrationImportRecord
    {
        [DataMember]
        public string FileName { get; set; }

        [DataMember]
        public string RelativePath { get; set; }

        [DataMember]
        public string SourceId { get; set; }

        [DataMember]
        public string SourceName { get; set; }

        [DataMember]
        public string LicenseName { get; set; }

        [DataMember]
        public string LicenseUrl { get; set; }

        [DataMember]
        public bool AttributionRequired { get; set; }

        [DataMember]
        public bool Decomposable { get; set; }

        [DataMember]
        public string ImportedUtc { get; set; }

        [DataMember]
        public string Sha256 { get; set; }
    }

    [DataContract]
    public class IllustrationImportCatalog
    {
        [DataMember]
        public string Version { get; set; } = "1.0";

        [DataMember]
        public List<IllustrationImportRecord> Imports { get; set; } = new List<IllustrationImportRecord>();
    }

    /// <summary>Outcome of one import run, sized for a plain-language summary.</summary>
    public sealed class IllustrationImportResult
    {
        public int Accepted { get; set; }

        public int Skipped { get; set; }

        public int Rejected { get; set; }

        public List<string> Messages { get; } = new List<string>();

        public string Folder { get; set; }

        public int Total
        {
            get { return Accepted + Skipped + Rejected; }
        }
    }

    /// <summary>
    /// Stores the SVG files the user fetched from a source listed in the Illustration Library window.
    /// The add-in never downloads anything: it screens the picked files, copies the accepted ones below
    /// %LOCALAPPDATA%\oPenEfficiency\Illustrations and keeps a license trail next to them so the source
    /// of every file stays traceable.
    /// </summary>
    public sealed class IllustrationImportService
    {
        public const string UnsortedSourceId = "unsorted";

        public const string CatalogFileName = "imported-manifest.json";

        /// <summary>User library location; per user, no administrator rights and outside the installer.</summary>
        public static string ResolveLibraryRoot()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "oPenEfficiency",
                "Illustrations");
        }

        public string LibraryRoot
        {
            get { return ResolveLibraryRoot(); }
        }

        public IllustrationImportResult Import(IEnumerable<string> files, IllustrationSource source)
        {
            if (files == null) throw new ArgumentNullException(nameof(files));

            string sourceId = source == null || string.IsNullOrWhiteSpace(source.Id) ? UnsortedSourceId : source.Id;
            string targetFolder = Path.Combine(LibraryRoot, sourceId);
            Directory.CreateDirectory(targetFolder);

            var result = new IllustrationImportResult { Folder = targetFolder };
            var catalog = LoadCatalog();

            foreach (string file in files.Where(path => !string.IsNullOrWhiteSpace(path)))
            {
                string fileName = Path.GetFileName(file);
                string text;
                try
                {
                    text = File.ReadAllText(file);
                }
                catch (Exception ex)
                {
                    result.Rejected++;
                    AddMessage(result, fileName + ": could not be read (" + ex.Message + ")");
                    continue;
                }

                var decision = IllustrationImportPolicy.Evaluate(fileName, text);
                if (!decision.Accepted)
                {
                    result.Rejected++;
                    AddMessage(result, fileName + ": " + decision.Reason);
                    continue;
                }

                string targetName = ResolveTargetName(targetFolder, fileName, text, result);
                if (targetName == null)
                {
                    result.Skipped++;
                    continue;
                }

                string targetPath = Path.Combine(targetFolder, targetName);
                File.WriteAllText(targetPath, text, new UTF8Encoding(false));

                catalog.Imports.Add(new IllustrationImportRecord
                {
                    FileName = targetName,
                    RelativePath = sourceId + "/" + targetName,
                    SourceId = sourceId,
                    SourceName = source == null ? string.Empty : source.Name,
                    LicenseName = source == null ? string.Empty : source.LicenseName,
                    LicenseUrl = source == null ? string.Empty : source.LicenseUrl,
                    AttributionRequired = source != null && source.AttributionRequired,
                    Decomposable = decision.Decomposable,
                    ImportedUtc = DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'"),
                    Sha256 = ComputeSha256(text)
                });

                result.Accepted++;
            }

            if (result.Accepted > 0) SaveCatalog(catalog);
            return result;
        }

        /// <summary>All files already sitting in the user library, newest import trail first.</summary>
        public IReadOnlyList<IllustrationImportRecord> ListImported()
        {
            return LoadCatalog().Imports
                .OrderByDescending(record => record.ImportedUtc, StringComparer.Ordinal)
                .ToList();
        }

        /// <summary>Opens a web page or a folder in the shell so the user can fetch artwork themselves.</summary>
        public static void OpenUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return;

            if (url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                url.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                Process.Start(url);
                return;
            }

            Process.Start("explorer.exe", "\"" + url + "\"");
        }

        private static void AddMessage(IllustrationImportResult result, string message)
        {
            if (result.Messages.Count < 8) result.Messages.Add(message);
        }

        /// <summary>
        /// Returns the file name to write, or null when the same content is already imported. A name
        /// clash with different content gets a numeric suffix instead of overwriting the earlier file.
        /// </summary>
        private static string ResolveTargetName(string folder, string fileName, string text, IllustrationImportResult result)
        {
            string candidate = fileName;
            for (int index = 1; index < 100; index++)
            {
                string path = Path.Combine(folder, candidate);
                if (!File.Exists(path)) return candidate;

                try
                {
                    if (string.Equals(File.ReadAllText(path), text, StringComparison.Ordinal)) return null;
                }
                catch (IOException)
                {
                    // Unreadable existing file: fall through and pick the next free name.
                }

                string stem = Path.GetFileNameWithoutExtension(fileName);
                candidate = stem + " (" + (index + 1) + ").svg";
            }

            result.Messages.Add(fileName + ": too many files with the same name in the library");
            return null;
        }

        private IllustrationImportCatalog LoadCatalog()
        {
            string path = Path.Combine(LibraryRoot, CatalogFileName);
            if (!File.Exists(path)) return new IllustrationImportCatalog();

            try
            {
                using (var stream = File.OpenRead(path))
                {
                    var serializer = new DataContractJsonSerializer(typeof(IllustrationImportCatalog));
                    var catalog = (IllustrationImportCatalog)serializer.ReadObject(stream);
                    if (catalog == null) return new IllustrationImportCatalog();
                    if (catalog.Imports == null) catalog.Imports = new List<IllustrationImportRecord>();
                    return catalog;
                }
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationImportService.LoadCatalog");
                return new IllustrationImportCatalog();
            }
        }

        private void SaveCatalog(IllustrationImportCatalog catalog)
        {
            string path = Path.Combine(LibraryRoot, CatalogFileName);
            using (var stream = File.Create(path))
            {
                var serializer = new DataContractJsonSerializer(typeof(IllustrationImportCatalog));
                serializer.WriteObject(stream, catalog);
            }
        }

        private static string ComputeSha256(string text)
        {
            using (var sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(text));
                var builder = new StringBuilder(hash.Length * 2);
                foreach (byte value in hash) builder.Append(value.ToString("x2"));
                return builder.ToString();
            }
        }
    }
}
