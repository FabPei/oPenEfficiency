using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Feeds the Illustration Library window: it loads the bundled manifest from the add-in directory,
    /// derives the bundled items from the shipped SVG files, merges the files the user imported into their
    /// own library, and resolves the path and plan for a selected item.
    ///
    /// Bundled and imported items are kept apart by id and resolved through an explicit path map, so a file
    /// from the user library can never be mistaken for a shipped asset.
    /// </summary>
    public sealed class IllustrationCatalogService
    {
        private const string ManifestRelativePath = @"Resources\Illustrations\illustrations-manifest.json";

        private readonly IllustrationImportService _importService = new IllustrationImportService();
        private readonly Dictionary<string, string> _pathsById = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly Dictionary<string, IllustrationItem> _itemsById = new Dictionary<string, IllustrationItem>(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _collectionNames = new Dictionary<string, string>(StringComparer.Ordinal);

        private IllustrationManifest _manifest;
        private List<IllustrationItem> _items;

        /// <summary>Folder the bundled illustration assets are deployed to.</summary>
        public string AssetsRoot
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Resources\Illustrations"); }
        }

        public string ManifestPath
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ManifestRelativePath); }
        }

        public IllustrationManifest LoadManifest()
        {
            if (_manifest != null) return _manifest;

            if (!File.Exists(ManifestPath))
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(
                    "Illustration manifest is missing at " + ManifestPath,
                    "IllustrationCatalogService.LoadManifest");
                _manifest = new IllustrationManifest();
                return _manifest;
            }

            _manifest = IllustrationCatalog.LoadManifest(File.ReadAllText(ManifestPath));
            return _manifest;
        }

        /// <summary>Drops the cached item list so the next LoadItems picks up newly imported files.</summary>
        public void Invalidate()
        {
            _items = null;
            _pathsById.Clear();
            _itemsById.Clear();
        }

        /// <summary>Bundled items first, then the user's imported files; cached per session.</summary>
        public IReadOnlyList<IllustrationItem> LoadItems()
        {
            if (_items != null) return _items;

            var items = new List<IllustrationItem>();
            var manifest = LoadManifest();

            var build = IllustrationCatalog.Build(
                manifest,
                relativeFolder => ListSvgFiles(relativeFolder),
                relativePath => ReadSvgFile(relativePath));

            foreach (var item in build.Items)
            {
                Track(item, Path.Combine(AssetsRoot, item.RelativePath.Replace('/', Path.DirectorySeparatorChar)));
                items.Add(item);
            }

            foreach (var collection in manifest.Collections)
            {
                if (collection == null || string.IsNullOrWhiteSpace(collection.Id)) continue;
                _collectionNames[collection.Id] = string.IsNullOrWhiteSpace(collection.Name) ? collection.Id : collection.Name;
            }

            items.AddRange(LoadImportedItems());

            _items = items;
            return _items;
        }

        /// <summary>Collection filter values: bundled collection ids plus one entry per imported source.</summary>
        public IReadOnlyList<string> GetCollections(IEnumerable<IllustrationItem> items = null)
        {
            var source = (items ?? LoadItems())
                .Select(item => item.CollectionId)
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(id => id, StringComparer.Ordinal)
                .ToList();
            return source;
        }

        /// <summary>Readable collection label, e.g. "Open Peeps" or "Imported from Storyset".</summary>
        public string DescribeCollection(string collectionId)
        {
            if (string.IsNullOrWhiteSpace(collectionId)) return string.Empty;

            string name;
            if (_collectionNames.TryGetValue(collectionId, out name)) return name;

            const string importedPrefix = "imported-from-";
            if (collectionId.StartsWith(importedPrefix, StringComparison.Ordinal))
            {
                string sourceId = collectionId.Substring(importedPrefix.Length);
                var source = IllustrationSourceCatalog.Find(sourceId);
                return source == null ? "Imported (" + sourceId + ")" : "Imported from " + source.Name;
            }

            return collectionId;
        }

        public IReadOnlyList<IllustrationItem> Filter(IEnumerable<IllustrationItem> items, string query, string collectionId)
        {
            var filtered = items ?? LoadItems();

            if (!string.IsNullOrWhiteSpace(collectionId))
            {
                filtered = filtered.Where(item => string.Equals(item.CollectionId, collectionId, StringComparison.Ordinal));
            }

            return IllustrationCatalog.Search(filtered, query).ToList();
        }

        /// <summary>Absolute file for an item, or null when it has no resolvable file.</summary>
        public string ResolveItemPath(IllustrationItem item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.Id)) return null;

            string path;
            return _pathsById.TryGetValue(item.Id, out path) ? path : null;
        }

        public string LoadSvg(IllustrationItem item)
        {
            string path = ResolveItemPath(item);
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return null;

            try
            {
                return File.ReadAllText(path);
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationCatalogService.LoadSvg");
                return null;
            }
        }

        public IllustrationSvgPlan LoadPlan(IllustrationItem item)
        {
            string svg = LoadSvg(item);
            return string.IsNullOrWhiteSpace(svg) ? null : IllustrationSvgPlanner.Describe(svg);
        }

        private void Track(IllustrationItem item, string path)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.Id)) return;

            _itemsById[item.Id] = item;
            _pathsById[item.Id] = path;
        }

        private IList<string> ListSvgFiles(string relativeFolder)
        {
            string folder = Path.Combine(AssetsRoot, relativeFolder.Replace('/', Path.DirectorySeparatorChar));
            if (!Directory.Exists(folder)) return new List<string>();

            return Directory
                .GetFiles(folder, "*.svg", SearchOption.TopDirectoryOnly)
                .Select(path => relativeFolder.TrimEnd('/', '\\') + "/" + Path.GetFileName(path))
                .OrderBy(path => path, StringComparer.Ordinal)
                .ToList();
        }

        private string ReadSvgFile(string relativePath)
        {
            string path = Path.Combine(AssetsRoot, relativePath.Replace('/', Path.DirectorySeparatorChar));
            return File.Exists(path) ? File.ReadAllText(path) : null;
        }

        /// <summary>
        /// Items the user imported into %LOCALAPPDATA%\oPenEfficiency\Illustrations. They carry their source
        /// in the id and collection so the license trail stays attached to the artwork.
        /// </summary>
        private IEnumerable<IllustrationItem> LoadImportedItems()
        {
            var imported = new List<IllustrationItem>();

            foreach (var record in _importService.ListImported())
            {
                if (record == null || string.IsNullOrWhiteSpace(record.RelativePath)) continue;

                string path = Path.Combine(
                    _importService.LibraryRoot,
                    record.RelativePath.Replace('/', Path.DirectorySeparatorChar));
                if (!File.Exists(path)) continue;

                string text;
                try
                {
                    text = File.ReadAllText(path);
                }
                catch (Exception ex)
                {
                    oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationCatalogService.LoadImportedItems");
                    continue;
                }

                var plan = IllustrationSvgPlanner.Describe(text);
                if (!plan.IsSafe) continue;

                string sourceId = string.IsNullOrWhiteSpace(record.SourceId)
                    ? IllustrationImportService.UnsortedSourceId
                    : record.SourceId;

                var item = new IllustrationItem
                {
                    Id = "imported:" + sourceId + "/" + record.FileName,
                    Name = IllustrationSvgPlanner.DisplayName(Path.GetFileNameWithoutExtension(record.FileName)),
                    CollectionId = "imported-from-" + sourceId,
                    RelativePath = record.RelativePath,
                    Palette = plan.Palette,
                    Decomposable = plan.IsDecomposable,
                    Tags = new List<string> { "imported", sourceId }
                };

                if (plan.Parts != null)
                {
                    item.Parts = plan.Parts
                        .Where(part => !part.IsPlaceholder)
                        .Select(part => part.Id)
                        .ToList();
                }

                Track(item, path);
                imported.Add(item);
            }

            return imported;
        }
    }
}
