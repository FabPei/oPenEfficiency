using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using oPenEfficiency.Models;

namespace oPenEfficiency.Services
{
    /// <summary>Outcome of a catalog build: usable items plus the assets that had to be left out.</summary>
    public sealed class IllustrationCatalogBuildResult
    {
        public List<IllustrationItem> Items { get; } = new List<IllustrationItem>();

        public List<string> MissingFiles { get; } = new List<string>();

        public List<string> UnsafeFiles { get; } = new List<string>();

        public bool IsClean
        {
            get { return MissingFiles.Count == 0 && UnsafeFiles.Count == 0; }
        }
    }

    /// <summary>
    /// Builds the illustration catalog from the bundled packs. The item list is derived from the SVG
    /// files themselves (named parts, palette, decomposability), so the manifest only has to carry the
    /// license and folder facts and the two can never drift apart.
    /// </summary>
    public static class IllustrationCatalog
    {
        public static IllustrationManifest LoadManifest(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) throw new ArgumentException("Illustration manifest json is empty.", nameof(json));

            using (var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(json)))
            {
                var serializer = new DataContractJsonSerializer(typeof(IllustrationManifest));
                var manifest = (IllustrationManifest)serializer.ReadObject(stream);
                if (manifest == null) throw new InvalidDataException("Illustration manifest could not be read.");
                if (manifest.Collections == null) manifest.Collections = new List<IllustrationCollectionInfo>();
                return manifest;
            }
        }

        /// <summary>
        /// Derives the items for every collection. <paramref name="listSvgFiles"/> returns the SVG paths
        /// relative to Resources\Illustrations for a collection folder, <paramref name="readSvg"/> their
        /// content; both are injected so the same rules can run against the shipped pack or a fixture.
        /// </summary>
        public static IllustrationCatalogBuildResult Build(
            IllustrationManifest manifest,
            Func<string, IList<string>> listSvgFiles,
            Func<string, string> readSvg)
        {
            var result = new IllustrationCatalogBuildResult();
            if (manifest == null || manifest.Collections == null) return result;

            foreach (var collection in manifest.Collections)
            {
                if (collection == null || string.IsNullOrWhiteSpace(collection.Id)) continue;
                if (string.IsNullOrWhiteSpace(collection.RelativeFolder)) continue;

                var files = listSvgFiles(collection.RelativeFolder) ?? new List<string>();

                foreach (var relativePath in files
                    .Where(path => path != null && path.EndsWith(".svg", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(path => path, StringComparer.Ordinal))
                {
                    string svgText = readSvg(relativePath);
                    if (string.IsNullOrEmpty(svgText))
                    {
                        result.MissingFiles.Add(relativePath);
                        continue;
                    }

                    var plan = IllustrationSvgPlanner.Describe(svgText);
                    if (!plan.IsSafe)
                    {
                        result.UnsafeFiles.Add(relativePath + ": " + plan.UnsafeReason);
                        continue;
                    }

                    string stem = Path.GetFileNameWithoutExtension(relativePath);
                    var item = new IllustrationItem
                    {
                        Id = collection.Id + "/" + stem,
                        Name = IllustrationSvgPlanner.DisplayName(stem),
                        CollectionId = collection.Id,
                        RelativePath = relativePath,
                        Parts = plan.Parts.Select(part => part.Id).ToList(),
                        Palette = plan.Palette,
                        Decomposable = plan.IsDecomposable && collection.SupportsPartDecomposition,
                        Tags = BuildTags(collection, plan)
                    };

                    result.Items.Add(item);
                }
            }

            result.Items.Sort((left, right) => string.CompareOrdinal(left.Id, right.Id));
            return result;
        }

        /// <summary>Case-insensitive token search across id, name, collection, tags and part ids.</summary>
        public static IList<IllustrationItem> Search(IEnumerable<IllustrationItem> items, string query)
        {
            var all = (items ?? Enumerable.Empty<IllustrationItem>()).ToList();
            if (string.IsNullOrWhiteSpace(query)) return all;

            var tokens = query.ToLowerInvariant().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return all.Where(item => tokens.All(token => item.SearchText.IndexOf(token, StringComparison.Ordinal) >= 0)).ToList();
        }

        private static List<string> BuildTags(IllustrationCollectionInfo collection, IllustrationSvgPlan plan)
        {
            var tags = new List<string>();
            if (collection != null && !string.IsNullOrWhiteSpace(collection.Name)) tags.Add(collection.Name.ToLowerInvariant());

            foreach (var part in plan.Parts.Where(part => !part.IsPlaceholder && IllustrationSvgPlanner.IsTaggableRole(part.Role)))
            {
                if (!string.IsNullOrWhiteSpace(part.Name) && !tags.Contains(part.Name.ToLowerInvariant()))
                {
                    tags.Add(part.Name.ToLowerInvariant());
                }
            }

            tags.Sort(StringComparer.Ordinal);
            return tags;
        }
    }
}
