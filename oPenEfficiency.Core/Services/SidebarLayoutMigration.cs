using System;
using System.Collections.Generic;
using System.Linq;
using oPenEfficiency.Models;

namespace oPenEfficiency.Services
{
    public sealed class SidebarLayoutMigrationResult
    {
        public bool Changed { get; internal set; }
        public int RefreshedNameCount { get; internal set; }
        public int RemovedObsoleteCount { get; internal set; }
        public int ConsolidatedLegacyCount { get; internal set; }
        public IReadOnlyList<string> UnknownFeatureIds { get; internal set; } = Array.Empty<string>();
    }

    public static class SidebarLayoutMigration
    {
        private static readonly IReadOnlyDictionary<string, string> LegacyIdMap =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["BtnAlignToTable"] = "BtnAlignToTableCell",
                ["BtnManageStickyNotes"] = "BtnStickyNoteManager",
                ["BtnMatchAngles"] = "BtnAlignShapeAdjustments"
            };

        private static readonly HashSet<string> ObsoleteIds =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "BtnSwapHorizontal",
                "BtnSwapVertical"
            };

        public static SidebarLayoutMigrationResult Migrate(
            SidebarConfig config,
            IEnumerable<SidebarFeature> catalogFeatures)
        {
            var result = new SidebarLayoutMigrationResult();
            if (config?.Sections == null) return result;

            var knownById = (catalogFeatures ?? Enumerable.Empty<SidebarFeature>())
                .Where(feature => feature != null && !string.IsNullOrWhiteSpace(feature.Id))
                .GroupBy(feature => feature.Id, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
            var unknownIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var section in config.Sections.Where(section => section != null))
            {
                if (section.Features == null) continue;

                for (int index = section.Features.Count - 1; index >= 0; index--)
                {
                    var saved = section.Features[index];
                    if (saved == null || ObsoleteIds.Contains(saved.Id))
                    {
                        section.Features.RemoveAt(index);
                        result.RemovedObsoleteCount++;
                        result.Changed = true;
                        continue;
                    }

                    if (LegacyIdMap.TryGetValue(saved.Id ?? string.Empty, out string currentId))
                    {
                        bool canonicalAlreadyPresent = section.Features
                            .Where((feature, candidateIndex) => candidateIndex != index && feature != null)
                            .Any(feature => string.Equals(feature.Id, currentId, StringComparison.OrdinalIgnoreCase));
                        if (canonicalAlreadyPresent)
                        {
                            section.Features.RemoveAt(index);
                            result.ConsolidatedLegacyCount++;
                            result.Changed = true;
                            continue;
                        }

                        saved.Id = currentId;
                        result.Changed = true;
                    }

                    if (!knownById.TryGetValue(saved.Id ?? string.Empty, out SidebarFeature current))
                    {
                        if (!string.IsNullOrWhiteSpace(saved.Id)) unknownIds.Add(saved.Id);
                        continue;
                    }

                    if (!string.Equals(saved.Id, current.Id, StringComparison.Ordinal))
                    {
                        saved.Id = current.Id;
                        result.Changed = true;
                    }

                    if (!string.Equals(saved.Name, current.Name, StringComparison.Ordinal))
                    {
                        saved.Name = current.Name;
                        result.RefreshedNameCount++;
                        result.Changed = true;
                    }
                }
            }

            result.UnknownFeatureIds = unknownIds.OrderBy(id => id, StringComparer.OrdinalIgnoreCase).ToList();
            return result;
        }
    }
}
