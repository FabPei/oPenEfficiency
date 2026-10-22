using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free library rules behind the Template Library (QA-S9 wave 3): which files count as
    /// templates and how the three sources (registered files, watched directories, default
    /// locations) merge into the final list. Production implementation used by TemplateService.
    /// </summary>
    /// <remarks>
    /// The merge order is the contract: registered files first, then watched directories, then the
    /// default locations. The list de-duplicates by path case-insensitively keeping the FIRST
    /// occurrence (so a registered file wins over the same file found later by a scan), drops
    /// blank paths and removes everything on the case-insensitive blacklist.
    /// </remarks>
    public static class TemplateLibraryPolicy
    {
        public static readonly string[] SupportedExtensions = { ".potx", ".pot", ".pptx", ".pptm" };

        /// <summary>Whether a path is a supported template file (case-insensitive extension check).</summary>
        public static bool IsSupportedPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            string extension = Path.GetExtension(path);
            return SupportedExtensions.Any(ext => extension.Equals(ext, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Merges the three sources in priority order: registered, watched, defaults. Keeps the
        /// first occurrence per path (case-insensitive), drops blank paths and blacklist hits.
        /// </summary>
        public static List<string> MergePaths(
            IEnumerable<string> registered,
            IEnumerable<string> watched,
            IEnumerable<string> defaults,
            IEnumerable<string> blacklist)
        {
            var candidates = new List<string>();
            foreach (var source in new[] { registered, watched, defaults })
            {
                if (source == null) continue;
                candidates.AddRange(source.Where(path => !string.IsNullOrWhiteSpace(path)));
            }

            var blocked = new HashSet<string>(
                (blacklist ?? Enumerable.Empty<string>()).Where(path => !string.IsNullOrWhiteSpace(path)),
                StringComparer.OrdinalIgnoreCase);

            return candidates
                .Where(path => !blocked.Contains(path))
                .GroupBy(path => path, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToList();
        }
    }
}
