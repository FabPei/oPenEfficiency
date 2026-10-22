using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace oPenEfficiency.Models
{
    public sealed class CountryFlagEntry
    {
        public CountryFlagEntry(string name, string alpha2, string alpha3, string assetId, IReadOnlyList<string> aliases)
        {
            Name = name;
            Alpha2 = alpha2;
            Alpha3 = alpha3;
            AssetId = assetId;
            Aliases = aliases ?? Array.Empty<string>();
            SearchText = Normalize(string.Join(" ", new[] { name, alpha2, alpha3 }.Concat(Aliases)));
        }

        public string Name { get; }
        public string Alpha2 { get; }
        public string Alpha3 { get; }
        public string AssetId { get; }
        public IReadOnlyList<string> Aliases { get; }
        internal string SearchText { get; }

        internal static string Normalize(string value)
        {
            var builder = new StringBuilder();
            foreach (char character in (value ?? string.Empty).Normalize(NormalizationForm.FormD))
            {
                if (CharUnicodeInfo.GetUnicodeCategory(character) != UnicodeCategory.NonSpacingMark)
                    builder.Append(char.ToLowerInvariant(character));
            }
            return builder.ToString().Normalize(NormalizationForm.FormC);
        }
    }

    public static partial class CountryFlagCatalog
    {
        public const string DatasetVersion = "ISO-3166-1 / flag-icons-7.3.2";
        private static readonly IReadOnlyList<CountryFlagEntry> SortedEntries;
        static CountryFlagCatalog()
        {
            SortedEntries = GeneratedEntries
                .OrderBy(entry => entry.Name, StringComparer.Ordinal)
                .ToList()
                .AsReadOnly();
        }
        public static IReadOnlyList<CountryFlagEntry> Entries => SortedEntries;

        public static IReadOnlyList<CountryFlagEntry> Search(string query)
        {
            string normalized = CountryFlagEntry.Normalize(query).Trim();
            if (normalized.Length == 0) return Entries;
            return Entries
                .Where(entry => entry.SearchText.Contains(normalized))
                .OrderBy(entry => IsExactMatch(entry, normalized) ? 0 : entry.SearchText.StartsWith(normalized, StringComparison.Ordinal) ? 1 : 2)
                .ThenBy(entry => entry.Name, StringComparer.Ordinal)
                .ToList()
                .AsReadOnly();
        }

        private static bool IsExactMatch(CountryFlagEntry entry, string normalized)
        {
            return CountryFlagEntry.Normalize(entry.Name) == normalized ||
                entry.Alpha2.Equals(normalized, StringComparison.OrdinalIgnoreCase) ||
                entry.Alpha3.Equals(normalized, StringComparison.OrdinalIgnoreCase) ||
                entry.Aliases.Any(alias => CountryFlagEntry.Normalize(alias) == normalized);
        }

        public static IReadOnlyList<CountryFlagEntry> ResolveRecent(IEnumerable<string> ids)
        {
            var byId = Entries.ToDictionary(entry => entry.AssetId, StringComparer.OrdinalIgnoreCase);
            var result = new List<CountryFlagEntry>();
            foreach (string id in ids ?? Enumerable.Empty<string>())
            {
                if (id != null && byId.TryGetValue(id, out var entry) && !result.Contains(entry)) result.Add(entry);
            }
            return result.AsReadOnly();
        }
    }
}
