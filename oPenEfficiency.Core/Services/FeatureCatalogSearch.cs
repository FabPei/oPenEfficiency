using System;
using System.Collections.Generic;
using System.Linq;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Services
{
    public static class FeatureCatalogSearch
    {
        public static bool Matches(FeatureCatalogRecord record, string query)
        {
            return GetMatchScore(record, query) < int.MaxValue;
        }

        public static int GetMatchScore(FeatureCatalogRecord record, string query)
        {
            if (record == null) return int.MaxValue;

            string normalizedQuery = query?.Trim();
            if (string.IsNullOrEmpty(normalizedQuery)) return 0;

            var aliases = record.SearchAliases ?? Array.Empty<string>();
            if (EqualsValue(record.Name, normalizedQuery) ||
                EqualsValue(record.Id, normalizedQuery)) return 0;
            if (aliases.Any(alias => EqualsValue(alias, normalizedQuery))) return 1;
            if (StartsWithValue(record.Name, normalizedQuery)) return 2;
            if (aliases.Any(alias => StartsWithValue(alias, normalizedQuery))) return 3;
            if (StartsWithValue(record.Id, normalizedQuery)) return 4;
            if (ContainsValue(record.Name, normalizedQuery)) return 5;
            if (aliases.Any(alias => ContainsValue(alias, normalizedQuery))) return 6;
            if (ContainsValue(record.Id, normalizedQuery)) return 7;
            if (ContainsValue(record.Description, normalizedQuery)) return 8;
            if (ContainsValue(record.DetailedHelpText, normalizedQuery)) return 9;
            return int.MaxValue;
        }

        public static FeatureCatalogRecord ResolveByNameOrAlias(
            IEnumerable<FeatureCatalogRecord> records,
            string nameOrAlias)
        {
            string query = nameOrAlias?.Trim();
            if (string.IsNullOrEmpty(query)) return null;

            return (records ?? Enumerable.Empty<FeatureCatalogRecord>())
                .Where(record => record != null)
                .Select(record => new
                {
                    Record = record,
                    Score = GetMatchScore(record, query)
                })
                .Where(match => match.Score <= 1)
                .OrderBy(match => match.Score)
                .ThenBy(match => match.Record.Id, StringComparer.Ordinal)
                .Select(match => match.Record)
                .FirstOrDefault();
        }

        private static bool EqualsValue(string value, string query)
        {
            return string.Equals(value, query, StringComparison.OrdinalIgnoreCase);
        }

        private static bool StartsWithValue(string value, string query)
        {
            return value?.StartsWith(query, StringComparison.OrdinalIgnoreCase) == true;
        }

        private static bool ContainsValue(string value, string query)
        {
            return value?.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
