using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace oPenEfficiency.Models
{
    /// <summary>Compatibility boundary between token editors and comma-delimited persisted settings.</summary>
    public static class DelimitedSettingTokens
    {
        private static readonly Regex HexColorPattern = new Regex(
            "^#(?:[0-9A-F]{6}|[0-9A-F]{8})$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        public static IReadOnlyList<string> Parse(string serialized)
        {
            return (serialized ?? string.Empty)
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(value => value.Trim())
                .Where(value => value.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        public static string Serialize(IEnumerable<string> tokens)
        {
            return string.Join(", ", (tokens ?? Enumerable.Empty<string>())
                .Select(value => (value ?? string.Empty).Trim())
                .Where(value => value.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase));
        }

        public static bool TryNormalizeHexColor(string candidate, out string normalized)
        {
            normalized = (candidate ?? string.Empty).Trim().ToUpperInvariant();
            return HexColorPattern.IsMatch(normalized);
        }

        public static bool TryNormalizeFontName(string candidate, out string normalized)
        {
            normalized = (candidate ?? string.Empty).Trim();
            return normalized.Length > 0 && normalized.Length <= 100 && normalized.IndexOf(',') < 0;
        }
    }
}
