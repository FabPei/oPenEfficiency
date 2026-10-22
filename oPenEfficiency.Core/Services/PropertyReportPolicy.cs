using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free report for Inspect Fonts &amp; Colors (QA-S9 dialogs/workflows wave): the
    /// COLORREF-to-hex conversion, the deduplication/normalization rules for the collected
    /// properties and the exact report layout. Production implementation used by
    /// PropertyExtractionFeature; the COM shape/text traversal stays in the feature.
    /// </summary>
    public static class PropertyReportPolicy
    {
        /// <summary>
        /// Formats a COLORREF (0x00BBGGRR) as "#RRGGBB".
        /// </summary>
        public static string ToHex(int colorRef)
        {
            int red = colorRef & 0xFF;
            int green = (colorRef >> 8) & 0xFF;
            int blue = (colorRef >> 16) & 0xFF;
            return $"#{red:X2}{green:X2}{blue:X2}";
        }

        /// <summary>Drops blank font names, deduplicates and sorts.</summary>
        public static List<string> NormalizeFonts(IEnumerable<string> fonts)
        {
            return fonts.Where(value => !string.IsNullOrWhiteSpace(value)).Distinct().OrderBy(value => value).ToList();
        }

        /// <summary>Drops non-positive sizes, deduplicates and sorts.</summary>
        public static List<float> NormalizeSizes(IEnumerable<float> sizes)
        {
            return sizes.Where(value => value > 0).Distinct().OrderBy(value => value).ToList();
        }

        /// <summary>Drops negative color values, deduplicates and sorts.</summary>
        public static List<int> NormalizeColors(IEnumerable<int> colors)
        {
            return colors.Where(value => value >= 0).Distinct().OrderBy(value => value).ToList();
        }

        /// <summary>
        /// The exact report layout: the scope line, one labeled line per property group
        /// ("None" when empty, sizes suffixed with " pt", colors as hex), trailing
        /// whitespace trimmed.
        /// </summary>
        public static string BuildReport(
            string scope,
            IEnumerable<string> fonts,
            IEnumerable<float> sizes,
            IEnumerable<int> textColors,
            IEnumerable<int> fillColors,
            IEnumerable<int> lineColors)
        {
            var report = new StringBuilder();
            report.AppendLine("Scope: " + scope);
            Append(report, "Fonts", fonts);
            Append(report, "Font sizes", sizes.Select(value => value.ToString("0.##") + " pt"));
            Append(report, "Text colors", textColors.Select(ToHex));
            Append(report, "Fill colors", fillColors.Select(ToHex));
            Append(report, "Line colors", lineColors.Select(ToHex));
            return report.ToString().TrimEnd();
        }

        private static void Append(StringBuilder report, string label, IEnumerable<string> values)
        {
            var items = values.ToList();
            report.AppendLine(label + ": " + (items.Count == 0 ? "None" : string.Join(", ", items)));
        }
    }
}
