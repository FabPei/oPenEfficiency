using System;
using System.Globalization;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free state contract for the Progress Series indicator (QA-S9 wave 4): the persistent
    /// shape-state codec
    /// "oPE_Series|{count}|{activeIndex}|{mode}|{layout}|{highlightColor}|{shapeType}" with the
    /// historical lenient partial parse and the documented defaults. Production implementation
    /// used by ProgressSeriesFeature.SeriesSettings.
    /// </summary>
    public static class ProgressSeriesStatePolicy
    {
        public const string SeriesTag = "oPE_Series";

        public const int DefaultCount = 5;
        public const int DefaultActiveIndex = 3;
        public const string DefaultMode = "Numbers";
        public const string DefaultLayout = "H";
        public const string DefaultHighlightColor = "#E53935";
        public const string DefaultShapeType = "Square";

        /// <summary>Formats the persistent shape state exactly as the historical writer did.</summary>
        public static string FormatTag(int count, int activeIndex, string mode, string layout, string highlightColor, string shapeType)
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0}|{1}|{2}|{3}|{4}|{5}|{6}",
                SeriesTag, count, activeIndex, mode, layout, highlightColor, shapeType);
        }

        /// <summary>
        /// The historical lenient parse: starts from the documented defaults and overwrites only
        /// the parts that are present (an unparsable integer becomes 0, matching the historical
        /// <c>int.TryParse</c> out semantics). Returns false for foreign or empty tags, in which
        /// case every output carries its default.
        /// </summary>
        public static bool TryParseTag(
            string tag,
            out int count,
            out int activeIndex,
            out string mode,
            out string layout,
            out string highlightColor,
            out string shapeType)
        {
            count = DefaultCount;
            activeIndex = DefaultActiveIndex;
            mode = DefaultMode;
            layout = DefaultLayout;
            highlightColor = DefaultHighlightColor;
            shapeType = DefaultShapeType;

            if (string.IsNullOrEmpty(tag) || !tag.StartsWith(SeriesTag, StringComparison.Ordinal)) return false;

            var parts = tag.Split('|');
            if (parts.Length >= 2) int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out count);
            if (parts.Length >= 3) int.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out activeIndex);
            if (parts.Length >= 4) mode = parts[3];
            if (parts.Length >= 5) layout = parts[4];
            if (parts.Length >= 6) highlightColor = parts[5];
            if (parts.Length >= 7) shapeType = parts[6];
            return true;
        }
    }
}
