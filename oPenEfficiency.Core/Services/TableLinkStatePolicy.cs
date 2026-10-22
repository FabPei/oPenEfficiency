using System;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free state contract for a native PowerPoint table linked to an Excel range
    /// (QA-S9 external-data wave): the "oPE_TableLink|flags|path|sheet|range" AlternativeText
    /// codec, including the optional flags token, the legacy flag-less format and sheet names
    /// that themselves contain pipes. Production implementation used by TableLink, which the
    /// Excel-side creation and the PowerPoint-side refresh share.
    /// </summary>
    public static class TableLinkStatePolicy
    {
        public const string Tag = "oPE_TableLink";

        /// <summary>
        /// The flags token: 'F' = linked with formatting, 'R' = auto-resize on refresh,
        /// "V" = neither. Empty input yields "V".
        /// </summary>
        public static string BuildFlags(bool formatted, bool autoResize)
        {
            string flags = (formatted ? "F" : "") + (autoResize ? "R" : "");
            return flags.Length == 0 ? "V" : flags;
        }

        /// <summary>
        /// Whether a token can be the optional flags field: at most three letters from
        /// {'V','F','R'}. A workbook path always contains ':' or a slash, so it can never
        /// look like a flags token; this cleanly distinguishes new tags from the old
        /// flag-less format.
        /// </summary>
        public static bool IsFlagsToken(string token)
        {
            if (string.IsNullOrEmpty(token) || token.Length > 3) return false;
            foreach (char ch in token)
            {
                if (ch != 'V' && ch != 'F' && ch != 'R') return false;
            }
            return true;
        }

        /// <summary>Parses a flags token into its two decisions.</summary>
        public static void ParseFlagsToken(string token, out bool formatted, out bool autoResize)
        {
            formatted = token.IndexOf('F') >= 0;
            autoResize = token.IndexOf('R') >= 0;
        }

        /// <summary>
        /// The historical lenient tag parse: requires the tag prefix and at least a path,
        /// sheet and range; an optional flags token directly after the tag selects the new
        /// format; the range is always the last field and the sheet rejoins every pipe in
        /// between.
        /// </summary>
        public static bool TryParseTag(
            string altText,
            out string path,
            out string sheet,
            out string range,
            out bool formatted,
            out bool autoResize)
        {
            path = null;
            sheet = null;
            range = null;
            formatted = false;
            autoResize = false;
            if (string.IsNullOrEmpty(altText)) return false;

            string[] parts = altText.Split('|');
            if (parts.Length < 4 || parts[0] != Tag) return false;

            int pathIdx = 1;
            if (parts.Length >= 5 && IsFlagsToken(parts[1]))
            {
                ParseFlagsToken(parts[1], out formatted, out autoResize);
                pathIdx = 2;
            }

            path = parts[pathIdx];
            range = parts[parts.Length - 1];
            sheet = string.Join("|", parts, pathIdx + 1, parts.Length - pathIdx - 2);
            return true;
        }
    }
}
