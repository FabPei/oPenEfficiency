using System;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Host-free matching rules of the Replace Formatting command (BtnReplaceFont). The command collects its format
    /// changes through this policy, so the documented contract - the 0.1 pt font-size tolerance, case-insensitive font
    /// matching and the no-op guards that keep a self-replacement out of the change list - is asserted against the
    /// production decision instead of a second test-only implementation.
    /// </summary>
    public static class BulkFormatMatchPolicy
    {
        /// <summary>Documented font-size tolerance: a run matches a 0.1 pt band around the configured find size.</summary>
        public const double SizeMatchTolerancePoints = 0.1;

        /// <summary>Find and replace sizes closer than this are the same formatting: no change is reported.</summary>
        public const double SizeNoOpTolerancePoints = 0.01;

        /// <summary>A run's font name matches case-insensitively, and a same-name replacement is a no-op.</summary>
        public static bool MatchesFont(string actualName, string findName, string replaceName)
        {
            return !string.IsNullOrWhiteSpace(findName)
                && !string.IsNullOrWhiteSpace(replaceName)
                && string.Equals(actualName, findName, StringComparison.OrdinalIgnoreCase)
                && !string.Equals(actualName, replaceName, StringComparison.Ordinal);
        }

        /// <summary>A run's font size matches within the documented tolerance, and a near-identical replacement is a no-op.</summary>
        public static bool MatchesSize(float actualSize, float? findSize, float? replaceSize)
        {
            return findSize.HasValue
                && replaceSize.HasValue
                && Math.Abs(actualSize - findSize.Value) < SizeMatchTolerancePoints
                && Math.Abs(actualSize - replaceSize.Value) >= SizeNoOpTolerancePoints;
        }

        /// <summary>Run colors: equal to the find value and different from the replace value.</summary>
        public static bool MatchesColor(int actualColor, int? findColor, int? replaceColor)
        {
            return MatchesColor(actualColor, findColor, replaceColor, true);
        }

        /// <summary>
        /// Fill and line colors additionally require the paint to be visible; an invisible fill or line can never
        /// match, and a replace value equal to the shape's own color is a no-op.
        /// </summary>
        public static bool MatchesColor(int actualColor, int? findColor, int? replaceColor, bool visible)
        {
            return visible
                && findColor.HasValue
                && replaceColor.HasValue
                && actualColor == findColor.Value
                && actualColor != replaceColor.Value;
        }
    }
}
