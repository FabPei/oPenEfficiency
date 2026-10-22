using System;
using System.Globalization;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free decision rules behind the Font Panel widget (QA-S9 wave 3): whether a requested
    /// font size is valid at all, whether a range actually differs from the requested value (so an
    /// apply skips ranges that are already there), and how the typed size text is parsed.
    /// Production implementation used by FontPanel.
    /// </summary>
    /// <remarks>
    /// The size-change epsilon is inclusive on the no-change side: a difference of exactly 0.01
    /// counts as unchanged. Size parsing deliberately uses the current culture because the user
    /// types the value in their own locale (for example 12,5 on de-DE machines).
    /// </remarks>
    public static class FontEditPolicy
    {
        public const float SizeEpsilon = 0.01f;

        /// <summary>Font sizes are exclusive-bounded: 0 and 4000 or above are rejected.</summary>
        public static bool IsValidFontSize(float size)
        {
            return size > 0f && size < 4000f;
        }

        /// <summary>Whether a range's font name differs from the requested name (case-insensitive).</summary>
        public static bool NeedsNameChange(string current, string requested)
        {
            return !string.Equals(current, requested, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>Whether a range's font size differs from the requested size (within-epsilon counts as equal).</summary>
        public static bool NeedsSizeChange(float current, float requested)
        {
            return Math.Abs(current - requested) > SizeEpsilon;
        }

        /// <summary>Parses typed size text in the user's current culture.</summary>
        public static bool TryParseSize(string text, out float value)
        {
            return float.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value);
        }
    }
}
