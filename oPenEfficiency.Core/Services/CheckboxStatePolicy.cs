using System;
using System.Globalization;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free state contract for the Checkbox indicator (QA-S9 wave 5): the persistent
    /// AlternativeText codec, the Wingdings symbol map, and the centered insert placement.
    /// Production implementation used by CheckboxFeature, the checkbox toolbar and ThisAddIn.
    /// </summary>
    /// <remarks>
    /// Culture safety: the codec writes the font size invariant; the lenient reader keeps the
    /// invariant parse and falls back to the current culture so decks authored before the
    /// invariant write (e.g. "24,5" on de-DE) keep opening with their real font size.
    /// </remarks>
    public static class CheckboxStatePolicy
    {
        public const string CheckboxTag = "oPE_Checkbox";

        /// <summary>The Wingdings font the tick/cross glyphs are written with.</summary>
        public const string SymbolFontName = "Wingdings";

        /// <summary>The side length of a freshly inserted checkbox, in points.</summary>
        public const float DefaultInsertSize = 50f;

        /// <summary>Whether an AlternativeText identifies a tagged checkbox shape.</summary>
        public static bool IsCheckboxTag(string alternativeText)
        {
            return alternativeText != null && alternativeText.StartsWith(CheckboxTag, StringComparison.Ordinal);
        }

        /// <summary>
        /// Formats the persistent shape state; the font size is written invariant so the tag is
        /// portable across locales. A colorRgb of -1 is the historical "keep the color" sentinel
        /// and is written verbatim.
        /// </summary>
        public static string FormatTag(int state, bool hasFrame, float fontSize, int colorRgb)
        {
            return CheckboxTag + "|" + state + "|" + hasFrame + "|" +
                fontSize.ToString(CultureInfo.InvariantCulture) + "|" + colorRgb.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// The historical lenient state read for a full "tag|state|frame|size|color" tag: fields
        /// that fail to parse fall back to zero values and a short or missing tag leaves the raw
        /// state text unset (the caller keeps its own defaults, as before). The font size parses
        /// invariant first and falls back to the current culture for legacy current-culture tags.
        /// </summary>
        public static void ParseTagFields(
            string alternativeText,
            out string stateText,
            out bool hasFrame,
            out float fontSize,
            out int colorRgb)
        {
            stateText = null;
            hasFrame = false;
            fontSize = 0f;
            colorRgb = 0;
            if (string.IsNullOrEmpty(alternativeText)) return;

            string[] parts = alternativeText.Split('|');
            if (parts.Length < 5) return;

            stateText = parts[1];
            bool.TryParse(parts[2], out hasFrame);
            if (!float.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out fontSize))
                float.TryParse(parts[3], NumberStyles.Float, CultureInfo.CurrentCulture, out fontSize);
            int.TryParse(parts[4], NumberStyles.Integer, CultureInfo.InvariantCulture, out colorRgb);
        }

        /// <summary>
        /// Maps a checkbox state to its Wingdings glyph: 0 = tick, 1 = cross, 2 = empty space.
        /// An unknown state keeps the historical empty glyph.
        /// </summary>
        public static string GetSymbol(int state)
        {
            switch (state)
            {
                case 0: return "ü"; // Wingdings tick
                case 1: return "û"; // Wingdings cross
                case 2: return " ";
                default: return "";
            }
        }

        /// <summary>Places a fresh checkbox at the center of the slide.</summary>
        public static void CenterOnSlide(float slideWidth, float slideHeight, float size, out float left, out float top)
        {
            left = (slideWidth / 2f) - (size / 2f);
            top = (slideHeight / 2f) - (size / 2f);
        }
    }
}
