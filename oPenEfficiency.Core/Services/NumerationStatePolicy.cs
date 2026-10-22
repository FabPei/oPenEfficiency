using System;
using System.Drawing;
using System.Globalization;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free state contract for the numbered marker (QA-S9 wave 5): the persistent
    /// AlternativeText codec, the documented defaults with the 5..200 point size window and the
    /// 0..255 alpha clamp, and the replacement placement that centers a new marker on the
    /// predecessor shape. Production implementation used by NumerationFeature.
    /// </summary>
    public static class NumerationStatePolicy
    {
        public const string NumerationTag = "oPE_Numeration";
        public const float MinimumSize = 5f;
        public const float MaximumSize = 200f;

        /// <summary>The documented marker defaults: number 1, circle, 0.8cm, oPenEfficiency blue, opaque.</summary>
        public static void Defaults(
            out int number,
            out string shapeType,
            out float size,
            out string colorHex,
            out int alpha)
        {
            number = 1;
            shapeType = "Circle";
            size = 22.68f;
            colorHex = "#005C94";
            alpha = 255;
        }

        /// <summary>Whether a marker size lies in the documented 5..200 point window.</summary>
        public static bool IsValidSize(float size)
        {
            return size >= MinimumSize && size <= MaximumSize;
        }

        /// <summary>Clamps an opacity into 0..255.</summary>
        public static int ClampAlpha(int alpha)
        {
            return Math.Max(0, Math.Min(255, alpha));
        }

        /// <summary>
        /// Formats the persistent shape state; numbers are written invariant. A null shape type
        /// or color keeps the documented default in the tag, matching the historical codec.
        /// </summary>
        public static string FormatTag(int number, string shapeType, float size, string colorHex, int alpha)
        {
            return string.Join("|",
                NumerationTag,
                number.ToString(CultureInfo.InvariantCulture),
                shapeType ?? "Circle",
                size.ToString(CultureInfo.InvariantCulture),
                colorHex ?? "#005C94",
                alpha.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// The historical lenient state read: requires the tag prefix and starts from the
        /// documented defaults, then applies whichever fields are present. A size outside the
        /// 5..200 window, an unparsable color or a missing field keeps the default; a parsed
        /// alpha is clamped into 0..255.
        /// </summary>
        public static void ParseState(
            string tag,
            out int number,
            out string shapeType,
            out float size,
            out string colorHex,
            out int alpha)
        {
            Defaults(out number, out shapeType, out size, out colorHex, out alpha);
            if (string.IsNullOrEmpty(tag) ||
                !tag.StartsWith(NumerationTag + "|", StringComparison.Ordinal))
                return;

            string[] parts = tag.Split('|');
            if (parts.Length >= 2 &&
                int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedNumber))
                number = parsedNumber;
            if (parts.Length >= 3 && !string.IsNullOrWhiteSpace(parts[2]))
                shapeType = parts[2];
            if (parts.Length >= 4 &&
                float.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out float parsedSize) &&
                parsedSize >= MinimumSize && parsedSize <= MaximumSize)
                size = parsedSize;
            if (parts.Length >= 5 && !string.IsNullOrWhiteSpace(parts[4]))
            {
                try
                {
                    ColorTranslator.FromHtml(parts[4]);
                    colorHex = parts[4];
                }
                catch { /* An unparsable color keeps the documented default. */ }
            }
            if (parts.Length >= 6 &&
                int.TryParse(parts[5], NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedAlpha))
                alpha = ClampAlpha(parsedAlpha);
        }

        /// <summary>
        /// Places a replacement marker centered on the predecessor shape, so a restyled marker
        /// does not jump on the slide.
        /// </summary>
        public static void CenterOnShape(
            float existingLeft,
            float existingTop,
            float existingWidth,
            float existingHeight,
            float size,
            out float left,
            out float top)
        {
            left = existingLeft + (existingWidth / 2f) - (size / 2f);
            top = existingTop + (existingHeight / 2f) - (size / 2f);
        }
    }
}
