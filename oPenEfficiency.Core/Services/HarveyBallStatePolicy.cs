using System;
using System.Globalization;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free state contract for the Harvey Ball indicator: the value cycle
    /// (0 -> 0.25 -> 0.5 -> 0.75 -> 1.0 -> 0) and the shape-state codec that persists the
    /// percentage in the shape's <c>AlternativeText</c> as "oPE_HarveyBall|{percentage}".
    /// </summary>
    /// <remarks>
    /// This is the production implementation used by HarveyBallFeature and the sidebar context
    /// menu (QA-S9 wave 4). The codec is culture-invariant on write and accepts both the
    /// invariant and the legacy current-culture form on read, because the original
    /// <c>$"{tag}|{percentage}"</c> formatting wrote the decimal separator in the authoring
    /// machine's culture (e.g. "oPE_HarveyBall|0,25" on de-DE decks).
    /// </remarks>
    public static class HarveyBallStatePolicy
    {
        public const string HarveyBallTag = "oPE_HarveyBall";

        /// <summary>Returns the next percentage in the 0 -> 0.25 -> 0.5 -> 0.75 -> 1.0 -> 0 cycle.</summary>
        public static float NextPercentage(float current)
        {
            float clamped = Math.Max(0f, Math.Min(1f, current));
            float next = (float)Math.Round(clamped * 4f + 1f) / 4f;
            return next > 1f ? 0f : next;
        }

        /// <summary>Formats the persistent shape state, e.g. "oPE_HarveyBall|0.25".</summary>
        public static string FormatTag(float percentage)
        {
            return HarveyBallTag + "|" + percentage.ToString("0.####", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Parses a shape's AlternativeText. Accepts the invariant form and legacy
        /// decimal-comma forms; returns false when the text is not a Harvey Ball state.
        /// </summary>
        public static bool TryParseTag(string alternativeText, out float percentage)
        {
            percentage = 0f;
            if (string.IsNullOrEmpty(alternativeText) ||
                !alternativeText.StartsWith(HarveyBallTag, StringComparison.Ordinal))
            {
                return false;
            }

            var parts = alternativeText.Split('|');
            if (parts.Length != 2) return false;

            string value = parts[1];
            // Historical interpolation of a float uses its culture's decimal separator but
            // never adds thousands separators. Accept one decimal comma regardless of the
            // reader's culture. Legacy percentages are fractions below one; requiring that
            // range also prevents an ambiguous value such as "1,000" being read as either 1
            // or a thousands-grouped 1000.
            if (value.IndexOf(',') >= 0)
            {
                if (value.IndexOf('.') >= 0 || value.IndexOf(',') != value.LastIndexOf(',')) return false;

                string invariantValue = value.Replace(',', '.');
                return float.TryParse(invariantValue, NumberStyles.Float, CultureInfo.InvariantCulture, out percentage) &&
                    percentage >= 0f && percentage < 1f;
            }

            if (float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out percentage)) return true;
            return float.TryParse(value, NumberStyles.Float, CultureInfo.CurrentCulture, out percentage);
        }
    }
}
