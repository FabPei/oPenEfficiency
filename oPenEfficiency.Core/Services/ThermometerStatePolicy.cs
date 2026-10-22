using System;
using System.Globalization;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free state contract for the Thermometer indicator (QA-S9 wave 4): the fill-percentage
    /// cycle (0.25 steps, wrapping 1.0 -> 0), the clamping rule and the persistent shape-state
    /// codec "{tag}|{percentage}|{colorRgb}". Production implementation used by ThermometerFeature.
    /// </summary>
    /// <remarks>
    /// The historical tag literal is "oPen_Thermometer" (not "oPE_Thermometer") and is preserved
    /// verbatim for saved-deck compatibility.
    /// </remarks>
    public static class ThermometerStatePolicy
    {
        public const string ThermometerTag = "oPen_Thermometer";

        /// <summary>Clamps a fill percentage into 0..1.</summary>
        public static float ClampPercentage(float percentage)
        {
            return Math.Max(0f, Math.Min(1f, percentage));
        }

        /// <summary>Returns the next cycle value: +0.25 steps, wrapping 1.0 back to 0.</summary>
        public static float NextPercentage(float current)
        {
            return current >= 1f ? 0f : Math.Min(1f, current + 0.25f);
        }

        /// <summary>Formats the persistent shape state; percentage is written invariant.</summary>
        public static string FormatTag(float percentage, int colorRgb)
        {
            return ThermometerTag + "|" + percentage.ToString(CultureInfo.InvariantCulture) + "|" + colorRgb;
        }

        /// <summary>
        /// The historical lenient state read: starts from the documented defaults (0.5, 0x0000FF),
        /// then applies whichever fields are present. The tag prefix is NOT required here because
        /// the feature accepts EE4P indicator shapes whose AlternativeText carries no oPE_ state;
        /// tag recognition stays in the feature's IsThermometer check.
        /// </summary>
        public static void ParseState(string alternativeText, out float percentage, out int colorRgb)
        {
            percentage = 0.5f;
            colorRgb = 0x0000FF;

            string[] parts = (alternativeText ?? string.Empty).Split('|');
            if (parts.Length >= 2 && float.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out float parsedPercentage))
                percentage = ClampPercentage(parsedPercentage);
            if (parts.Length >= 3 && int.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedColor))
                colorRgb = parsedColor;
        }
    }
}
