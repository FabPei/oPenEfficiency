using System;
using System.Globalization;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free state contract for the Traffic Light indicator (QA-S9 wave 4): the persistent
    /// shape-state codec
    /// "oPE_TrafficLight|{state}|H/V|0/1|{borderWeight}" with strict validation and the legacy
    /// comma-tolerant float parse. Production implementation used by TrafficLightFeature; the
    /// feature maps its TrafficLightState enum to the persisted integer.
    /// </summary>
    public static class TrafficLightStatePolicy
    {
        public const string TrafficLightTag = "oPE_TrafficLight";
        public const string TrafficLightStateTag = "oPE_State";

        /// <summary>Serializes the persistent shape state exactly as the historical writer did.</summary>
        public static string Serialize(int state, bool horizontal, bool hasBorder, float borderWeight)
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0}|{1}|{2}|{3}|{4}",
                TrafficLightTag, state, horizontal ? "H" : "V", hasBorder ? 1 : 0, borderWeight);
        }

        /// <summary>
        /// Parses a persistent shape state with the historical strict validation: the tag must
        /// match exactly, state is an integer, orientation is H or V, the border flag is 0 or 1 and
        /// the weight accepts both the invariant and the legacy comma-decimal form.
        /// </summary>
        public static bool TryParse(
            string serializedState,
            out int state,
            out bool horizontal,
            out bool hasBorder,
            out float borderWeight)
        {
            state = 0; horizontal = true; hasBorder = false; borderWeight = 1f;
            if (string.IsNullOrEmpty(serializedState)) return false;

            var parts = serializedState.Split('|');
            if (parts.Length < 5 ||
                !string.Equals(parts[0], TrafficLightTag, StringComparison.Ordinal) ||
                !int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out state) ||
                (parts[2] != "H" && parts[2] != "V") ||
                (parts[3] != "0" && parts[3] != "1") ||
                !TryParsePersistedFloat(parts[4], out borderWeight) ||
                borderWeight <= 0f)
            {
                return false;
            }

            horizontal = parts[2] == "H";
            hasBorder = parts[3] == "1";
            return true;
        }

        /// <summary>Accepts the invariant form and the legacy comma-decimal form.</summary>
        public static bool TryParsePersistedFloat(string value, out float result)
        {
            return float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out result) ||
                   float.TryParse(value?.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out result);
        }
    }
}
