using System;
using System.Globalization;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Serializes and parses the position-only clipboard payload used by Copy/Paste XY-Coordinates.
    /// Current payloads are branded so destructive clipboard operations can prove ownership.
    /// </summary>
    public static class CoordinateClipboardPayload
    {
        private const string CurrentPrefix = "openEfficiency:XY:v1:";
        private const string FormatPrefix = "openEfficiency:XY:";

        public static bool TrySerializeCurrent(float x, float y, out string payload)
        {
            payload = null;
            if (!IsFinite(x) || !IsFinite(y)) return false;

            payload = string.Format(CultureInfo.InvariantCulture, CurrentPrefix + "{0:R},{1:R}", x, y);
            return true;
        }

        /// <summary>
        /// Parses valid current payloads and the retained markerless legacy two- and four-value formats.
        /// This method is intentionally permissive for paste compatibility and must not be used as an ownership check.
        /// </summary>
        public static bool TryParseForPaste(string payload, out float x, out float y)
        {
            x = 0f;
            y = 0f;
            if (string.IsNullOrWhiteSpace(payload)) return false;

            if (payload.StartsWith(CurrentPrefix, StringComparison.Ordinal))
            {
                return TryParseValues(payload.Substring(CurrentPrefix.Length), 2, out x, out y);
            }

            if (payload.StartsWith(FormatPrefix, StringComparison.Ordinal)) return false;

            return TryParseLegacyValues(payload, out x, out y);
        }

        /// <summary>
        /// Returns true only for a structurally and numerically valid current branded payload.
        /// Markerless legacy payloads remain paste-compatible but are never owned for clearing.
        /// </summary>
        public static bool IsOwnedCurrentPayload(string payload)
        {
            if (string.IsNullOrWhiteSpace(payload) ||
                !payload.StartsWith(CurrentPrefix, StringComparison.Ordinal))
            {
                return false;
            }

            return TryParseValues(payload.Substring(CurrentPrefix.Length), 2, out _, out _);
        }

        private static bool TryParseLegacyValues(string payload, out float x, out float y)
        {
            x = 0f;
            y = 0f;
            string[] parts = payload.Split(',');
            if (parts.Length != 2 && parts.Length != 4) return false;

            return TryParseValues(parts, out x, out y);
        }

        private static bool TryParseValues(string payload, int expectedValueCount, out float x, out float y)
        {
            x = 0f;
            y = 0f;
            string[] parts = payload.Split(',');
            if (parts.Length != expectedValueCount) return false;

            return TryParseValues(parts, out x, out y);
        }

        private static bool TryParseValues(string[] parts, out float x, out float y)
        {
            x = 0f;
            y = 0f;
            float[] values = new float[parts.Length];
            for (int index = 0; index < parts.Length; index++)
            {
                if (!float.TryParse(parts[index], NumberStyles.Float, CultureInfo.InvariantCulture, out values[index]) ||
                    !IsFinite(values[index]))
                {
                    return false;
                }
            }

            x = values[0];
            y = values[1];
            return true;
        }

        private static bool IsFinite(float value)
        {
            return !float.IsNaN(value) && !float.IsInfinity(value);
        }
    }
}
