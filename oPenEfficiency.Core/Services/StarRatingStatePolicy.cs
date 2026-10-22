using System;
using System.Globalization;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free state contract for the Star Rating indicator (QA-S9 wave 4): the persistent
    /// shape-state codec "oPE_StarRating|{count}|{rating}|{iconMode}|{hasFrame}|{colorBgr}" and the
    /// per-star fill decision (full / dim / partial gradient with the >= 0.5 solid fallback).
    /// Production implementation used by StarRatingFeature.
    /// </summary>
    public static class StarRatingStatePolicy
    {
        public const string StarRatingTag = "oPE_StarRating";

        public enum FillKind
        {
            Dim,
            Full,
            Gradient
        }

        /// <summary>Formats the persistent shape state; rating is written invariant with one decimal.</summary>
        public static string FormatTag(int count, double rating, int iconMode, bool hasFrame, int colorBgr)
        {
            string ratingStr = rating.ToString("F1", CultureInfo.InvariantCulture);
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0}|{1}|{2}|{3}|{4}|{5}",
                StarRatingTag, count, ratingStr, iconMode, hasFrame ? 1 : 0, colorBgr);
        }

        /// <summary>Parses a persistent shape state; rejects foreign or malformed tags.</summary>
        public static bool TryParseTag(
            string alternativeText,
            out int count,
            out double rating,
            out int iconMode,
            out bool hasFrame,
            out int colorBgr)
        {
            count = 0; rating = 0; iconMode = 0; hasFrame = false; colorBgr = 0;
            if (string.IsNullOrEmpty(alternativeText)) return false;

            var parts = alternativeText.Split('|');
            if (parts.Length != 6 || !string.Equals(parts[0], StarRatingTag, StringComparison.Ordinal)) return false;
            if (!int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out count)) return false;
            if (!double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out rating)) return false;
            if (!int.TryParse(parts[3], NumberStyles.Integer, CultureInfo.InvariantCulture, out iconMode)) return false;
            if (parts[4] != "0" && parts[4] != "1") return false;
            hasFrame = parts[4] == "1";
            return int.TryParse(parts[5], NumberStyles.Integer, CultureInfo.InvariantCulture, out colorBgr);
        }

        /// <summary>
        /// Decides how star <paramref name="index"/> (0-based) is filled for the given rating:
        /// fully when the rating reaches the star's end, dim when it does not reach the star's start,
        /// otherwise a partial gradient with the historical solid fallback at >= 0.5.
        /// </summary>
        public static void ResolveFill(int index, double rating, out FillKind kind, out double fraction)
        {
            double starStart = index;
            double starEnd = index + 1;
            fraction = 0;

            if (rating >= starEnd)
            {
                kind = FillKind.Full;
            }
            else if (rating <= starStart)
            {
                kind = FillKind.Dim;
            }
            else
            {
                fraction = rating - starStart;
                kind = FillKind.Gradient;
            }
        }

        /// <summary>The historical solid-color fallback used when gradient rendering fails.</summary>
        public static bool UseSolidFallback(double fraction)
        {
            return fraction >= 0.5;
        }
    }
}
