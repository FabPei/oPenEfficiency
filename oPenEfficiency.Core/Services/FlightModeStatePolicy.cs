namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free state contract for Master Covers & Overlays (QA-S9 dialogs/workflows wave):
    /// the persistent shape-tag vocabulary and the overlay transparency levels, including the
    /// boundary that classifies a persisted overlay as Light versus Heavy. Production
    /// implementation used by FlightModeFeature.
    /// </summary>
    public static class FlightModeStatePolicy
    {
        public const string FlightModeTag = "oPE_FlightMode";
        public const string ImageCoverTag = "ImageCover";
        public const string OverlayTag = "Overlay";

        /// <summary>The Light overlay tint: 5% white, written as 95% transparency.</summary>
        public const float LightOverlayTransparency = 0.95f;

        /// <summary>The Heavy overlay tint: 15% white, written as 85% transparency.</summary>
        public const float HeavyOverlayTransparency = 0.85f;

        /// <summary>
        /// Whether a persisted overlay's fill transparency classifies as the Light overlay.
        /// The boundary is inclusive at 0.9, between the two written levels.
        /// </summary>
        public static bool IsLightOverlay(float fillTransparency)
        {
            return fillTransparency >= 0.9f;
        }
    }
}
