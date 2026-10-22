using System;

namespace oPenEfficiency.Models
{
    public static class IconColorModes
    {
        public const string PowerPoint = "PowerPoint";
        public const string FeatureColors = "Feature colors";

        public static string Normalize(string value)
        {
            if (string.Equals(value, FeatureColors, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "FeatureColors", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "Feature Colors", StringComparison.OrdinalIgnoreCase))
            {
                return FeatureColors;
            }

            return PowerPoint;
        }
    }

}

namespace oPenEfficiency.UI
{
    using oPenEfficiency.Models;

    public static class FeatureIconSizing
    {
        public const double OpticalExpansion = 4d;

        public static double ResolveGlyphSize(double configuredIconSize)
        {
            return configuredIconSize + OpticalExpansion;
        }
    }

    public enum FeatureIconVisualState
    {
        Normal,
        Hover,
        Pressed,
        Selected,
        Disabled
    }

    public enum FeatureIconAccentRole
    {
        None,
        ActionBlue,
        CreateGreen,
        SuccessGreen,
        AttentionYellow,
        DestructiveRed,
        ContentColor
    }

    public sealed class FeatureIconColorDecision
    {
        private FeatureIconColorDecision(string resourceKey, string storedFeatureColor)
        {
            ResourceKey = resourceKey;
            StoredFeatureColor = storedFeatureColor;
        }

        public string ResourceKey { get; }
        public string StoredFeatureColor { get; }
        public bool UsesStoredFeatureColor => ResourceKey == null;

        public static FeatureIconColorDecision ThemeResource(string resourceKey)
        {
            return new FeatureIconColorDecision(resourceKey, null);
        }

        public static FeatureIconColorDecision FeatureColor(string storedFeatureColor)
        {
            return new FeatureIconColorDecision(null, storedFeatureColor);
        }
    }

    /// <summary>
    /// Host-neutral policy for all command-icon consumers. WPF surfaces translate the returned
    /// semantic resource key into a dynamic theme brush without duplicating mode decisions.
    /// Supports two-layer model: base layer uses Office-neutral foreground; accent layer uses
    /// a semantic accent resource.
    /// </summary>
    public static class FeatureIconColorPolicy
    {
        // Base layer resource keys
        public const string NormalResourceKey = "OfficeIconForegroundBrush";
        public const string DisabledResourceKey = "OfficeIconDisabledBrush";

        // Accent layer semantic resource keys
        public const string AccentBlueResourceKey = "AccentBlueBrush";
        public const string AccentGreenResourceKey = "AccentGreenBrush";
        public const string SuccessGreenResourceKey = "SuccessGreenBrush";
        public const string AccentYellowResourceKey = "AccentYellowBrush";
        public const string AccentRedResourceKey = "AccentRedBrush";

        public static FeatureIconColorDecision ResolveForLayer(
            string mode,
            FeatureIconVisualState state,
            string storedFeatureColor,
            bool isAccent,
            FeatureIconAccentRole accentRole)
        {
            if (state == FeatureIconVisualState.Disabled)
            {
                return FeatureIconColorDecision.ThemeResource(DisabledResourceKey);
            }

            if (IconColorModes.Normalize(mode) == IconColorModes.FeatureColors &&
                !string.IsNullOrWhiteSpace(storedFeatureColor))
            {
                return FeatureIconColorDecision.FeatureColor(storedFeatureColor);
            }

            if (!isAccent)
            {
                return FeatureIconColorDecision.ThemeResource(NormalResourceKey);
            }

            switch (accentRole)
            {
                case FeatureIconAccentRole.CreateGreen:
                    return FeatureIconColorDecision.ThemeResource(AccentGreenResourceKey);
                case FeatureIconAccentRole.SuccessGreen:
                    return FeatureIconColorDecision.ThemeResource(SuccessGreenResourceKey);
                case FeatureIconAccentRole.AttentionYellow:
                    return FeatureIconColorDecision.ThemeResource(AccentYellowResourceKey);
                case FeatureIconAccentRole.DestructiveRed:
                    return FeatureIconColorDecision.ThemeResource(AccentRedResourceKey);
                case FeatureIconAccentRole.ActionBlue:
                case FeatureIconAccentRole.ContentColor:
                case FeatureIconAccentRole.None:
                default:
                    return FeatureIconColorDecision.ThemeResource(AccentBlueResourceKey);
            }
        }

    }
}
