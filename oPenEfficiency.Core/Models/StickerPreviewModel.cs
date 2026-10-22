using System;
using System.Linq;

namespace oPenEfficiency.Models
{
    public sealed class StickyNotePreviewState
    {
        public string Style { get; set; }
        public string Text { get; set; }
        public string FontFamily { get; set; }
        public string BackgroundColor { get; set; }
        public string FontColor { get; set; }
    }

    public sealed class StatusStickerPreviewState
    {
        public string Text { get; set; }
        public string FontFamily { get; set; }
        public string Color { get; set; }
        public double FontSize { get; set; }
        public double LineSize { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool UseThemeColor { get; set; }
        public string PositionLabel { get; set; }
    }

    /// <summary>
    /// Resolves persisted sticker settings into host-neutral visual state. PowerPoint rendering and
    /// Settings previews can share fallbacks without requiring Office interop in the view model.
    /// </summary>
    public static class StickerPreviewModel
    {
        public const string DefaultStickyText = "Sticky note";
        public const string DefaultStatusText = "Work in progress";

        public static StickyNotePreviewState CreateStickyNote(
            SidebarConfig config,
            string themeFont,
            string styleOverride = null)
        {
            string style = NormalizeStickyStyle(styleOverride ?? config?.StickyNoteStyle);
            return new StickyNotePreviewState
            {
                Style = style,
                Text = string.IsNullOrWhiteSpace(config?.StickyNoteDefaultText)
                    ? DefaultStickyText
                    : config.StickyNoteDefaultText.Trim(),
                FontFamily = ResolveFont(config?.StickyNoteFont, themeFont),
                BackgroundColor = ResolveStickyColor(config, style, true),
                FontColor = ResolveStickyColor(config, style, false)
            };
        }

        public static void SetStickyNoteColors(
            SidebarConfig config,
            string style,
            string backgroundColor,
            string fontColor)
        {
            if (config == null) return;
            string normalizedStyle = NormalizeStickyStyle(style);
            string resolvedBackground = backgroundColor ?? GetDefaultStickyColor(normalizedStyle, true);
            string resolvedFont = fontColor ?? GetDefaultStickyColor(normalizedStyle, false);

            switch (normalizedStyle)
            {
                case "Consulting":
                    config.StickyNoteConsultingBackgroundColor = resolvedBackground;
                    config.StickyNoteConsultingFontColor = resolvedFont;
                    break;
                case "GreenB":
                    config.StickyNoteGreenBBackgroundColor = resolvedBackground;
                    config.StickyNoteGreenBFontColor = resolvedFont;
                    break;
                default:
                    config.StickyNoteThemeBackgroundColor = resolvedBackground;
                    config.StickyNoteThemeFontColor = resolvedFont;
                    break;
            }

            // Keep older versions and existing export consumers aligned with the active style.
            config.StickyNoteBackgroundColor = resolvedBackground;
            config.StickyNoteFontColor = resolvedFont;
        }

        public static bool NormalizeStickyNoteColors(SidebarConfig config)
        {
            if (config == null) return false;

            bool hadStyleColors = config.StickyNoteThemeBackgroundColor != null ||
                config.StickyNoteThemeFontColor != null ||
                config.StickyNoteConsultingBackgroundColor != null ||
                config.StickyNoteConsultingFontColor != null ||
                config.StickyNoteGreenBBackgroundColor != null ||
                config.StickyNoteGreenBFontColor != null;
            bool changed = false;

            if (config.StickyNoteThemeBackgroundColor == null) { config.StickyNoteThemeBackgroundColor = "#FFF494"; changed = true; }
            if (config.StickyNoteThemeFontColor == null) { config.StickyNoteThemeFontColor = "#323232"; changed = true; }
            if (config.StickyNoteConsultingBackgroundColor == null) { config.StickyNoteConsultingBackgroundColor = "#FFFF99"; changed = true; }
            if (config.StickyNoteConsultingFontColor == null) { config.StickyNoteConsultingFontColor = "#333333"; changed = true; }
            if (config.StickyNoteGreenBBackgroundColor == null) { config.StickyNoteGreenBBackgroundColor = "#FFFF00"; changed = true; }
            if (config.StickyNoteGreenBFontColor == null) { config.StickyNoteGreenBFontColor = "#575757"; changed = true; }

            if (!hadStyleColors)
            {
                string legacyBackground = NormalizeHexColor(config.StickyNoteBackgroundColor, "#FFF494");
                string legacyFont = NormalizeHexColor(config.StickyNoteFontColor, "#323232");
                if (!string.Equals(legacyBackground, "#FFF494", StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(legacyFont, "#323232", StringComparison.OrdinalIgnoreCase))
                {
                    SetStickyNoteColors(config, config.StickyNoteStyle, legacyBackground, legacyFont);
                    changed = true;
                }
            }

            var active = CreateStickyNote(config, null, config.StickyNoteStyle);
            if (!string.Equals(config.StickyNoteBackgroundColor, active.BackgroundColor, StringComparison.Ordinal) ||
                !string.Equals(config.StickyNoteFontColor, active.FontColor, StringComparison.Ordinal))
            {
                config.StickyNoteBackgroundColor = active.BackgroundColor;
                config.StickyNoteFontColor = active.FontColor;
                changed = true;
            }
            return changed;
        }

        public static StatusStickerPreviewState CreateStatusSticker(
            SidebarConfig config,
            string themeFont,
            DateTime? today = null,
            string author = null)
        {
            string label = GetFirstStatusLabel(config?.IllustrativeStickerTexts);
            return new StatusStickerPreviewState
            {
                Text = BuildStatusDisplayText(
                    label,
                    config?.IllustrativeStickerIncludeDate ?? false,
                    config?.IllustrativeStickerIncludeAuthor ?? false,
                    today ?? DateTime.Today,
                    author ?? Environment.UserName),
                FontFamily = ResolveFont(config?.IllustrativeStickerFont, themeFont),
                Color = NormalizeHexColor(config?.IllustrativeStickerColor, "#FF0000"),
                FontSize = NormalizePositive(config?.IllustrativeStickerFontSize, 14d),
                LineSize = NormalizePositive(config?.IllustrativeStickerLineSize, 3d),
                Bold = config?.IllustrativeStickerFontBold ?? true,
                Italic = config?.IllustrativeStickerFontItalic ?? false,
                UseThemeColor = config?.IllustrativeStickerUseThemeColor ?? false,
                PositionLabel = HumanizePosition(config?.IllustrativeStickerPosition)
            };
        }

        public static string GetFirstStatusLabel(string labels)
        {
            return (labels ?? string.Empty)
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(label => label.Trim())
                .FirstOrDefault(label => label.Length > 0) ?? DefaultStatusText;
        }

        public static string BuildStatusDisplayText(
            string label,
            bool includeDate,
            bool includeAuthor,
            DateTime today,
            string author)
        {
            var parts = new[]
            {
                string.IsNullOrWhiteSpace(label) ? DefaultStatusText : label.Trim(),
                includeDate ? today.ToString("yyyy-MM-dd") : null,
                includeAuthor && !string.IsNullOrWhiteSpace(author) ? author.Trim() : null
            };
            return string.Join(" · ", parts.Where(part => !string.IsNullOrWhiteSpace(part)));
        }

        public static double NormalizePositive(double? value, double fallback)
        {
            return value.HasValue && value.Value > 0 && !double.IsNaN(value.Value) && !double.IsInfinity(value.Value)
                ? value.Value
                : fallback;
        }

        public static string NormalizeHexColor(string value, string fallback)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                string candidate = value.Trim();
                string digits = candidate.StartsWith("#", StringComparison.Ordinal) ? candidate.Substring(1) : candidate;
                if (digits.Length == 6 && digits.All(Uri.IsHexDigit)) return "#" + digits.ToUpperInvariant();
            }
            return fallback;
        }

        public static string GetDefaultStickyColor(string style, bool background)
        {
            switch (NormalizeStickyStyle(style))
            {
                case "Consulting": return background ? "#FFFF99" : "#333333";
                case "GreenB": return background ? "#FFFF00" : "#575757";
                default: return background ? "#FFF494" : "#323232";
            }
        }

        private static string ResolveStickyColor(SidebarConfig config, string style, bool background)
        {
            string configured = null;
            switch (NormalizeStickyStyle(style))
            {
                case "Consulting":
                    configured = background
                        ? config?.StickyNoteConsultingBackgroundColor
                        : config?.StickyNoteConsultingFontColor;
                    break;
                case "GreenB":
                    configured = background
                        ? config?.StickyNoteGreenBBackgroundColor
                        : config?.StickyNoteGreenBFontColor;
                    break;
                default:
                    configured = background
                        ? config?.StickyNoteThemeBackgroundColor
                        : config?.StickyNoteThemeFontColor;
                    break;
            }
            if (!string.IsNullOrWhiteSpace(configured))
                return NormalizeHexColor(configured, GetDefaultStickyColor(style, background));

            // Legacy configurations only contain one color pair. Preserve an intentional custom
            // value for the formerly active style, but restore historical defaults otherwise.
            if (config != null &&
                string.Equals(NormalizeStickyStyle(config.StickyNoteStyle), NormalizeStickyStyle(style), StringComparison.Ordinal))
            {
                string legacy = background ? config.StickyNoteBackgroundColor : config.StickyNoteFontColor;
                string legacyDefault = background ? "#FFF494" : "#323232";
                string normalizedLegacy = NormalizeHexColor(legacy, legacyDefault);
                if (!string.Equals(normalizedLegacy, legacyDefault, StringComparison.OrdinalIgnoreCase))
                    return normalizedLegacy;
            }
            return GetDefaultStickyColor(style, background);
        }

        private static string NormalizeStickyStyle(string value)
        {
            if (string.Equals(value, "Consulting", StringComparison.OrdinalIgnoreCase)) return "Consulting";
            if (string.Equals(value, "GreenB", StringComparison.OrdinalIgnoreCase)) return "GreenB";
            return "Theme";
        }

        private static string ResolveFont(string configuredFont, string themeFont)
        {
            if (!string.IsNullOrWhiteSpace(configuredFont) &&
                !string.Equals(configuredFont, "Theme Font", StringComparison.OrdinalIgnoreCase))
                return configuredFont.Trim();
            return string.IsNullOrWhiteSpace(themeFont) ? "Segoe UI" : themeFont.Trim();
        }

        private static string HumanizePosition(string value)
        {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "topleft": return "Top left";
                case "bottomleft": return "Bottom left";
                case "bottomright": return "Bottom right";
                default: return "Top right";
            }
        }

    }
}
