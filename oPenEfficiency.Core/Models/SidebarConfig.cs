using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace oPenEfficiency.Models
{
    [DataContract]
    public class SidebarFeature
    {
        [DataMember]
        public string Id { get; set; }
        [DataMember]
        public string Name { get; set; }

        public override string ToString() => Name;

        public override bool Equals(object obj)
        {
            if (obj is SidebarFeature other)
                return Id == other.Id;
            return false;
        }

        public override int GetHashCode() => Id?.GetHashCode() ?? 0;
    }

    [DataContract]
    public class SidebarSection
    {
        [DataMember]
        public string Name { get; set; }
        [DataMember]
        public bool IsVisible { get; set; } = true;
        [DataMember]
        public string Color { get; set; }
        [DataMember]
        public List<SidebarFeature> Features { get; set; } = new List<SidebarFeature>();
    }

    [DataContract]
    public class SidebarConfig
    {
        [DataMember]
        public List<SidebarSection> Sections { get; set; } = new List<SidebarSection>();
        [DataMember]
        public List<string> LibraryPaths { get; set; } = new List<string>();
        [DataMember]
        public double IconSize { get; set; } = 12;
        [DataMember]
        public string IconColorMode { get; set; } = IconColorModes.PowerPoint;
        [DataMember]
        public bool ShowTooltips { get; set; } = true;
        [DataMember]
        public string BackgroundColor { get; set; } = "Transparent";
        // Legacy invariant: DataContractJsonSerializer leaves this null when the member is absent; Normalize then infers the mode from BackgroundColor.
        [DataMember(EmitDefaultValue = false)]
        public string BackgroundMode { get; set; } = SidebarBackgroundModes.Theme;
        [DataMember]
        public string SectionFontColor { get; set; } = "#9CA3AF";
        [DataMember(EmitDefaultValue = false)]
        public string SectionFontColorMode { get; set; } = SidebarSectionFontColorModes.Theme;
        [DataMember]
        public double SectionFontSize { get; set; } = 9;
        [DataMember]
        public string ButtonBackgroundColor { get; set; } = "Transparent";
        [DataMember]
        public string StickyNoteStyle { get; set; } = "Theme";
        [DataMember]
        public string StickyNoteDefaultText { get; set; } = "";
        [DataMember]
        public string StickyNoteFont { get; set; } = "";
        [DataMember]
        public string MasterObjectMode { get; set; } = "First selected";
        
        [DataMember]
        public string AppTheme { get; set; } = "Dark";

        [DataMember]
        public string AppFontFamily { get; set; } = "Segoe UI";
        
        // Illustrative Sticker Settings
        [DataMember]
        public string IllustrativeStickerTexts { get; set; } = "Work in progress, Updated, For discussion, Backup, Confidential, Strictly confidential";
        [DataMember]
        public string IllustrativeStickerColor { get; set; } = "#FF0000";
        [DataMember]
        public double IllustrativeStickerLineSize { get; set; } = 3.0; // Point size for line thickness (approx 3 pt ~ 0.1cm standard)
        [DataMember]
        public double IllustrativeStickerFontSize { get; set; } = 14;
        [DataMember]
        public string IllustrativeStickerFont { get; set; } = "";
        [DataMember]
        public bool IllustrativeStickerFontBold { get; set; } = true;
        [DataMember]
        public bool IllustrativeStickerFontItalic { get; set; } = false;
        [DataMember]
        public string IllustrativeStickerPosition { get; set; } = "TopRight";
        [DataMember]
        public bool IllustrativeStickerUseThemeColor { get; set; } = false;
        [DataMember]
        public bool IllustrativeStickerIncludeDate { get; set; } = false;
        [DataMember]
        public bool IllustrativeStickerIncludeAuthor { get; set; } = false;
        [DataMember]
        public double SpacingIncrement { get; set; } = 10.0;
        [DataMember]
        public string StickyNoteBackgroundColor { get; set; } = "#FFF494"; // Legacy active-style compatibility
        [DataMember]
        public string StickyNoteFontColor { get; set; } = "#323232"; // Legacy active-style compatibility
        [DataMember(EmitDefaultValue = false)]
        public string StickyNoteThemeBackgroundColor { get; set; } = "#FFF494";
        [DataMember(EmitDefaultValue = false)]
        public string StickyNoteThemeFontColor { get; set; } = "#323232";
        [DataMember(EmitDefaultValue = false)]
        public string StickyNoteConsultingBackgroundColor { get; set; } = "#FFFF99";
        [DataMember(EmitDefaultValue = false)]
        public string StickyNoteConsultingFontColor { get; set; } = "#333333";
        [DataMember(EmitDefaultValue = false)]
        public string StickyNoteGreenBBackgroundColor { get; set; } = "#FFFF00";
        [DataMember(EmitDefaultValue = false)]
        public string StickyNoteGreenBFontColor { get; set; } = "#575757";
        
        // Style Check Settings
        [DataMember]
        public string StyleCheckCorrectColors { get; set; } = ""; // Comma separated hex codes (e.g. "#FF0000, #00FF00")
        [DataMember]
        public string StyleCheckCorrectFonts { get; set; } = ""; // Comma separated font names (e.g. "Arial, Calibri")
        [DataMember]
        public string StyleCheckCorrectFontSizes { get; set; } = "10,12,14,16,18,20,24,28,32"; // Comma separated font sizes
        [DataMember]
        public string StyleCheckAllowedBullets { get; set; } = ""; // Empty means native allowed. Or specific bullet character.
        [DataMember]
        public string StyleCheckProtectionArea { get; set; } = ""; // E.g., "0,0,100,50" (Left,Top,Width,Height)
        [DataMember]
        public string StyleCheckLineSpacing { get; set; } = "";
        [DataMember]
        public string StyleCheckSpaceBefore { get; set; } = "";
        [DataMember]
        public string StyleCheckSpaceAfter { get; set; } = "";


        [DataMember(EmitDefaultValue = false)]
        public string DeepLApiKey { get; set; } = "";
        [DataMember]
        public string TranslationTargetLanguage { get; set; } = "EN-US";
        [DataMember]
        public string SelectedTranslationTool { get; set; } = "DeepL"; // DeepL, OpenAI, Claude, LibreTranslate
        [DataMember(EmitDefaultValue = false)]
        public string OpenAIApiKey { get; set; } = "";
        [DataMember]
        public string OpenAIModel { get; set; } = "gpt-4o";
        [DataMember(EmitDefaultValue = false)]
        public string ClaudeApiKey { get; set; } = "";
        [DataMember]
        public string ClaudeModel { get; set; } = "claude-3-5-sonnet-20240620";
        [DataMember]
        public string LibreTranslateBaseUrl { get; set; } = "https://libretranslate.com/";
        [DataMember(EmitDefaultValue = false)]
        public string LibreTranslateApiKey { get; set; } = "";

        // Quick Context Menu (opt-in; stable catalog IDs only).
        [DataMember]
        public bool QuickContextMenuEnabled { get; set; } = false;
        [DataMember]
        public List<string> QuickContextMenuFeatureIds { get; set; } = new List<string>();

        // Performance Mode persists only user-selected stable COM add-in identities.
        [DataMember]
        public List<string> PerformanceModeTargetIds { get; set; } = new List<string>();
        [DataMember]
        public List<string> RecentShapeFillColors { get; set; } = new List<string>();
        [DataMember]
        public List<string> RecentCountryFlagIds { get; set; } = new List<string>();
    }

    public static class QuickContextMenuSettings
    {
        public const int MinimumItems = 1;
        public const int MaximumItems = 12;

        public static readonly string[] DefaultFeatureIds =
        {
            "BtnAlignLeft", "BtnAlignTop", "BtnMatchWidth", "BtnMatchHeight",
            "BtnCopyXY", "BtnPasteXY"
        };

        public static List<string> Normalize(IEnumerable<string> ids, ISet<string> knownIds)
        {
            var result = new List<string>();
            if (ids != null)
            {
                foreach (string id in ids)
                {
                    if (string.IsNullOrWhiteSpace(id) ||
                        (knownIds != null && !knownIds.Contains(id)) ||
                        result.Contains(id))
                        continue;
                    result.Add(id);
                    if (result.Count == MaximumItems) break;
                }
            }
            return result;
        }
    }

    public static class ShapeFillPalette
    {
        public static readonly float[] Tints = { 0f, 0.8f, 0.6f, 0.4f, -0.25f, -0.5f };
        public static readonly int[] StandardRgb = { 0xC00000, 0xFF0000, 0xFFC000, 0xFFFF00, 0x92D050, 0x00B050, 0x00B0F0, 0x0070C0, 0x002060, 0x7030A0 };

        public static int ApplyTint(int rgb, float tint)
        {
            int r = (rgb >> 16) & 0xFF, g = (rgb >> 8) & 0xFF, b = rgb & 0xFF;
            System.Func<int, int> channel = value => tint >= 0
                ? (int)System.Math.Round(value + (255 - value) * tint)
                : (int)System.Math.Round(value * (1 + tint));
            return (Clamp(channel(r)) << 16) | (Clamp(channel(g)) << 8) | Clamp(channel(b));
        }

        public static bool TryParseHex(string value, out int rgb)
        {
            rgb = 0;
            if (string.IsNullOrWhiteSpace(value)) return false;
            return int.TryParse(value.Trim().TrimStart('#'), System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out rgb) && rgb >= 0 && rgb <= 0xFFFFFF;
        }

        public static string ToHex(int rgb) => $"#{rgb & 0xFFFFFF:X6}";
        public static string TintLabel(float tint) => tint == 0 ? "Base" : tint > 0 ? $"Lighter {tint:0.##}" : $"Darker {-tint:0.##}";
        private static int Clamp(int value) => System.Math.Max(0, System.Math.Min(255, value));
    }

    public static class ExternalAddInIdentity
    {
        public static bool IsOpenEfficiency(string progId, string description)
        {
            return ContainsProductName(progId) || ContainsProductName(description);
        }

        public static IEnumerable<string> Normalize(IEnumerable<string> ids)
        {
            return (ids ?? Enumerable.Empty<string>())
                .Where(id => !string.IsNullOrWhiteSpace(id) && !IsOpenEfficiency(id, null))
                .Select(id => id.Trim())
                .Distinct(System.StringComparer.OrdinalIgnoreCase)
                .OrderBy(id => id, System.StringComparer.OrdinalIgnoreCase);
        }

        private static bool ContainsProductName(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            string compact = new string(value.Where(char.IsLetterOrDigit).ToArray());
            return compact.IndexOf("openefficiency", System.StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }

    public static class SidebarBackgroundModes
    {
        public const string Theme = "Theme";
        public const string Custom = "Custom";

        public static bool UsesTheme(SidebarConfig config)
        {
            if (config == null) return true;
            if (string.Equals(config.BackgroundMode, Theme, System.StringComparison.OrdinalIgnoreCase))
                return true;
            if (string.Equals(config.BackgroundMode, Custom, System.StringComparison.OrdinalIgnoreCase))
                return false;
            return IsLegacyThemeColor(config.BackgroundColor);
        }

        public static bool Normalize(SidebarConfig config)
        {
            if (config == null) return false;

            string normalized = UsesTheme(config) ? Theme : Custom;
            bool changed = !string.Equals(config.BackgroundMode, normalized, System.StringComparison.Ordinal);
            config.BackgroundMode = normalized;

            if (normalized == Theme &&
                !string.Equals(config.BackgroundColor, "Transparent", System.StringComparison.Ordinal))
            {
                config.BackgroundColor = "Transparent";
                changed = true;
            }

            return changed;
        }

        private static bool IsLegacyThemeColor(string color)
        {
            return string.IsNullOrWhiteSpace(color) ||
                string.Equals(color, "Transparent", System.StringComparison.OrdinalIgnoreCase) ||
                string.Equals(color, "#252525", System.StringComparison.OrdinalIgnoreCase);
        }
    }

    public static class SidebarSectionFontColorModes
    {
        public const string Theme = "Theme";
        public const string Custom = "Custom";

        public static bool UsesTheme(SidebarConfig config)
        {
            if (config == null) return true;
            if (string.Equals(config.SectionFontColorMode, Theme, System.StringComparison.OrdinalIgnoreCase))
                return true;
            if (string.Equals(config.SectionFontColorMode, Custom, System.StringComparison.OrdinalIgnoreCase))
                return false;

            return string.IsNullOrWhiteSpace(config.SectionFontColor) ||
                   string.Equals(config.SectionFontColor, "#9CA3AF", System.StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(config.SectionFontColor, "#888888", System.StringComparison.OrdinalIgnoreCase);
        }

        public static bool Normalize(SidebarConfig config)
        {
            if (config == null) return false;

            string normalized = UsesTheme(config) ? Theme : Custom;
            bool changed = !string.Equals(config.SectionFontColorMode, normalized, System.StringComparison.Ordinal);
            config.SectionFontColorMode = normalized;
            return changed;
        }
    }
}
