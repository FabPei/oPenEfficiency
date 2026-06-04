using System.Collections.Generic;
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
        [DataMember]
        public string Keywords { get; set; }

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
        public bool IsDarkMode { get; set; } = true;
        [DataMember]
        public double IconSize { get; set; } = 12;
        [DataMember]
        public bool ShowTooltips { get; set; } = true;
        [DataMember]
        public string BackgroundColor { get; set; } = "#252525";
        [DataMember]
        public string SectionFontColor { get; set; } = "#9CA3AF";
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
        public string MasterObjectMode { get; set; } = "First Selected";
        
        [DataMember]
        public string AppTheme { get; set; } = "Dark";

        [DataMember]
        public string AppFontFamily { get; set; } = "Segoe UI";
        
        // Illustrative Sticker Settings
        [DataMember]
        public string IllustrativeStickerTexts { get; set; } = "Work in progress, Updated, For discussion, Backup, Confidential, Strictly confidential";
        [DataMember]
        public bool EnableSwapPositionsContextMenu { get; set; } = true;
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
        public double SpacingIncrement { get; set; } = 10.0;
        [DataMember]
        public string StickyNoteBackgroundColor { get; set; } = "#FFF494"; // Default yellow
        [DataMember]
        public string StickyNoteFontColor { get; set; } = "#323232"; // Default dark grey
        
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


        [DataMember]
        public string DeepLApiKey { get; set; } = "";
        [DataMember]
        public string TranslationTargetLanguage { get; set; } = "EN-US";
        [DataMember]
        public string SelectedTranslationTool { get; set; } = "DeepL"; // DeepL, OpenAI, Claude, LibreTranslate
        [DataMember]
        public string OpenAIApiKey { get; set; } = "";
        [DataMember]
        public string OpenAIModel { get; set; } = "gpt-4o";
        [DataMember]
        public string ClaudeApiKey { get; set; } = "";
        [DataMember]
        public string ClaudeModel { get; set; } = "claude-3-5-sonnet-20240620";
        [DataMember]
        public string LibreTranslateBaseUrl { get; set; } = "https://libretranslate.com/";
        [DataMember]
        public string LibreTranslateApiKey { get; set; } = "";
    }
}
