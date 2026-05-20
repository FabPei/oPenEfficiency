using System;
using System.Collections.Generic;
using System.Linq;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI
{
    public struct FeatureDisplayInfo
    {
        public string Tooltip;
        public string IconData;
        public string Color;
        public string Description;
        public string DetailedHelpText;
        public string HelpImagePath;
        public int MinSelection { get; set; }
        public int MaxSelection { get; set; }
        public PowerPoint.PpSelectionType RequiredType { get; set; }
        public bool RequiresTable { get; set; }
        public bool RequiresHiddenShapes { get; set; }
        public bool IsToggle { get; set; }
    }

    public static class FeatureLibrary
    {
        public static List<SidebarFeature> AllFeatures = new List<SidebarFeature>
        {
            new SidebarFeature { Id = "BtnSearchBar", Name = "Search Command Palette" },
            new SidebarFeature { Id = "BtnSizePanel", Name = "Shape Size Panel" },
            new SidebarFeature { Id = "BtnMatchSize", Name = "Match Size" },
            new SidebarFeature { Id = "BtnMatchWidth", Name = "Match Width" },
            new SidebarFeature { Id = "BtnMatchHeight", Name = "Match Height" },
            new SidebarFeature { Id = "BtnLockDimensions", Name = "Lock Aspect Ratio" },
            new SidebarFeature { Id = "BtnShapeLock", Name = "Shape Locking" },
            new SidebarFeature { Id = "BtnSlideSize", Name = "Slide Size Analysis" },
            new SidebarFeature { Id = "BtnAlignLeft", Name = "Align Left" },
            new SidebarFeature { Id = "BtnAlignCenter", Name = "Align Center" },
            new SidebarFeature { Id = "BtnAlignRight", Name = "Align Right" },
            new SidebarFeature { Id = "BtnAlignTop", Name = "Align Top" },
            new SidebarFeature { Id = "BtnAlignCenterVertical", Name = "Align Middle" },
            new SidebarFeature { Id = "BtnAlignBottom", Name = "Align Bottom" },
            new SidebarFeature { Id = "BtnDiagonalAlign", Name = "Diagonal Alignment" },
            new SidebarFeature { Id = "BtnDistributeHorizontal", Name = "Distribute Horizontal" },
            new SidebarFeature { Id = "BtnDistributeVertical", Name = "Distribute Vertical" },
            new SidebarFeature { Id = "BtnRemoveHorizontalGaps", Name = "Remove Horiz Gaps" },
            new SidebarFeature { Id = "BtnRemoveVerticalGaps", Name = "Remove Vert Gaps" },
            new SidebarFeature { Id = "BtnIncreaseHorizontalSpacing", Name = "Incr Horiz Spacing" },
            new SidebarFeature { Id = "BtnDecreaseHorizontalSpacing", Name = "Decr Horiz Spacing" },
            new SidebarFeature { Id = "BtnIncreaseVerticalSpacing", Name = "Incr Vert Spacing" },
            new SidebarFeature { Id = "BtnDecreaseVerticalSpacing", Name = "Decr Vert Spacing" },
            new SidebarFeature { Id = "BtnAdjustSpacing", Name = "Advanced Spacing" },
            new SidebarFeature { Id = "BtnStretchLeft", Name = "Stretch Left" },
            new SidebarFeature { Id = "BtnStretchRight", Name = "Stretch Right" },
            new SidebarFeature { Id = "BtnStretchTop", Name = "Stretch Top" },
            new SidebarFeature { Id = "BtnStretchBottom", Name = "Stretch Bottom" },
            new SidebarFeature { Id = "BtnStretchLeftEdge", Name = "Stretch to Left Edge" },
            new SidebarFeature { Id = "BtnStretchRightEdge", Name = "Stretch to Right Edge" },
            new SidebarFeature { Id = "BtnStretchTopEdge", Name = "Stretch to Top Edge" },
            new SidebarFeature { Id = "BtnStretchBottomEdge", Name = "Stretch to Bottom Edge" },
            new SidebarFeature { Id = "BtnDockLeft", Name = "Dock Left" },
            new SidebarFeature { Id = "BtnDockRight", Name = "Dock Right" },
            new SidebarFeature { Id = "BtnDockTop", Name = "Dock Top" },
            new SidebarFeature { Id = "BtnDockBottom", Name = "Dock Bottom" },
            new SidebarFeature { Id = "BtnCloneLeft", Name = "Clone Left" },
            new SidebarFeature { Id = "BtnCloneRight", Name = "Clone Right" },
            new SidebarFeature { Id = "BtnCloneTop", Name = "Clone Top" },
            new SidebarFeature { Id = "BtnCloneBottom", Name = "Clone Bottom" },
            new SidebarFeature { Id = "BtnSwapPositions", Name = "Swap Positions" },
            new SidebarFeature { Id = "BtnCopyXY", Name = "Copy Coordinates" },
            new SidebarFeature { Id = "BtnPasteXY", Name = "Paste Coordinates" },
            new SidebarFeature { Id = "BtnSelectSameType", Name = "Select Same Type" },
            new SidebarFeature { Id = "BtnSyncObjects", Name = "Sync Across Slides" },
            new SidebarFeature { Id = "BtnArrangeCircle", Name = "Arrange in Circle" },
            new SidebarFeature { Id = "BtnArrangePro", Name = "Arrange Pro" },
            new SidebarFeature { Id = "BtnRectifyRotation", Name = "Rectify Rotation" },
            new SidebarFeature { Id = "BtnAlignShapeAdjustments", Name = "Align Handles" },
            new SidebarFeature { Id = "BtnHideSelected", Name = "Hide Selected" },
            new SidebarFeature { Id = "BtnShowHidden", Name = "Show All Hidden" },
            new SidebarFeature { Id = "BtnHideMasterObjects", Name = "Hide Master Background" },
            new SidebarFeature { Id = "BtnCreateMotionPath", Name = "Create Motion Path" },
            new SidebarFeature { Id = "BtnBringToFront", Name = "Bring to Front" },
            new SidebarFeature { Id = "BtnBringForward", Name = "Bring Forward" },
            new SidebarFeature { Id = "BtnSendBackward", Name = "Send Backward" },
            new SidebarFeature { Id = "BtnSendToBack", Name = "Send to Back" },
            new SidebarFeature { Id = "BtnSnapShape", Name = "Snap Shape to Object" },
            new SidebarFeature { Id = "BtnSnapToGrid", Name = "Snap to Grid" },
            new SidebarFeature { Id = "BtnSnapToObjects", Name = "Snap to Objects" },
            new SidebarFeature { Id = "BtnFontPanel", Name = "Quick Font Panel" },
            new SidebarFeature { Id = "BtnApplyTextTool", Name = "Apply Text Formatting" },
            new SidebarFeature { Id = "BtnFitFormToText", Name = "Fit Shape to Text" },
            new SidebarFeature { Id = "BtnTranslate", Name = "Translate Text" },
            new SidebarFeature { Id = "BtnReplaceFont", Name = "Bulk Format Replacer" },
            new SidebarFeature { Id = "BtnParagraphDialog", Name = "Paragraph Dialog" },
            new SidebarFeature { Id = "BtnSpecialChars", Name = "Insert Special Chars" },
            new SidebarFeature { Id = "BtnFormatNumbers", Name = "Format Numbers" },
            new SidebarFeature { Id = "BtnAlignText", Name = "Align Text in Box" },
            new SidebarFeature { Id = "BtnTextDirection", Name = "Change Text Direction" },
            new SidebarFeature { Id = "BtnChangeCase", Name = "Change Case" },
            new SidebarFeature { Id = "BtnCharacterSpacing", Name = "Character Spacing" },
            new SidebarFeature { Id = "BtnSpellCheck", Name = "Spell Check Language" },
            new SidebarFeature { Id = "BtnDeleteText", Name = "Clear Shape Text" },
            new SidebarFeature { Id = "BtnSwapTextFormatted", Name = "Swap Text (Styles)" },
            new SidebarFeature { Id = "BtnSwapTextPlain", Name = "Swap Text (Plain)" },
            new SidebarFeature { Id = "BtnPickColor", Name = "Color Picker" },
            new SidebarFeature { Id = "BtnPickFillColor", Name = "Pick Fill Color" },
            new SidebarFeature { Id = "BtnPickLineColor", Name = "Pick Line Color" },
            new SidebarFeature { Id = "BtnPickTextColor", Name = "Pick Text Color" },
            new SidebarFeature { Id = "BtnThemeColor", Name = "Absolute Theme Colors" },
            new SidebarFeature { Id = "BtnTransparentColor", Name = "Set Transparency" },
            new SidebarFeature { Id = "BtnFormatShapeDialog", Name = "Format Shape Dialog" },
            new SidebarFeature { Id = "BtnOptimizeFreeForm", Name = "Optimize Freeform" },
            new SidebarFeature { Id = "BtnStyleCheck", Name = "Presentation Style Check" },
            new SidebarFeature { Id = "BtnCleaner", Name = "Deep Presentation Cleaner" },
            new SidebarFeature { Id = "BtnFormatBold", Name = "Bold" },
            new SidebarFeature { Id = "BtnFormatItalic", Name = "Italic" },
            new SidebarFeature { Id = "BtnFormatUnderline", Name = "Underline" },
            new SidebarFeature { Id = "BtnFormatStrikethrough", Name = "Strikethrough" },
            new SidebarFeature { Id = "BtnFormatSuperscript", Name = "Superscript" },
            new SidebarFeature { Id = "BtnFormatSubscript", Name = "Subscript" },
            new SidebarFeature { Id = "BtnInsertColumnLeft", Name = "Insert Column Left" },
            new SidebarFeature { Id = "BtnInsertColumnRight", Name = "Insert Column Right" },
            new SidebarFeature { Id = "BtnInsertRowTop", Name = "Insert Row Top" },
            new SidebarFeature { Id = "BtnInsertRowBottom", Name = "Insert Row Bottom" },
            new SidebarFeature { Id = "BtnTableTranspose", Name = "Transpose Table" },
            new SidebarFeature { Id = "BtnTableSplit", Name = "Split Table" },
            new SidebarFeature { Id = "BtnConvertTableToShapes", Name = "Convert Table to Shapes" },
            new SidebarFeature { Id = "BtnTableDimensions", Name = "Table Size Manager" },
            new SidebarFeature { Id = "BtnTableFormatPainter", Name = "Table Format Painter" },
            new SidebarFeature { Id = "BtnTableSort", Name = "Table Sorting" },
            new SidebarFeature { Id = "BtnTableSum", Name = "Table Auto-Sum" },
            new SidebarFeature { Id = "BtnTableBranding", Name = "Table Branding" },
            new SidebarFeature { Id = "BtnAlignToTableCell", Name = "Align to Cell" },
            new SidebarFeature { Id = "BtnIllustrativeSticker", Name = "Status Stickers" },
            new SidebarFeature { Id = "BtnAddSticker", Name = "Illustrative Stickers" },
            new SidebarFeature { Id = "BtnHarveyBall", Name = "Harvey Ball" },
            new SidebarFeature { Id = "BtnTrafficLight", Name = "Traffic Light" },
            new SidebarFeature { Id = "BtnThermometer", Name = "Thermometer Chart" },
            new SidebarFeature { Id = "BtnStarRating", Name = "Star Rating" },
            new SidebarFeature { Id = "BtnCheckbox", Name = "Vector Checkbox" },
            new SidebarFeature { Id = "BtnNumeration", Name = "Auto Numeration" },
            new SidebarFeature { Id = "BtnProgressSeries", Name = "Progress Series" },
            new SidebarFeature { Id = "BtnRepeatShape", Name = "Repeat Shape" },
            new SidebarFeature { Id = "BtnSplitShape", Name = "Split Shape to Grid" },
            new SidebarFeature { Id = "BtnSplitByParagraphs", Name = "Split by Paragraphs" },
            new SidebarFeature { Id = "BtnAgendaWizard", Name = "Agenda Wizard" },
            new SidebarFeature { Id = "BtnSaveAgendaLayout", Name = "Save Agenda Layout" },
            new SidebarFeature { Id = "BtnStoryline", Name = "Storyline" },
            new SidebarFeature { Id = "BtnSlideGuidelines", Name = "Slide Grid Guidelines" },
            new SidebarFeature { Id = "BtnConsolidateMasters", Name = "Consolidate Masters" },
            new SidebarFeature { Id = "BtnAssetManager", Name = "Asset Library" },
            new SidebarFeature { Id = "BtnWinnerPicker", Name = "Lottery / Winner Picker" },
            new SidebarFeature { Id = "BtnUpdateExcelCharts", Name = "Update Excel Data" },
            new SidebarFeature { Id = "BtnQRCode", Name = "QR Code Generator" },
            new SidebarFeature { Id = "BtnExportWizard", Name = "Export Wizard" },
            new SidebarFeature { Id = "BtnSettings", Name = "oPenEfficiency Settings" },
            new SidebarFeature { Id = "BtnTemplateManager", Name = "Template Library" },
            new SidebarFeature { Id = "BtnStickyNote", Name = "Sticky Note" },
            new SidebarFeature { Id = "BtnSlidePaste", Name = "Smart Slide Paste" }
        };

        public static SidebarConfig GetEmergencyDefault() => GetStandardLayout();

        public static SidebarConfig GetStandardLayout()
        {
            var config = new SidebarConfig();

            var sizeSection = new SidebarSection { Name = "PROPERTIES & SIZE", Color = "#6366F1" };
            string[] sizeIds = { "BtnSearchBar", "BtnSizePanel", "BtnMatchSize", "BtnMatchWidth", "BtnMatchHeight", "BtnLockDimensions", "BtnShapeLock", "BtnSlideSize" };
            foreach (var id in sizeIds) sizeSection.Features.Add(AllFeatures.Find(f => f.Id == id));
            config.Sections.Add(sizeSection);

            var alignSection = new SidebarSection { Name = "ALIGN", Color = "#10B981" };
            string[] alignIds = { "BtnAlignLeft", "BtnAlignCenter", "BtnAlignRight", "BtnAlignTop", "BtnAlignCenterVertical", "BtnAlignBottom", "BtnDiagonalAlign", "BtnDistributeHorizontal", "BtnDistributeVertical", "BtnRemoveHorizontalGaps", "BtnRemoveVerticalGaps", "BtnIncreaseHorizontalSpacing", "BtnDecreaseHorizontalSpacing", "BtnIncreaseVerticalSpacing", "BtnDecreaseVerticalSpacing", "BtnAdjustSpacing" };
            foreach (var id in alignIds) alignSection.Features.Add(AllFeatures.Find(f => f.Id == id));
            config.Sections.Add(alignSection);

            var stretchSection = new SidebarSection { Name = "DOCK", Color = "#F59E0B" };
            string[] stretchIds = { "BtnStretchLeft", "BtnStretchRight", "BtnStretchTop", "BtnStretchBottom", "BtnStretchLeftEdge", "BtnStretchRightEdge", "BtnStretchTopEdge", "BtnStretchBottomEdge", "BtnDockLeft", "BtnDockRight", "BtnDockTop", "BtnDockBottom" };
            foreach (var id in stretchIds) stretchSection.Features.Add(AllFeatures.Find(f => f.Id == id));
            config.Sections.Add(stretchSection);

            var arrangeSection = new SidebarSection { Name = "ARRANGE & ORGANIZE", Color = "#FBBF24" };
            string[] arrangeIds = { "BtnSwapPositions", "BtnCopyXY", "BtnPasteXY", "BtnSelectSameType", "BtnSyncObjects", "BtnArrangeCircle", "BtnArrangePro", "BtnRectifyRotation", "BtnAlignShapeAdjustments", "BtnHideSelected", "BtnShowHidden", "BtnHideMasterObjects", "BtnCreateMotionPath", "BtnBringToFront", "BtnBringForward", "BtnSendBackward", "BtnSendToBack", "BtnSnapShape", "BtnSnapToGrid", "BtnSnapToObjects" };
            foreach (var id in arrangeIds) arrangeSection.Features.Add(AllFeatures.Find(f => f.Id == id));
            config.Sections.Add(arrangeSection);

            var textSection = new SidebarSection { Name = "TEXT & FONTS", Color = "#FDE047" };
            string[] textIds = { "BtnFontPanel", "BtnApplyTextTool", "BtnFitFormToText", "BtnTranslate", "BtnReplaceFont", "BtnParagraphDialog", "BtnSpecialChars", "BtnFormatNumbers", "BtnAlignText", "BtnTextDirection", "BtnChangeCase", "BtnCharacterSpacing", "BtnSpellCheck", "BtnDeleteText", "BtnSwapTextFormatted", "BtnSwapTextPlain" };
            foreach (var id in textIds) textSection.Features.Add(AllFeatures.Find(f => f.Id == id));
            config.Sections.Add(textSection);

            var formatSection = new SidebarSection { Name = "FORMAT & STYLE", Color = "#06B6D4" };
            string[] formatIds = { "BtnPickColor", "BtnPickFillColor", "BtnPickLineColor", "BtnPickTextColor", "BtnThemeColor", "BtnTransparentColor", "BtnFormatShapeDialog", "BtnOptimizeFreeForm", "BtnStyleCheck", "BtnCleaner", "BtnFormatBold", "BtnFormatItalic", "BtnFormatUnderline", "BtnFormatStrikethrough", "BtnFormatSuperscript", "BtnFormatSubscript" };
            foreach (var id in formatIds) formatSection.Features.Add(AllFeatures.Find(f => f.Id == id));
            config.Sections.Add(formatSection);

            var tablesSection = new SidebarSection { Name = "TABLE TOOLS", Color = "#3B82F6" };
            string[] tablesIds = { "BtnInsertColumnLeft", "BtnInsertColumnRight", "BtnInsertRowTop", "BtnInsertRowBottom", "BtnTableTranspose", "BtnTableSplit", "BtnConvertTableToShapes", "BtnTableDimensions", "BtnTableFormatPainter", "BtnTableSort", "BtnTableSum", "BtnTableBranding", "BtnAlignToTableCell" };
            foreach (var id in tablesIds) tablesSection.Features.Add(AllFeatures.Find(f => f.Id == id));
            config.Sections.Add(tablesSection);

            var visualsSection = new SidebarSection { Name = "VISUAL ELEMENTS", Color = "#EF4444" };
            string[] visualsIds = { "BtnIllustrativeSticker", "BtnAddSticker", "BtnHarveyBall", "BtnTrafficLight", "BtnThermometer", "BtnStarRating", "BtnCheckbox", "BtnNumeration", "BtnProgressSeries", "BtnRepeatShape", "BtnSplitShape", "BtnSplitByParagraphs" };
            foreach (var id in visualsIds) visualsSection.Features.Add(AllFeatures.Find(f => f.Id == id));
            config.Sections.Add(visualsSection);

            var storySection = new SidebarSection { Name = "STORYLINING & REVIEW", Color = "#FDE047" };
            string[] storyIds = { "BtnAgendaWizard", "BtnSaveAgendaLayout", "BtnStoryline", "BtnStyleCheck", "BtnSlideGuidelines", "BtnCleaner", "BtnFlightMode", "BtnStickyNote", "BtnSlideNotes" };
            foreach (var id in storyIds) storySection.Features.Add(AllFeatures.Find(f => f.Id == id));
            config.Sections.Add(storySection);

            var utilSection = new SidebarSection { Name = "UTILITIES & WIZARDS", Color = "#10B981" };
            string[] utilIds = { "BtnAssetManager", "BtnWinnerPicker", "BtnUpdateExcelCharts", "BtnQRCode", "BtnExportWizard", "BtnSettings", "BtnTemplateManager", "BtnStickyNote", "BtnSlidePaste" };
            foreach (var id in utilIds) utilSection.Features.Add(AllFeatures.Find(f => f.Id == id));
            config.Sections.Add(utilSection);

            return config;
        }

        public static SidebarConfig GetAllFeaturesLayout() => GetStandardLayout();

        public static SidebarConfig GetMinimalistLayout()
        {
            var config = new SidebarConfig();
            var essentialSection = new SidebarSection { Name = "ESSENTIALS", Color = "#3B82F6" };
            string[] essentialIds = { "BtnSearchBar", "BtnSizePanel", "BtnAlignLeft", "BtnAlignCenter", "BtnAlignRight", "BtnAlignTop", "BtnAlignCenterVertical", "BtnAlignBottom", "BtnFontPanel" };
            foreach (var id in essentialIds) essentialSection.Features.Add(AllFeatures.Find(f => f.Id == id));
            config.Sections.Add(essentialSection);
            return config;
        }

        public static FeatureDisplayInfo GetFeatureInfo(string id)
        {
            var autoInfo = FeatureDiscovery.GetFeatureInfo(id);
            if (autoInfo != null) return autoInfo.Value;

            switch (id)
            {
                case "BtnCloneLeft": return new FeatureDisplayInfo { Tooltip = "Clone Left", Color = "#FBBF24", IconData = "M19,19H5V5H19M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M12,16H10V12H6V10H10V6H12V10H16V12H12", Description = "Duplicates the selection and places the copy adjacent to the left edge.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "Duplicates the selection and places the copy adjacent to the left edge." };
                case "BtnCloneRight": return new FeatureDisplayInfo { Tooltip = "Clone Right", Color = "#FBBF24", IconData = "M19,19H5V5H19M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M12,16H10V12H6V10H10V6H12V10H16V12H12", Description = "Duplicates the selection and places the copy adjacent to the right edge.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "Duplicates the selection and places the copy adjacent to the right edge." };
                case "BtnCloneTop": return new FeatureDisplayInfo { Tooltip = "Clone Top", Color = "#FBBF24", IconData = "M19,19H5V5H19M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M12,16H10V12H6V10H10V6H12V10H16V12H12", Description = "Duplicates the selection and places the copy adjacent to the top edge.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "Duplicates the selection and places the copy adjacent to the top edge." };
                case "BtnCloneBottom": return new FeatureDisplayInfo { Tooltip = "Clone Bottom", Color = "#FBBF24", IconData = "M19,19H5V5H19M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M12,16H10V12H6V10H10V6H12V10H16V12H12", Description = "Duplicates the selection and places the copy adjacent to the bottom edge.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "Duplicates the selection and places the copy adjacent to the bottom edge." };
                case "BtnTableFormatPainter": return new FeatureDisplayInfo { Tooltip = "Table Format Painter", Color = "#06B6D4", Description = "Quickly paint background and text formatting from selected table cells to neighboring cells in specific directions (Up, Down, Left, Right).", RequiresTable = true, DetailedHelpText = "Quickly paint background and text formatting from selected table cells to neighboring cells in specific directions." };
                case "BtnSearchBar": return new FeatureDisplayInfo { Tooltip = "Search Command Palette", Color = "#10B981", Description = "Instantly find and execute any oPenEfficiency command or feature by name using a keyboard-first search bar.", RequiredType = PowerPoint.PpSelectionType.ppSelectionNone, DetailedHelpText = "Instantly find and execute any command or feature by name." };
                case "BtnExportWizard": return new FeatureDisplayInfo { Tooltip = "Export Wizard", Color = "#10B981", Description = "Automates the process of saving individual slides or selections as PDF, PNG, or separate PPTX files.", MinSelection = 0, DetailedHelpText = "Automates the process of saving individual slides or selections as PDF, PNG, or separate PPTX files." };
                case "BtnTemplateManager": return new FeatureDisplayInfo { Tooltip = "PowerPoint Template Manager", Color = "#3B82F6", Description = "Manage and quickly open PowerPoint templates from default locations and custom directories.", RequiredType = PowerPoint.PpSelectionType.ppSelectionNone, DetailedHelpText = "Manage and quickly open PowerPoint templates from default locations and custom directories.\n\nFeatures:\n• Discover default Windows templates\n• Add watched folders\n• Add individual template files\n• Open as new presentation" };
                case "BtnParagraphDialog": return new FeatureDisplayInfo { Tooltip = "Paragraph Dialog", Color = "#FDE047", Description = "Opens the native PowerPoint paragraph formatting dialog to adjust line spacing, indentation, and alignment.", MinSelection = 1, DetailedHelpText = "Opens the native PowerPoint paragraph formatting dialog." };
                case "BtnSpecialChars": return new FeatureDisplayInfo { Tooltip = "Special Characters", Color = "#FDE047", Description = "Opens a specialized picker for inserting symbols, currencies, and mathematical operators not easily found on the keyboard.", MinSelection = 1, DetailedHelpText = "Opens a specialized picker for inserting symbols and special characters." };
                case "BtnFlightMode": return new FeatureDisplayInfo { Tooltip = "Flight Mode", Color = "#D946EF", Description = "Quickly obscures or hides sensitive slide content, logos, and master backgrounds for clean presentation sharing and previews.", RequiredType = PowerPoint.PpSelectionType.ppSelectionNone, DetailedHelpText = "Quickly obscures sensitive content and branding elements." };
                case "BtnSizePanel": return new FeatureDisplayInfo { Tooltip = "Shape Size Panel (Width × Height)", Color = "#6366F1", Description = "Displays and allows editing of exact Width and Height for selected shapes. Supports centimeter and point inputs.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "Displays and allows editing of exact Width and Height for selected shapes. Supports centimeter and point inputs.", IsToggle = true };
                case "BtnStoryline": return new FeatureDisplayInfo { Tooltip = "Storyline", Color = "#FDE047", Description = "Consolidates all slide titles into a single narrative flow view to check the 'red thread' of your story.", RequiredType = PowerPoint.PpSelectionType.ppSelectionNone, DetailedHelpText = "Consolidates all slide titles into a single narrative flow view to check the 'red thread' of your story." };
                case "BtnStickyNote": return new FeatureDisplayInfo { Tooltip = "Sticky Note", Color = "#10B981", Description = "Inserts a sticky note with a random professional color to the side of your slide for internal comments. Right-click for bulk management (hide, move, delete).", RequiredType = PowerPoint.PpSelectionType.ppSelectionNone, DetailedHelpText = "Inserts a sticky note for internal comments.\n\nRight-click options:\n• Move off slide/selected slides\n• Move back on slide\n• Delete notes\n• Convert comments/shapes to notes" };
                case "BtnIllustrativeSticker": return new FeatureDisplayInfo { Tooltip = "Illustrative Sticker", Color = "#F43F5E", IconData = "M12,16L12.01,16 M12,8L12,12 M15.31,2L20,6.69L20,13.31L15.31,18L8.69,18L4,13.31L4,8.69L8.69,4Z", Description = "Inserts professional status stickers like 'Work in Progress' or 'Backup'. Right-click for more variants.", MinSelection = 0, DetailedHelpText = "Inserts professional status stickers like 'Work in Progress' or 'Backup'.\n\nRight-click options:\n• Insert specific stickers: Updated, For discussion, Confidential, etc." };
                case "BtnAssetManager": return new FeatureDisplayInfo { Tooltip = "Asset Library (Slides, Images, Shapes)", Color = "#8B5CF6", IconData = "M21.54,15L17,15L17,19.54 M7,3.34L7,5C7,6.1 7.9,7 9,7C9,5.9 9.9,5 11,5L14.17,5 M11,21.95L11,18L11,17L2.05,17 M12,2C17.522847498307936,2.0 22.0,6.477152501692064 22.0,12.0C22.0,17.522847498307936 17.522847498307936,22.0 12.0,22.0C6.477152501692064,22.0 2.0,17.522847498307936 2.0,12.0C2.0,6.477152501692064 6.477152501692064,2.0 12.0,2.0", Description = "Central repository for your reusable slide templates, custom shapes, and high-quality image assets. Searchable and categorized.", MinSelection = 0, DetailedHelpText = "Central repository for your reusable slide templates, custom shapes, and high-quality image assets. Searchable and categorized.", IsToggle = true };
                case "BtnCopyXY": return new FeatureDisplayInfo { Tooltip = "Copy XY-Coordinates", Color = "#F59E0B", IconData = "M6,20L10,20 M14,10L18,10 M6,14L8,14L8,20 M14,4L16,4L16,10 M14,16L14,18Q14,20 16,20L16,20Q18,20 18,18L18,16Q18,14 16,14L16,14Q14,14 14,16Z M6,6L6,8Q6,10 8,10L8,10Q10,10 10,8L10,6Q10,4 8,4L8,4Q6,4 6,6Z", Description = "Grabs the exact top-left core coordinates of the select object. Use 'Paste XY' to move another object to this exact spot.", MinSelection = 1, MaxSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "Grabs the exact top-left core coordinates of the select object. Use 'Paste XY' to move another object to this exact spot." };
                case "BtnPasteXY": return new FeatureDisplayInfo { Tooltip = "Paste XY-Coordinates", Color = "#F59E0B", Description = "Moves selected objects to the X/Y coordinates previously copied. Essential for aligning objects across multiple slides.", MinSelection = 1, MaxSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "Moves selected objects to the X/Y coordinates previously copied. Essential for aligning objects across multiple slides.\n\nRight-click options:\n• Paste Coordinates (X/Y)\n• Paste Size (Width/Height)\n• Paste Both (Coordinates & Size)" };
                case "BtnConsolidateMasters": return new FeatureDisplayInfo { Tooltip = "Consolidate Masters", Color = "#10B981", IconData = "M2,2H11V6H9V4H4V9H6V11H2V2M22,22H13V18H15V20H20V15H18V13H22V22M8,8H16V16H8V8Z", Description = "Cleans up your presentation by merging duplicate slide masters and layouts. Reduces file size and improves template consistency.", RequiredType = PowerPoint.PpSelectionType.ppSelectionNone, DetailedHelpText = "Cleans up your presentation by merging duplicate slide masters and layouts. Reduces file size and improves template consistency." };
                case "BtnHarveyBall": return new FeatureDisplayInfo { Tooltip = "Harvey Ball", Color = "#F43F5E", IconData = "M21,12", Description = "Cycles through Harvey Ball states (0%, 25%, 50%, 75%, 100%) to represent progress or qualitative data. Right-click to set a specific value.", MinSelection = 0, DetailedHelpText = "Cycles through Harvey Ball states (0%, 25%, 50%, 75%, 100%) to represent progress or qualitative data.\n\nRight-click options:\n• Set specific percentage directly." };
                case "BtnHideSelected": return new FeatureDisplayInfo { Tooltip = "Hide Selected", Color = "#D946EF", IconData = "M10.73,5.08 M14.08,14.16 M17.48,17.5 M2,2L22,22", Description = "Instantly hides selected shapes from the slide. Use 'Show Hidden' or the selection pane to bring them back.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "Instantly hides selected shapes from the slide. Use 'Show Hidden' or the selection pane to bring them back." };
                case "BtnShowHidden": return new FeatureDisplayInfo { Tooltip = "Show Hidden", Color = "#D946EF", IconData = "M2.06,12.35 M12,9C13.65685424949238,9.0 15.0,10.34314575050762 15.0,12.0C15.0,13.65685424949238 13.65685424949238,15.0 12.0,15.0C10.34314575050762,15.0 9.0,13.65685424949238 9.0,12.0C9.0,10.34314575050762 10.34314575050762,9.0 12.0,9.0", Description = "Makes all currently hidden shapes on the slide visible again.", RequiresHiddenShapes = true, DetailedHelpText = "Makes all currently hidden shapes on the slide visible again." };
                case "BtnCheckbox": return new FeatureDisplayInfo { Tooltip = "Checkbox", Color = "#F43F5E", IconData = "M9,12L11,14L15,10 M3,5L3,19Q3,21 5,21L19,21Q21,21 21,19L21,5Q21,3 19,3L5,3Q3,3 3,5Z", Description = "Inserts a vector-based checkmark or cross inside a box. Useful for to-do lists and project status slides.", MinSelection = 0, DetailedHelpText = "Inserts a vector-based checkmark or cross inside a box. Useful for to-do lists and project status slides." };
                case "BtnMagicResizer": return new FeatureDisplayInfo { Tooltip = "Magic Resizer", Color = "#6366F1", IconData = "M136,112L48,112L48,200L136,200L136,120ZM128,200L56,200L56,128L128,128ZM216,184L216,200L176,200L200,200L200,184ZM216,112L216,144L216,112ZM216,56L216,72L216,56L184,56L200,56ZM152,48L112,48L144,48ZM40,80L40,56L72,56L56,56L56,80Z", Description = "Propagates a resize transformation across multiple objects relative to their individual centers, preserving layout spacing.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "Propagates a resize transformation across multiple objects relative to their individual centers, preserving layout spacing.", IsToggle = true };
                case "BtnSplitByParagraphs": return new FeatureDisplayInfo { Tooltip = "Split by Paragraphs", Color = "#EAB308", IconData = "M19,20H5V4H7V7H17V4H19M12,2A1,1 0 0,1 13,3A1,1 0 0,1 12,4A1,1 0 0,1 11,3A1,1 0 0,1 12,2M19,2H14.82C14.4,0.84 13.3,0 12,0C10.7,0 9.6,0.84 9.18,2H5A2,2 0 0,0 3,4V20A2,2 0 0,0 5,22H19A2,2 0 0,0 21,20V4A2,2 0 0,0 19,2M12,18L8,14H11V10H13V14H16L12,18Z", Description = "Takes a single text box with multiple paragraphs and splits it into individual text boxes for each paragraph.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "Takes a single text box with multiple paragraphs and splits it into individual text boxes for each paragraph." };
                case "BtnDeleteText": return new FeatureDisplayInfo { Tooltip = "Delete Text", Color = "#EAB308", IconData = "M16,5L3,5 M11,12L3,12 M16,19L3,19 M15.5,9.5L20.5,14.5 M20.5,9.5L15.5,14.5", Description = "Clears all text content within selected shapes while keeping the shapes and their formatting intact.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "Clears all text content within selected shapes while keeping the shapes and their formatting intact." };
                case "BtnSwapTextFormatted": return new FeatureDisplayInfo { Tooltip = "Swap Text Formatted", Color = "#EAB308", IconData = "M18,14V11H12V9H18V6L22,10L18,14M6,10V13H12V15H6V18L2,14L6,10Z", Description = "Exchanges text between two shapes while maintaining the specific formatting (font, size, color) of the original content.", MinSelection = 2, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "Exchanges text between two shapes while maintaining the specific formatting (font, size, color) of the original content." };
                case "BtnSwapTextPlain": return new FeatureDisplayInfo { Tooltip = "Swap Text Plain", Color = "#EAB308", IconData = "M21,9L17,5V8H10V10H17V13M7,11L3,15L7,19V16H14V14H7V11Z", Description = "Exchanges text between two shapes, but makes the text inherit the formatting of the destination shape.", MinSelection = 2, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "Exchanges text between two shapes, but makes the text inherit the formatting of the destination shape." };
                case "BtnTrafficLight": return new FeatureDisplayInfo { Tooltip = "Traffic Light", Color = "#F43F5E", IconData = "M128,56ZM128,104ZM128,136ZM128,184ZM216,144L200,144L200,80L216,80L200,80L200,40L72,40L72,64L40,64L56,64L56,128L40,128L56,128L56,184L184,184L184,160L200,160ZM184,216L72,216L72,40L184,40L184,216Z", Description = "Inserts or cycles a traffic light symbol (Red, Yellow, Green) to indicate project health or risk levels.", MinSelection = 0, DetailedHelpText = "Inserts or cycles a traffic light symbol (Red, Yellow, Green) to indicate project health or risk levels." };
                case "BtnStarRating": return new FeatureDisplayInfo { Tooltip = "Star Rating", Color = "#F43F5E", IconData = "M11.53,2.29L13.84,6.97L19.0,7.73L15.27,11.37L16.15,16.51L11.53,14.08L6.4,21.01L7.28,15.87L2.16,9.79L7.33,9.04Z", Description = "Inserts a sequence of stars representing a rating from 1 to 5. Great for feedback or performance summaries.", MinSelection = 0, DetailedHelpText = "Inserts a sequence of stars representing a rating from 1 to 5. Great for feedback or performance summaries." };
                case "BtnExploreFeatures": return new FeatureDisplayInfo { Tooltip = "Explore Features", Color = "#6366F1", IconData = "M12,2A10,10 0 1,0 22,12A10,10 0 0,0 12,2M12,17A1,1 0 1,1 11,16A1,1 0 0,1 12,17M13,14H11V9H13V14Z", Description = "The Feature Explorer provides a searchable index of all features available in oPenEfficiency. It shows detailed usage instructions, animated examples, and current shortcuts.", DetailedHelpText = "The Feature Explorer provides a searchable index of all features available in oPenEfficiency. It shows detailed usage instructions, animated examples, and current shortcuts." };
                case "BtnSettings": return new FeatureDisplayInfo { Tooltip = "Settings", Color = "#94A3B8", Description = "Configure your oPenEfficiency experience, including sidebar layout, spacing increments, and keyboard shortcuts.", DetailedHelpText = "Configure your oPenEfficiency experience, including sidebar layout, spacing increments, and keyboard shortcuts." };
                default: return new FeatureDisplayInfo { Tooltip = id, Color = "#9CA3AF", IconData = "M12,2L14.5,9H22L16,13.5L18.5,21L12,16.5L5.5,21L8,13.5L2,9H9.5L12,2Z" };
            }
        }
    }
}
