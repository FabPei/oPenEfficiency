using System;
using System.Collections.Generic;
using System.Linq;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.UI
{
    /// <summary>
    /// P0.1 policy data for the current command set. These assignments are explicit
    /// catalog data; consumers must not infer category or maturity from IDs or names.
    /// Visible terminology and preset membership are intentionally left to later slices.
    /// </summary>
    public static class FeatureCatalogPolicy
    {
        private static readonly HashSet<string> Widgets =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "BtnSearchBar", "BtnSizePanel", "BtnFontPanel"
            };

        private static readonly Dictionary<string, FeatureCategory> Categories =
            new Dictionary<string, FeatureCategory>(StringComparer.Ordinal);

        private static readonly HashSet<string> Experimental =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "Btn2048", "BtnChess", "BtnMinesweeper", "BtnWinnerPicker"
            };

        private static readonly HashSet<string> Preview =
            new HashSet<string>(StringComparer.Ordinal)
            {
            };

        private static readonly HashSet<string> ShortcutIneligible =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "BtnSearchBar", "BtnSizePanel", "BtnFontPanel"
            };

        // Curated for stable, broadly useful consulting workflows. Placement here is
        // authoritative catalog policy; it is intentionally independent of the legacy layout.
        private static readonly HashSet<string> ProfessionalDefault =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "BtnSearchBar",
                "BtnSizePanel", "BtnMagicResizer", "BtnMatchSize", "BtnMatchWidth", "BtnMatchHeight",
                "BtnAlignLeft", "BtnAlignCenter", "BtnAlignRight", "BtnAlignTop",
                "BtnAlignCenterVertical", "BtnAlignBottom",
                "BtnDistributeHorizontal", "BtnDistributeVertical",
                "BtnRemoveHorizontalGaps", "BtnRemoveVerticalGaps", "BtnAdjustSpacing",
                "BtnDockLeft", "BtnDockRight", "BtnDockTop", "BtnDockBottom",
                "BtnSwapPositions", "BtnCopyXY", "BtnPasteXY", "BtnSelectSameType",
                "BtnBringToFront", "BtnSendToBack", "BtnPositionPainter",
                "BtnFontPanel", "BtnApplyTextTool", "BtnFitFormToText", "BtnReplaceFont",
                "BtnParagraphDialog", "BtnFormatNumbers", "BtnTextFormatPainter",
                "BtnPickColor", "BtnThemeColor",
                "BtnTableDimensions", "BtnTableFormatPainter", "BtnTableSort",
                "BtnTableSum", "BtnAlignToTableCell",
                "BtnIllustrativeSticker", "BtnHarveyBall", "BtnTrafficLight",
                "BtnCheckbox", "BtnNumeration",
                "BtnAgendaWizard", "BtnStoryline", "BtnStyleCheck",
                "BtnSlideGuidelines", "BtnSlideNotes",
                "BtnUpdateExcelCharts", "BtnGanttChart", "BtnDocumentAutomation",
                "BtnAssetManager", "BtnTemplateManager", "BtnIconBrowser", "BtnIllustrationLibrary",
                "BtnSettings", "BtnExploreFeatures", "BtnQuickContextMenu"
                , "BtnShapeFill"
            };

        private static readonly Dictionary<string, string[]> Variants =
            new Dictionary<string, string[]>(StringComparer.Ordinal)
            {
                ["BtnAlignLeft"] = new[] { "Align Left Within Selection", "Align Left to Slide", "Align Left to First Selected", "Align Left to Last Selected", "Move Left Until Next Object" },
                ["BtnAlignCenter"] = new[] { "Align Centers Within Selection", "Align Centers to Slide", "Snap Centers to Nearest Guide", "Align to Configured Anchor" },
                ["BtnAlignRight"] = new[] { "Align Right Within Selection", "Align Right to Slide", "Align Right to First Selected", "Align Right to Last Selected", "Move Right Until Next Object" },
                ["BtnAlignTop"] = new[] { "Align Top Within Selection", "Align Top to Slide", "Align Top to First Selected", "Align Top to Last Selected", "Move Up Until Next Object" },
                ["BtnAlignCenterVertical"] = new[] { "Align Middles Within Selection", "Align Middles to Slide", "Snap Middles to Nearest Guide", "Align to Configured Anchor" },
                ["BtnAlignBottom"] = new[] { "Align Bottom Within Selection", "Align Bottom to Slide", "Align Bottom to First Selected", "Align Bottom to Last Selected", "Move Down Until Next Object" },
                ["BtnAlignToTableCell"] = new[] { "Top Left", "Top Center", "Top Right", "Middle Left", "Center", "Middle Right", "Bottom Left", "Bottom Center", "Bottom Right" },
                ["BtnAlignShapeAdjustments"] = new[] { "Use First Selected as Geometry Anchor", "Use Last Selected as Geometry Anchor" },
                ["BtnApplyTextTool"] = new[] { "Open Text Editor", "Replace Text from First Selected", "Replace Text from Last Selected" },
                ["BtnReplaceFont"] = new[] { "Replace Formats on Current Slide", "Replace Formats on Selected Slides", "Replace Formats in Entire Presentation" },
                ["BtnIconBrowser"] = new[] { "Search Icons", "Browse Favorite Icons", "Insert Icons as SVG", "Insert Icons as PowerPoint Shapes" },
                ["BtnArrangeCircle"] = new[] { "Triangle", "Square", "Star", "Octagon", "Pentagon" },
                ["BtnArrangeGrid"] = new[] { "Arrange in 2 Columns", "Arrange in 3 Columns", "Arrange in 4 Columns", "Arrange in 5 Columns" },
                ["BtnArrangeInShape"] = new[] { "Arrange Shapes in Circle", "Arrange in Triangle", "Arrange in Square", "Arrange in Star", "Arrange in Pentagon", "Arrange in Heart" },
                ["BtnSearchBar"] = new[] { "Focus Command Search", "Clear Command Search" },
                ["BtnSettings"] = new[] { "Open General Settings", "Open Layout Settings", "Open Appearance Settings", "Open Stickers & Notes Settings", "Open Shortcut Settings", "Open Backup & Config Settings", "Open About" },
                ["BtnExploreFeatures"] = new[] { "Search All Features", "Explore Align & Size", "Explore Text & Format", "Explore Data & Automation", "Explore Labs" },
                ["BtnFontPanel"] = new[] { "Focus Font Family", "Focus Font Size" },
                ["BtnSizePanel"] = new[] { "Focus Width", "Focus Height" },
                ["BtnAgendaWizard"] = new[] { "Update Existing Agenda", "Build Deck Structure", "Extract Agenda from Deck", "Validate Agenda & Sections", "Open Agenda Layout Designer" },
                ["BtnBringForward"] = new[] { "Rank Selected Shapes (First Selected on Top)", "Place Directly Above Last Selected", "Place Directly Below Last Selected" },
                ["BtnBringToFront"] = new[] { "Rank Selected Shapes (First Selected on Top)", "Place Directly Above Last Selected", "Place Directly Below Last Selected" },
                ["BtnChangeCase"] = new[] { "Sentence case", "lowercase", "UPPERCASE", "Capitalize Each Word", "tOGGLE cASE" },
                ["BtnCharacterSpacing"] = new[] { "More Spacing..." },
                ["BtnCheckbox"] = new[] { "Framed Tick", "Framed Cross", "Framed Empty", "Unframed Tick", "Unframed Cross", "Unframed Empty" },
                ["BtnConvertTableToShapes"] = new[] { "Ask After Conversion", "Convert and Keep Original", "Convert and Replace Original" },
                ["BtnCreateMotionPath"] = new[] { "First Shape → Second Shape", "Second Shape → First Shape", "Create Both Directions" },
                ["BtnCopyXY"] = new[] { "Clear Copied XY Coordinates" },
                ["BtnNumeration"] = new[] { "Insert Number 1", "Insert Number 2", "Insert Number 3", "Insert Number 4", "Insert Number 5" },
                ["BtnObjectConnector"] = new[] { "Connect Nearest Sides (Smart)", "Connect Right to Left", "Connect Left to Right", "Connect Bottom to Top", "Connect Top to Bottom" },
                ["BtnDeleteText"] = new[] { "Delete All Text (Keep Shapes)", "Delete Highlighted Text", "Delete Selected Paragraphs", "Replace Selected Text..." },
                ["BtnDecreaseFontSize"] = new[] { "Decrease 1 Step", "Decrease 2 Steps", "Decrease 5 Steps" },
                ["BtnDistributeHorizontal"] = new[] { "Distribute Across Slide", "Distribute Within Selection", "Align Middles & Distribute Within Selection", "Align Middles & Distribute Across Slide" },
                ["BtnDistributeVertical"] = new[] { "Distribute Across Slide", "Distribute Within Selection", "Align Centers & Distribute Within Selection", "Align Centers & Distribute Across Slide" },
                ["BtnDockLeft"] = new[] { "Stack Left", "Stack Right", "Stack Above", "Stack Below" },
                ["BtnDockRight"] = new[] { "Stack Left", "Stack Right", "Stack Above", "Stack Below" },
                ["BtnDockTop"] = new[] { "Stack Left", "Stack Right", "Stack Above", "Stack Below" },
                ["BtnDockBottom"] = new[] { "Stack Left", "Stack Right", "Stack Above", "Stack Below" },
                ["BtnFlightMode"] = new[] { "Cover Logos", "Light Overlay", "Heavy Overlay", "Disable Master Covers & Overlays" },
                ["BtnFlipHorizontal"] = new[] { "Flip Horizontal", "Flip Vertical" },
                ["BtnFlipVertical"] = new[] { "Flip Horizontal", "Flip Vertical" },
                ["BtnStyleCheck"] = new[] { "Check Current Slide", "Check Selected Slides", "Check Entire Presentation" },
                ["BtnFormatBold"] = new[] { "Apply Bold Formatting", "Remove Bold Formatting", "Toggle Bold Formatting" },
                ["BtnFormatItalic"] = new[] { "Apply Italic Formatting", "Remove Italic Formatting", "Toggle Italic Formatting" },
                ["BtnFormatStrikethrough"] = new[] { "Toggle Strikethrough Formatting" },
                ["BtnFormatSubscript"] = new[] { "Apply Subscript Formatting", "Remove Subscript Formatting", "Toggle Subscript Formatting" },
                ["BtnFormatSuperscript"] = new[] { "Apply Superscript Formatting", "Remove Superscript Formatting", "Toggle Superscript Formatting" },
                ["BtnFormatUnderline"] = new[] { "Apply Underline Formatting", "Remove Underline Formatting", "Toggle Underline Formatting" },
                ["BtnFormatNumbers"] = new[] { "Format Numbers: US (1,234.56)", "Format Numbers: German (1.234,56)", "Format Numbers: French (1 234,56)", "Format Numbers: Swiss (1'234.56)" },
                ["BtnFitFormToText"] = new[] { "Fit with Zero Margins", "Fit and Keep Existing Margins" },
                ["BtnInsertColumnLeft"] = new[] { "Insert Column Left & Preserve Width", "Duplicate Selected Column Left – Format Only", "Insert Multiple Columns Left..." },
                ["BtnInsertColumnRight"] = new[] { "Insert Column Right & Preserve Width", "Duplicate Selected Column Right – Format Only", "Insert Multiple Columns Right..." },
                ["BtnInsertRowBottom"] = new[] { "Insert Row Below & Preserve Height", "Duplicate Selected Row Below – Format Only", "Insert Multiple Rows Below..." },
                ["BtnInsertRowTop"] = new[] { "Insert Row Above & Preserve Height", "Duplicate Selected Row Above – Format Only", "Insert Multiple Rows Above..." },
                ["BtnHideSelected"] = new[] { "Hide Selected Shapes", "Show openEfficiency-Hidden Shapes", "Reveal All Hidden Shapes on Slide..." },
                ["BtnHideMasterObjects"] = new[] { "Hide Master Objects", "Show Master Objects", "Toggle Master Objects" },
                ["BtnInsertText"] = new[] { "Insert Line Break", "Insert Bullet", "Insert En Dash", "Insert Em Dash" },
                ["BtnIncreaseHorizontalSpacing"] = new[] { "Increase Horizontal Spacing", "Decrease Horizontal Spacing", "Increase Vertical Spacing", "Decrease Vertical Spacing" },
                ["BtnDecreaseHorizontalSpacing"] = new[] { "Increase Horizontal Spacing", "Decrease Horizontal Spacing", "Increase Vertical Spacing", "Decrease Vertical Spacing" },
                ["BtnIncreaseVerticalSpacing"] = new[] { "Increase Horizontal Spacing", "Decrease Horizontal Spacing", "Increase Vertical Spacing", "Decrease Vertical Spacing" },
                ["BtnIncreaseFontSize"] = new[] { "Increase 1 Step", "Increase 2 Steps", "Increase 5 Steps" },
                ["BtnDecreaseVerticalSpacing"] = new[] { "Increase Horizontal Spacing", "Decrease Horizontal Spacing", "Increase Vertical Spacing", "Decrease Vertical Spacing" },
                ["BtnMatchSize"] = new[] { "Match Width", "Match Height", "Match Width and Height" },
                ["BtnMatchWidth"] = new[] { "Match Width", "Match Height", "Match Width and Height" },
                ["BtnMatchHeight"] = new[] { "Match Width", "Match Height", "Match Width and Height", "Match Height and Swap Positions" },
                ["BtnLockDimensions"] = new[] { "Lock/Unlock Aspect Ratio (All Slides)", "Lock Aspect Ratio (All Slides)", "Unlock Aspect Ratio (All Slides)" },
                ["BtnMatchAngles"] = new[] { "Use First Selected as Geometry Anchor", "Use Last Selected as Geometry Anchor" },
                ["BtnLearnMargin"] = new[] { "Copy Text Margins from Selection", "Clear Copied Text Margins" },
                ["BtnWinnerPicker"] = new[] { "Pick 1 Winner", "Pick 2 Winners", "Pick 3 Winners", "Clear Winner Markers" },
                ["BtnChess"] = new[] { "New Easy Game", "New Medium Game", "New Hard Game", "Close Chess" },
                ["Btn2048"] = new[] { "Open / Continue 2048", "Start New 2048 Game", "Close 2048" },
                ["BtnMinesweeper"] = new[] { "Open / Continue Minesweeper", "Start New Minesweeper Game", "Close Minesweeper" },
                ["BtnMoveTableRowColumn"] = new[] { "Move rows", "Move columns", "Move to position...", "Swap selected", "Duplicate selected", "Repeat last table move" },
                ["BtnPasteXY"] = new[] { "Clear Copied XY Coordinates" },
                ["BtnPickColor"] = new[] { "Pick Fill Color", "Pick Line Color", "Pick Text Color" },
                ["BtnPickFillColor"] = new[] { "Pick Fill Color", "Pick Line Color", "Pick Text Color" },
                ["BtnPickLineColor"] = new[] { "Pick Fill Color", "Pick Line Color", "Pick Text Color" },
                ["BtnPickTextColor"] = new[] { "Pick Fill Color", "Pick Line Color", "Pick Text Color" },
                ["BtnSelectSameType"] = new[] { "Match Fill", "Match Line", "Match Fill and Line", "Match Size", "Match Width", "Match Height", "Match Type and Formatting" },
                ["BtnSmartCorners"] = new[] { "Make Corners Rounded", "Make Corners Square", "Toggle Rounded Corners" },
                ["BtnSendBackward"] = new[] { "Rank Selected Shapes (First Selected at Bottom)", "Place Directly Above Last Selected", "Place Directly Below Last Selected" },
                ["BtnSendToBack"] = new[] { "Rank Selected Shapes (First Selected at Bottom)", "Place Directly Above Last Selected", "Place Directly Below Last Selected" },
                ["BtnShapeLock"] = new[] { "Toggle Lock All Shapes", "Lock All Shapes", "Unlock All Shapes" },
                ["BtnShowHidden"] = new[] { "Show openEfficiency-Hidden Shapes", "Reveal All Hidden Shapes on Slide..." },
                ["BtnSlideGuidelines"] = new[] { "Three Columns", "Four Columns", "Five Columns", "1:2 Columns", "Rule of Thirds", "Remove Guidelines" },
                ["BtnSlideNotes"] = new[] { "Delete Notes on This Slide", "Delete All Notes in Presentation", "Show or Hide Review Note on Current Slide", "Open Review Notes Overview" },
                ["BtnSlideSize"] = new[] { "Analyze Slide Media Sizes", "Remove Analysis Labels" },
                ["BtnSetMargin"] = new[] { "Paste All Text Margins", "Paste Left and Right Text Margins", "Paste Top and Bottom Text Margins", "Reset Text Margins to Zero" },
                ["BtnSpellCheck"] = new[] { "Choose Language for Selected Shapes", "Choose Language for Selected Slides", "Choose Language for Entire Presentation" },
                ["BtnOptimizeFreeForm"] = new[] { "Light Cleanup (2 pt)", "Standard Cleanup (5 pt)", "Aggressive Cleanup (10 pt)" },
                ["BtnRefreshLinkedTable"] = new[] { "Refresh Linked Table", "Open Source in Excel", "Change Linked Range...", "Relink Source...", "Toggle Auto-Resize" },
                ["BtnNewSlide"] = new[] { "Open New Slide Gallery", "Insert Blank Slide", "Duplicate Current Slide" },
                ["BtnAddSticker"] = new[] { "Insert Draft Sticker", "Insert Review Sticker", "Insert Approved Sticker", "Insert Confidential Sticker" },
                ["BtnSnapToGrid"] = new[] { "Enable Snap to Grid", "Enable Snap to Objects", "Disable Snapping" },
                ["BtnSnapToObjects"] = new[] { "Enable Snap to Objects", "Enable Snap to Grid", "Disable Snapping" },
                ["BtnSnapShape"] = new[] { "Snap Left", "Snap Right", "Snap Above", "Snap Below" },
                ["BtnStarRating"] = new[] { "Set / Insert 1 of 5", "Set / Insert 2 of 5", "Set / Insert 3 of 5", "Set / Insert 4 of 5", "Set / Insert 5 of 5" },
                ["BtnStretchLeft"] = new[] { "Stretch Left", "Stretch Right", "Stretch Top", "Stretch Bottom", "Move Left Edge", "Move Right Edge", "Move Top Edge", "Move Bottom Edge" },
                ["BtnStretchRight"] = new[] { "Stretch Left", "Stretch Right", "Stretch Top", "Stretch Bottom", "Move Left Edge", "Move Right Edge", "Move Top Edge", "Move Bottom Edge" },
                ["BtnStretchTop"] = new[] { "Stretch Left", "Stretch Right", "Stretch Top", "Stretch Bottom", "Move Left Edge", "Move Right Edge", "Move Top Edge", "Move Bottom Edge" },
                ["BtnStretchBottom"] = new[] { "Stretch Left", "Stretch Right", "Stretch Top", "Stretch Bottom", "Move Left Edge", "Move Right Edge", "Move Top Edge", "Move Bottom Edge" },
                ["BtnStretchLeftEdge"] = new[] { "Stretch Left", "Stretch Right", "Stretch Top", "Stretch Bottom", "Move Left Edge", "Move Right Edge", "Move Top Edge", "Move Bottom Edge" },
                ["BtnStretchRightEdge"] = new[] { "Stretch Left", "Stretch Right", "Stretch Top", "Stretch Bottom", "Move Left Edge", "Move Right Edge", "Move Top Edge", "Move Bottom Edge" },
                ["BtnStretchTopEdge"] = new[] { "Stretch Left", "Stretch Right", "Stretch Top", "Stretch Bottom", "Move Left Edge", "Move Right Edge", "Move Top Edge", "Move Bottom Edge" },
                ["BtnStretchBottomEdge"] = new[] { "Stretch Left", "Stretch Right", "Stretch Top", "Stretch Bottom", "Move Left Edge", "Move Right Edge", "Move Top Edge", "Move Bottom Edge" },
                ["BtnSplitByParagraphs"] = new[] { "Split and Remove Empty Paragraphs", "Split and Keep Empty Paragraphs", "Merge Selected Text Boxes" },
                ["BtnSplitShape"] = new[] { "Split 1 × 2", "Split 2 × 1", "Split 2 × 2", "Split 3 × 2", "Split 3 × 3" },
                ["BtnStoryline"] = new[] { "Open Current Section", "Refresh Storyline", "Check Headline Quality", "Show Missing Titles", "Copy Numbered Storyline", "Copy Headlines Only", "Copy Grouped by Sections", "Copy Issues Report" },
                ["BtnSwapPositions"] = new[] { "Top Left", "Top Center", "Top Right", "Left Center", "Center (Default)", "Right Center", "Bottom Left", "Bottom Center", "Bottom Right" },
                ["BtnSwapTextPlain"] = new[] { "Swap Text & Shrink on Overflow", "Cycle Text Through Selection", "Swap Selected Paragraphs Only..." },
                ["BtnTableFormatPainter"] = new[] { "All Formats Horizontally", "All Formats Vertically", "All Formats in All Directions", "Text Formatting Horizontally", "Text Formatting Vertically", "Text Formatting in All Directions", "Cell Fill Horizontally", "Cell Fill Vertically", "Cell Fill in All Directions" },
                ["BtnTableColumnInsertion"] = new[] { "Insert Column Left", "Insert Column Right", "Insert Left and Preserve Width", "Insert Right and Preserve Width", "Duplicate Left Column Format", "Duplicate Right Column Format" },
                ["BtnTableRowInsertion"] = new[] { "Insert Row Above", "Insert Row Below", "Insert Above and Preserve Height", "Insert Below and Preserve Height", "Duplicate Above Row Format", "Duplicate Below Row Format" },
                ["BtnTableSortAZ"] = new[] { "Sort Selected Cells A–Z", "Sort Selected Cells Z–A", "Sort Whole Row Text A–Z", "Sort Whole Row Text Z–A" },
                ["BtnTagInspector"] = new[] { "Inspect Current Slide Tags", "Inspect Selected Shape Tags", "Inspect Presentation Tags" },
                ["BtnTableColumnWidth"] = new[] { "Copy Column Widths", "Paste Column Widths", "Make Column Widths Equal", "Clear Copied Column Widths" },
                ["BtnTableRowHeight"] = new[] { "Copy Row Heights", "Paste Row Heights", "Make Row Heights Equal", "Clear Copied Row Heights" },
                ["BtnTableTranspose"] = new[] { "Replace Original in Place", "Create Copy to the Right", "Create Copy Below" },
                ["BtnProgressSeries"] = new[] { "3-Step Number Series", "5-Step Number Series", "5-Step Vertical Series", "5-Step Bullet Series" },
                ["BtnPositionPainter"] = new[] { "Copy/Paste Position Only", "Copy/Paste Size Only", "Copy/Paste Position and Size" },
                ["BtnTableHeatmap"] = new[] { "Configure Row-Wise Heatmap", "Configure Column-Wise Heatmap", "Configure Entire-Table Heatmap" },
                ["BtnRectifyRotation"] = new[] { "Set Rotation to 0°", "Set Rotation to 90°", "Set Rotation to 180°", "Set Rotation to 270°" },
                ["BtnTextFormatPainter"] = new[] { "Apply Typography Only", "Apply Paragraphs & Bullets Only", "Apply While Preserving Emphasis" },
                ["BtnTrafficLight"] = new[] { "Set / Insert Red", "Set / Insert Yellow", "Set / Insert Green", "Set / Insert Off" },
                ["BtnTextDirection"] = new[] { "Horizontal", "Rotate 90°", "Rotate 270°", "Stacked" },
                ["BtnThemeColor"] = new[] { "Open Fill Theme Colors", "Open Line Theme Colors", "Open Text Theme Colors", "Show / Hide Docked Palette" },
                ["BtnTransparentColor"] = new[] { "Pick Color: Exact Match (0%)", "Pick Color: Balanced Tolerance (10%)", "Pick Color: Broad Tolerance (25%)" },
                ["BtnStickyNote"] = new[] { "Insert Theme Sticky Note", "Insert Consulting Sticky Note", "Insert Compact Banner Note" },
                ["BtnSyncObjects"] = new[] { "Open Sync Objects", "Sync Format", "Sync Position", "Sync Content", "Sync All Properties" },
                ["BtnArrangePro"] = new[] { "Circle Around Selection", "Circle Around Slide", "Arc Around Selection", "Rotate Shapes to Center", "Open Custom Arrangement" },
                ["BtnParagraphDialog"] = new[] { "Open Paragraph Dialog", "Single Line Spacing", "1.15 Line Spacing", "1.5 Line Spacing", "Remove Paragraph Spacing" },
                ["BtnSwapTextFormatted"] = new[] { "Swap Formatted Text", "Swap and Shrink Overflowing Text", "Swap and Fit Shapes to Text" },
                ["BtnFormatShapeDialog"] = new[] { "Open Format Shape", "Reset to Default Shape Formatting", "Remove Shadow", "Remove Reflection", "Remove Glow and Soft Edges", "Remove 3-D Format and Rotation", "Remove All Effects" },
                ["BtnInsertExcelValue"] = new[] { "Insert at Text Cursor", "Insert as Linked Text Box", "Insert as Unlinked Value", "Refresh Selected Linked Value", "Unlink Selected Value" },
                ["BtnConnectedTextLabel"] = new[] { "Send to Word", "Send to OneNote", "Refresh All Open Targets", "Disconnect Selected Source" },
                ["BtnAssetManager"] = new[] { "Open Slides", "Open Images", "Open Shapes", "Open My Assets", "Save Current Selection" },
                ["BtnExcelLinkManager"] = new[] { "Show All Links", "Show Broken Links", "Refresh All Links", "Relink Source", "Open Source Workbook" },
                ["BtnQRCode"] = new[] { "Enter URL or Text", "Use Selected Text", "Use Clipboard Text" },
                ["BtnExportWizard"] = new[] { "PPTX — Entire Presentation", "PDF — Entire Presentation", "PPTX + PDF — Entire Presentation", "PDF — Selected Slides", "Editable HTML — Entire Presentation" },
                ["BtnGanttChart"] = new[] { "Open Editor", "Create New Chart", "Update Selected Chart", "Import Excel", "Sync Linked Excel" },
                ["BtnTemplateManager"] = new[] { "Open Templates", "Search Templates", "Add Template File", "Manage Watched Folders", "Restore Hidden Templates" },
                ["BtnSlidePaste"] = new[] { "Insert with Links and Shadows", "Insert with Links", "Insert with Shadows", "Insert without Links or Shadows" },
                ["BtnHtmlSlideImport"] = new[] { "Preserve HTML Styling", "Apply Presentation Theme", "Preserve Background Only", "Local Resources Only", "Review Import Options" },
                ["BtnAnonymize"] = new[] { "Anonymize Current Slide", "Anonymize Selected Slides", "Anonymize Entire Presentation", "Preview Affected Word Count" },
                ["BtnChartOverlay"] = new[] { "Delta (Absolute)", "CAGR (%)", "Growth (%)", "Average Level Line", "Highlight Outliers" },
                ["BtnConvertToTable"] = new[] { "Ask After Conversion", "Keep Original Shapes", "Replace Original Shapes" },
                ["BtnDocumentAutomation"] = new[] { "Run Automation", "Design Schema" },
                ["BtnLocalSlideRebuild"] = new[] { "Rebuild Image Locally", "Rebuild Entire PDF Locally", "Rebuild PDF Page Range" },
                ["BtnMapWizard"] = new[] { "New World Map", "New Continent Map", "New Country Map", "Edit Selected Map", "Open Recent Map Configuration" },
                ["BtnMoveToBackup"] = new[] { "Move Selected Slides to Backup", "Copy Selected Slides to Backup", "Locate Backup Slide" },
                ["BtnPropertyExtraction"] = new[] { "Inspect Selection", "Inspect Current Slide", "Copy Selection Report", "Copy Slide Report", "Copy Color List" },
                ["BtnSeriesGenerator"] = new[] { "1–5 Numbers — Horizontal", "A–E Letters — Horizontal", "1–5 Numbers — Vertical", "1–9 Numbers — 3×3 Matrix", "Open Custom Series" }
            };

        static FeatureCatalogPolicy()
        {
            Add(FeatureCategory.AlignAndSize,
                "BtnAdjustSpacing", "BtnAlignBottom", "BtnAlignCenter", "BtnAlignCenterVertical",
                "BtnAlignLeft", "BtnAlignRight", "BtnAlignShapeAdjustments", "BtnAlignTop",
                "BtnDecreaseHorizontalSpacing", "BtnDecreaseVerticalSpacing", "BtnDiagonalAlign",
                "BtnDistributeHorizontal", "BtnDistributeVertical", "BtnDockBottom", "BtnDockLeft",
                "BtnDockRight", "BtnDockTop", "BtnIncreaseHorizontalSpacing", "BtnIncreaseVerticalSpacing",
                "BtnLockDimensions", "BtnMagicResizer", "BtnMatchAngles", "BtnMatchHeight", "BtnMatchSize", "BtnMatchWidth",
                "BtnRemoveHorizontalGaps", "BtnRemoveVerticalGaps", "BtnSizePanel", "BtnSlideSize",
                "BtnStretchBottom", "BtnStretchBottomEdge", "BtnStretchLeft", "BtnStretchLeftEdge",
                "BtnStretchRight", "BtnStretchRightEdge", "BtnStretchTop", "BtnStretchTopEdge");

            Add(FeatureCategory.ArrangeAndOrganize,
                "BtnArrangeCircle", "BtnArrangeGrid", "BtnArrangeInShape", "BtnArrangePro",
                "BtnBringForward", "BtnBringToFront", "BtnCopyXY", "BtnCreateMotionPath",
                "BtnFlipHorizontal", "BtnFlipVertical", "BtnHideMasterObjects", "BtnHideSelected",
                "BtnObjectConnector", "BtnOptimizeFreeForm", "BtnPasteXY", "BtnPositionPainter",
                "BtnRectifyRotation", "BtnRotateShape", "BtnSelectSameType", "BtnSendBackward",
                "BtnSendToBack", "BtnShapeLock", "BtnShowHidden", "BtnSmartCorners", "BtnSnapShape",
                "BtnSnapToGrid", "BtnSnapToObjects", "BtnSwapPositions", "BtnSyncObjects");

            Add(FeatureCategory.TextAndFormat,
                "BtnAlignText", "BtnApplyTextTool", "BtnChangeCase", "BtnCharacterSpacing",
                "BtnDecreaseFontSize", "BtnDeleteText", "BtnFitFormToText", "BtnFontPanel",
                "BtnFormatBold", "BtnFormatItalic", "BtnFormatNumbers", "BtnFormatShapeDialog",
                "BtnFormatStrikethrough", "BtnFormatSubscript", "BtnFormatSuperscript",
                "BtnFormatUnderline", "BtnIncreaseFontSize", "BtnInsertText", "BtnParagraphDialog",
                "BtnPickColor", "BtnPickFillColor", "BtnPickLineColor", "BtnPickTextColor",
                "BtnReplaceFont", "BtnSpecialChars", "BtnSpellCheck", "BtnSplitByParagraphs",
                "BtnSwapTextFormatted", "BtnSwapTextPlain", "BtnTextDirection", "BtnTextFormatPainter",
                "BtnThemeColor", "BtnTranslate", "BtnTransparency", "BtnTransparentColor", "BtnShapeFill");

            Add(FeatureCategory.Tables,
                "BtnAlignToTableCell", "BtnConvertTableToShapes", "BtnConvertToTable",
                "BtnInsertColumnLeft", "BtnInsertColumnRight", "BtnInsertRowBottom", "BtnInsertRowTop",
                "BtnMoveTableRowColumn", "BtnTableBranding", "BtnTableColumnInsertion",
                "BtnTableColumnWidth", "BtnTableDimensions", "BtnTableFormatPainter", "BtnTableHeatmap",
                "BtnTableRowHeight", "BtnTableRowInsertion", "BtnTableSort", "BtnTableSortAZ",
                "BtnTableSplit", "BtnTableSum", "BtnTableTranspose");

            Add(FeatureCategory.VisualElements,
                "BtnAddSticker", "BtnChartOverlay", "BtnCheckbox", "BtnHarveyBall",
                "BtnIllustrativeSticker", "BtnLearnMargin", "BtnMapWizard", "BtnNumeration", "BtnProgressSeries",
                "BtnQRCode", "BtnRepeatShape", "BtnSeriesGenerator", "BtnSetMargin", "BtnSplitShape",
                "BtnStarRating", "BtnThermometer", "BtnTrafficLight");

            Add(FeatureCategory.StorylineAndReview,
                "BtnAgendaWizard", "BtnAnonymize", "BtnCleaner", "BtnConsolidateMasters",
                "BtnExportWizard", "BtnFlightMode", "BtnMoveToBackup", "BtnNewSlide",
                "BtnSaveAgendaLayout", "BtnSlideGuidelines", "BtnSlideNotes", "BtnStickyNote",
                "BtnStickyNoteManager", "BtnStoryline", "BtnStyleCheck");

            Add(FeatureCategory.DataAndAutomation,
                "BtnDocumentAutomation", "BtnExcelLinkManager", "BtnGanttChart", "BtnHtmlSlideImport",
                "BtnConnectedTextLabel", "BtnInsertExcelValue", "BtnLocalSlideRebuild", "BtnPropertyExtraction",
                "BtnRefreshLinkedTable", "BtnSlidePaste", "BtnTagInspector", "BtnUpdateExcelCharts");

            Add(FeatureCategory.Libraries,
                "BtnAssetManager", "BtnCountryFlags", "BtnIconBrowser", "BtnIllustrationLibrary", "BtnTemplateManager");

            Add(FeatureCategory.SettingsAndHelp,
                "BtnExploreFeatures", "BtnFeatureAcceptanceWindow", "BtnPerformanceMode", "BtnQuickContextMenu", "BtnSearchBar", "BtnSettings");

            Add(FeatureCategory.Labs,
                "Btn2048", "BtnChess", "BtnMinesweeper", "BtnWinnerPicker");
        }

        public static void Apply(FeatureCatalogRecord record)
        {
            if (record == null) throw new ArgumentNullException(nameof(record));

            if (Categories.TryGetValue(record.Id ?? string.Empty, out var category))
                record.Category = category;

            if (Widgets.Contains(record.Id ?? string.Empty))
                record.ExecutionKind = FeatureExecutionKind.Widget;

            if (Experimental.Contains(record.Id))
            {
                record.Maturity = FeatureMaturity.Experimental;
                record.DefaultPlacement = FeatureDefaultPlacement.Labs;
            }
            else if (Preview.Contains(record.Id))
            {
                record.Maturity = FeatureMaturity.Preview;
                record.DefaultPlacement = FeatureDefaultPlacement.CatalogOnly;
            }
            else
            {
                if (record.Maturity == FeatureMaturity.Unknown)
                    record.Maturity = FeatureMaturity.Stable;
                if (record.DefaultPlacement == FeatureDefaultPlacement.Unknown)
                    record.DefaultPlacement = FeatureDefaultPlacement.CatalogOnly;
            }

            if (ProfessionalDefault.Contains(record.Id))
                record.DefaultPlacement = FeatureDefaultPlacement.ProfessionalDefault;

            if (ShortcutIneligible.Contains(record.Id))
                record.SupportsUserShortcut = false;
            if (Variants.TryGetValue(record.Id ?? string.Empty, out var variants))
            {
                record.Variants = variants;
                EnsureVariantDiscoveryText(record, variants);
            }

            if (NormalizeDisplayText(record.Name) == NormalizeDisplayText(record.Tooltip))
                record.Tooltip = null;
        }

        private static string NormalizeDisplayText(string value)
        {
            return new string((value ?? string.Empty)
                .Where(char.IsLetterOrDigit)
                .Select(char.ToLowerInvariant)
                .ToArray());
        }

        private static void EnsureVariantDiscoveryText(
            FeatureCatalogRecord record,
            IReadOnlyList<string> variants)
        {
            if (variants == null || variants.Count == 0) return;

            string help = record.DetailedHelpText ?? string.Empty;
            var missing = new List<string>();
            foreach (string variant in variants)
            {
                if (!string.IsNullOrWhiteSpace(variant) &&
                    help.IndexOf(variant, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    missing.Add(variant);
                }
            }

            if (missing.Count == 0) return;

            if (!string.IsNullOrWhiteSpace(help))
                help = help.TrimEnd() + Environment.NewLine + Environment.NewLine;
            help += "Right-click variants:" + Environment.NewLine +
                    string.Join(
                        Environment.NewLine,
                        missing.ConvertAll(variant => "- " + variant));
            record.DetailedHelpText = help;
        }

        private static void Add(FeatureCategory category, params string[] ids)
        {
            foreach (var id in ids)
                Categories.Add(id, category);
        }
    }
}
