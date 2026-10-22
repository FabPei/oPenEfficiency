using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI
{
    public struct FeatureDisplayInfo
    {
        public string Name;
        public string Tooltip;
        public string IconData;
        public string Color;
        public string Description;
        public string DetailedHelpText;
        public string HelpImagePath;
        public FeatureCategory Category;
        public IReadOnlyList<string> SearchAliases;
        public FeatureMaturity Maturity;
        public bool UserVisible;
        public FeatureDefaultPlacement DefaultPlacement;
        public bool SupportsUserShortcut;
        public IReadOnlyList<string> Variants;
        public int MinSelection { get; set; }
        public int MaxSelection { get; set; }
        public PowerPoint.PpSelectionType RequiredType { get; set; }
        public bool RequiresTable { get; set; }
        public bool RequiresHiddenShapes { get; set; }
        public bool IsToggle { get; set; }
    }

    public static class FeatureLibrary
    {
        private static IReadOnlyList<FeatureCatalogRecord> _catalog;
        private static IReadOnlyList<FeaturePresetDefinition> _presetDefinitions;
        public static IReadOnlyList<FeatureCatalogRecord> Catalog => _catalog;
        public static IReadOnlyList<FeaturePresetDefinition> PresetDefinitions => _presetDefinitions;

        public static List<SidebarFeature> AllFeatures = new List<SidebarFeature>
        {
            new SidebarFeature { Id = "BtnSearchBar", Name = "Command Search" },
            new SidebarFeature { Id = "BtnInsertExcelValue", Name = "Insert Excel Value" },
            new SidebarFeature { Id = "BtnConnectedTextLabel", Name = "Insert Linked Text Label" },
            new SidebarFeature { Id = "BtnSizePanel", Name = "Shape Size Panel" },
            new SidebarFeature { Id = "BtnMagicResizer", Name = "Scale Selected Shapes" },
            new SidebarFeature { Id = "BtnMatchSize", Name = "Match Size" },
            new SidebarFeature { Id = "BtnMatchWidth", Name = "Match Width" },
            new SidebarFeature { Id = "BtnMatchHeight", Name = "Match Height" },
            new SidebarFeature { Id = "BtnLockDimensions", Name = "Lock/Unlock Aspect Ratio" },
            new SidebarFeature { Id = "BtnShapeLock", Name = "Lock/Unlock Shapes" },
            new SidebarFeature { Id = "BtnSlideSize", Name = "Analyze Slide Media Size" },
            new SidebarFeature { Id = "BtnAlignLeft", Name = "Align Left" },
            new SidebarFeature { Id = "BtnAlignCenter", Name = "Align Center" },
            new SidebarFeature { Id = "BtnAlignRight", Name = "Align Right" },
            new SidebarFeature { Id = "BtnAlignTop", Name = "Align Top" },
            new SidebarFeature { Id = "BtnAlignCenterVertical", Name = "Align Center Vertical" },
            new SidebarFeature { Id = "BtnAlignBottom", Name = "Align Bottom" },
            new SidebarFeature { Id = "BtnDiagonalAlign", Name = "Align Diagonally" },
            new SidebarFeature { Id = "BtnDistributeHorizontal", Name = "Distribute Horizontally" },
            new SidebarFeature { Id = "BtnDistributeVertical", Name = "Distribute Vertically" },
            new SidebarFeature { Id = "BtnRemoveHorizontalGaps", Name = "Remove Horizontal Gaps" },
            new SidebarFeature { Id = "BtnRemoveVerticalGaps", Name = "Remove Vertical Gaps" },
            new SidebarFeature { Id = "BtnIncreaseHorizontalSpacing", Name = "Increase Horizontal Spacing" },
            new SidebarFeature { Id = "BtnDecreaseHorizontalSpacing", Name = "Decrease Horizontal Spacing" },
            new SidebarFeature { Id = "BtnIncreaseVerticalSpacing", Name = "Increase Vertical Spacing" },
            new SidebarFeature { Id = "BtnDecreaseVerticalSpacing", Name = "Decrease Vertical Spacing" },
            new SidebarFeature { Id = "BtnAdjustSpacing", Name = "Adjust Spacing" },
            new SidebarFeature { Id = "BtnStretchLeft", Name = "Stretch Left" },
            new SidebarFeature { Id = "BtnStretchRight", Name = "Stretch Right" },
            new SidebarFeature { Id = "BtnStretchTop", Name = "Stretch Top" },
            new SidebarFeature { Id = "BtnStretchBottom", Name = "Stretch Bottom" },
            new SidebarFeature { Id = "BtnStretchLeftEdge", Name = "Stretch Left Edge" },
            new SidebarFeature { Id = "BtnStretchRightEdge", Name = "Stretch Right Edge" },
            new SidebarFeature { Id = "BtnStretchTopEdge", Name = "Stretch Top Edge" },
            new SidebarFeature { Id = "BtnStretchBottomEdge", Name = "Stretch Bottom Edge" },
            new SidebarFeature { Id = "BtnDockLeft", Name = "Stack Left" },
            new SidebarFeature { Id = "BtnDockRight", Name = "Stack Right" },
            new SidebarFeature { Id = "BtnDockTop", Name = "Stack Above" },
            new SidebarFeature { Id = "BtnDockBottom", Name = "Stack Below" },
            new SidebarFeature { Id = "BtnSwapPositions", Name = "Swap Positions" },
            new SidebarFeature { Id = "BtnCopyXY", Name = "Copy XY-Coordinates" },
            new SidebarFeature { Id = "BtnPasteXY", Name = "Paste XY-Coordinates" },
            new SidebarFeature { Id = "BtnSelectSameType", Name = "Select Same Type" },
            new SidebarFeature { Id = "BtnSyncObjects", Name = "Sync Shapes Across Slides" },
            new SidebarFeature { Id = "BtnArrangeCircle", Name = "Arrange Shapes in Circle" },
            new SidebarFeature { Id = "BtnArrangePro", Name = "Arrange Shapes Geometrically" },
            new SidebarFeature { Id = "BtnRectifyRotation", Name = "Snap Rotation to 90°" },
            new SidebarFeature { Id = "BtnAlignShapeAdjustments", Name = "Copy Shape Adjustments" },
            new SidebarFeature { Id = "BtnHideSelected", Name = "Hide Selected" },
            new SidebarFeature { Id = "BtnShowHidden", Name = "Show Hidden" },
            new SidebarFeature { Id = "BtnHideMasterObjects", Name = "Hide/Show Master Objects" },
            new SidebarFeature { Id = "BtnCreateMotionPath", Name = "Create Motion Path" },
            new SidebarFeature { Id = "BtnBringToFront", Name = "Bring to Front" },
            new SidebarFeature { Id = "BtnBringForward", Name = "Bring Forward" },
            new SidebarFeature { Id = "BtnSendBackward", Name = "Send Backward" },
            new SidebarFeature { Id = "BtnSendToBack", Name = "Send to Back" },
            new SidebarFeature { Id = "BtnSnapShape", Name = "Snap Shape to Object" },
            new SidebarFeature { Id = "BtnSnapToGrid", Name = "Snap to Grid" },
            new SidebarFeature { Id = "BtnSnapToObjects", Name = "Snap to Objects" },
            new SidebarFeature { Id = "BtnFontPanel", Name = "Font Panel" },
            new SidebarFeature { Id = "BtnApplyTextTool", Name = "Replace Text" },
            new SidebarFeature { Id = "BtnFitFormToText", Name = "Fit Shape to Text" },
            new SidebarFeature { Id = "BtnTranslate", Name = "Translate Text" },
            new SidebarFeature { Id = "BtnReplaceFont", Name = "Replace Formatting" },
            new SidebarFeature { Id = "BtnParagraphDialog", Name = "Paragraph & Line Spacing" },
            new SidebarFeature { Id = "BtnSpecialChars", Name = "Insert Special Characters" },
            new SidebarFeature { Id = "BtnFormatNumbers", Name = "Format Numbers" },
            new SidebarFeature { Id = "BtnAlignText", Name = "Align Text" },
            new SidebarFeature { Id = "BtnTextDirection", Name = "Text Direction" },
            new SidebarFeature { Id = "BtnChangeCase", Name = "Change Case" },
            new SidebarFeature { Id = "BtnCharacterSpacing", Name = "Character Spacing" },
            new SidebarFeature { Id = "BtnSpellCheck", Name = "Set Spell Check Language" },
            new SidebarFeature { Id = "BtnDeleteText", Name = "Delete Text" },
            new SidebarFeature { Id = "BtnTextFormatPainter", Name = "Text Format Painter" },
            new SidebarFeature { Id = "BtnSwapTextFormatted", Name = "Swap Text (Formatted)" },
            new SidebarFeature { Id = "BtnSwapTextPlain", Name = "Swap Text (Plain)" },
            new SidebarFeature { Id = "BtnPickColor", Name = "Color Picker" },
            new SidebarFeature { Id = "BtnPickFillColor", Name = "Pick Fill Color" },
            new SidebarFeature { Id = "BtnPickLineColor", Name = "Pick Line Color" },
            new SidebarFeature { Id = "BtnPickTextColor", Name = "Pick Text Color" },
            new SidebarFeature { Id = "BtnIconBrowser", Name = "Icon Library" },
            new SidebarFeature { Id = "BtnIllustrationLibrary", Name = "Illustration Library" },
            new SidebarFeature { Id = "BtnThemeColor", Name = "Theme Color Palette (RGB)" },
            new SidebarFeature { Id = "BtnTransparentColor", Name = "Color Eraser" },
            new SidebarFeature { Id = "BtnFormatShapeDialog", Name = "Open Format Shape Pane" },
            new SidebarFeature { Id = "BtnOptimizeFreeForm", Name = "Simplify Freeform Nodes" },
            new SidebarFeature { Id = "BtnStyleCheck", Name = "Style Check" },
            new SidebarFeature { Id = "BtnCleaner", Name = "Presentation Cleaner" },
            new SidebarFeature { Id = "BtnFormatBold", Name = "Bold" },
            new SidebarFeature { Id = "BtnFormatItalic", Name = "Italic" },
            new SidebarFeature { Id = "BtnFormatUnderline", Name = "Underline" },
            new SidebarFeature { Id = "BtnFormatStrikethrough", Name = "Strikethrough" },
            new SidebarFeature { Id = "BtnFormatSuperscript", Name = "Superscript" },
            new SidebarFeature { Id = "BtnFormatSubscript", Name = "Subscript" },
            new SidebarFeature { Id = "BtnInsertColumnLeft", Name = "Insert Column Left" },
            new SidebarFeature { Id = "BtnInsertColumnRight", Name = "Insert Column Right" },
            new SidebarFeature { Id = "BtnInsertRowTop", Name = "Insert Row Above" },
            new SidebarFeature { Id = "BtnInsertRowBottom", Name = "Insert Row Below" },
            new SidebarFeature { Id = "BtnTableTranspose", Name = "Transpose Table" },
            new SidebarFeature { Id = "BtnTableSplit", Name = "Split Table" },
            new SidebarFeature { Id = "BtnConvertTableToShapes", Name = "Convert Table to Shapes" },
            new SidebarFeature { Id = "BtnTableDimensions", Name = "Copy/Paste Table Dimensions" },
            new SidebarFeature { Id = "BtnTableFormatPainter", Name = "Table Format Painter" },
            new SidebarFeature { Id = "BtnTableSort", Name = "Sort Table Cells" },
            new SidebarFeature { Id = "BtnMoveTableRowColumn", Name = "Move Table Rows or Columns" },
            new SidebarFeature { Id = "BtnTableSum", Name = "Calculate Table Totals" },
            new SidebarFeature { Id = "BtnTableBranding", Name = "Table Style Library" },
            new SidebarFeature { Id = "BtnAlignToTableCell", Name = "Align to Cell" },
            new SidebarFeature { Id = "BtnIllustrativeSticker", Name = "Status Sticker" },
            new SidebarFeature { Id = "BtnAddSticker", Name = "Add Corner Sticker" },
            new SidebarFeature { Id = "BtnHarveyBall", Name = "Harvey Ball" },
            new SidebarFeature { Id = "BtnTrafficLight", Name = "Traffic Light" },
            new SidebarFeature { Id = "BtnThermometer", Name = "Thermometer" },
            new SidebarFeature { Id = "BtnStarRating", Name = "Star Rating" },
            new SidebarFeature { Id = "BtnCheckbox", Name = "Checkbox" },
            new SidebarFeature { Id = "BtnNumeration", Name = "Numbered Marker" },
            new SidebarFeature { Id = "BtnProgressSeries", Name = "Progress Series" },
            new SidebarFeature { Id = "BtnRepeatShape", Name = "Repeat Shape" },
            new SidebarFeature { Id = "BtnSplitShape", Name = "Split Shape into Grid" },
            new SidebarFeature { Id = "BtnSetMargin", Name = "Paste Text Margins" },
            new SidebarFeature { Id = "BtnLearnMargin", Name = "Copy Text Margins" },
            new SidebarFeature { Id = "BtnSplitByParagraphs", Name = "Split Text Box by Paragraphs" },
            new SidebarFeature { Id = "BtnAgendaWizard", Name = "Agenda Creator" },
            new SidebarFeature { Id = "BtnSaveAgendaLayout", Name = "Agenda Layout Designer" },
            new SidebarFeature { Id = "BtnStoryline", Name = "Storyline" },
            new SidebarFeature { Id = "BtnSlideGuidelines", Name = "Create Slide Guide Grid" },
            new SidebarFeature { Id = "BtnConsolidateMasters", Name = "Consolidate Slide Masters" },
            new SidebarFeature { Id = "BtnAssetManager", Name = "Asset Library" },
            new SidebarFeature { Id = "BtnWinnerPicker", Name = "Random Winner Picker" },
            new SidebarFeature { Id = "BtnChess", Name = "Chess" },
            new SidebarFeature { Id = "Btn2048", Name = "2048" },
            new SidebarFeature { Id = "BtnMinesweeper", Name = "Minesweeper" },
            new SidebarFeature { Id = "BtnUpdateExcelCharts", Name = "Refresh Linked Excel Charts" },
            new SidebarFeature { Id = "BtnExcelLinkManager", Name = "Excel Link Manager" },
            new SidebarFeature { Id = "BtnRefreshLinkedTable", Name = "Refresh Linked Table" },
            new SidebarFeature { Id = "BtnQRCode", Name = "QR Code" },
            new SidebarFeature { Id = "BtnExportWizard", Name = "Export Presentation" },
            new SidebarFeature { Id = "BtnGanttChart", Name = "Gantt & Timeline Editor" },
            new SidebarFeature { Id = "BtnSettings", Name = "Settings" },
            new SidebarFeature { Id = "BtnTemplateManager", Name = "Template Library" },
            new SidebarFeature { Id = "BtnStickyNote", Name = "Sticky Note" },
            new SidebarFeature { Id = "BtnSlideNotes", Name = "Private Review Notes" },
            new SidebarFeature { Id = "BtnSlidePaste", Name = "Insert Slide Snapshots" },
            new SidebarFeature { Id = "BtnHtmlSlideImport", Name = "Import HTML Slides" }
        };

        static FeatureLibrary()
        {
            var featureTypes = GetLoadableTypes(Assembly.GetExecutingAssembly())
                .Select(type => new
                {
                    Type = type,
                    Metadata = type.GetCustomAttribute<FeatureMetadataAttribute>()
                })
                .Where(item => item.Metadata != null)
                .OrderBy(item => item.Metadata.Id, StringComparer.Ordinal)
                .ToList();
            var recordsById = new Dictionary<string, FeatureCatalogRecord>(StringComparer.Ordinal);
            var orderedRecords = new List<FeatureCatalogRecord>();

            // Preserve the historical command order while resolving the duplicate
            // name-only registrations into authoritative catalog records.
            foreach (var seed in AllFeatures)
            {
                var attributed = featureTypes.FirstOrDefault(item =>
                    string.Equals(item.Metadata.Id, seed.Id, StringComparison.Ordinal));
                var record = attributed != null
                    ? CreateAttributedRecord(attributed.Metadata)
                    : CreateManualRecord(seed);
                if (!recordsById.ContainsKey(record.Id))
                {
                    recordsById.Add(record.Id, record);
                    orderedRecords.Add(record);
                }
            }

            foreach (var attributed in featureTypes)
            {
                if (recordsById.ContainsKey(attributed.Metadata.Id)) continue;
                var record = CreateAttributedRecord(attributed.Metadata);
                recordsById.Add(record.Id, record);
                orderedRecords.Add(record);
            }

            // Feature Explorer remains a persistent sidebar action in this slice; adding
            // its catalog record does not move it into a preset.
            if (!recordsById.ContainsKey("BtnExploreFeatures"))
            {
                var explorer = CreateManualRecord(
                    new SidebarFeature { Id = "BtnExploreFeatures", Name = "Feature Explorer" });
                recordsById.Add(explorer.Id, explorer);
                orderedRecords.Add(explorer);
            }

            _catalog = orderedRecords.AsReadOnly();
            AllFeatures = orderedRecords
                .Where(record => record.UserVisible && !record.IsObsolete)
                .Select(record => new SidebarFeature { Id = record.Id, Name = record.Name })
                .ToList();

            _presetDefinitions = SidebarPresetFactory.CreateDefinitions(_catalog);
            var presetErrors = FeatureCatalogValidator.ValidatePresets(_catalog, _presetDefinitions);
            if (presetErrors.Count > 0)
            {
                throw new InvalidOperationException(
                    "Sidebar preset validation failed:" + Environment.NewLine +
                    string.Join(Environment.NewLine, presetErrors));
            }
        }

        public static SidebarConfig GetEmergencyDefault() => GetConsultingStandardLayout();

        public static SidebarConfig GetConsultingStandardLayout() =>
            SidebarPresetFactory.Create(SidebarPresetFactory.ConsultingStandardName, _catalog);

        public static SidebarConfig GetAllFeaturesLayout() =>
            SidebarPresetFactory.Create(SidebarPresetFactory.AllFeaturesName, _catalog);

        public static SidebarConfig GetMinimalLayout() =>
            SidebarPresetFactory.Create(SidebarPresetFactory.MinimalName, _catalog);

        public static SidebarConfig GetLabsLayout() =>
            SidebarPresetFactory.Create(SidebarPresetFactory.LabsName, _catalog);

        private static IEnumerable<Type> GetLoadableTypes(Assembly assembly)
        {
            try
            {
                return assembly.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                // The Windows Runtime types used by local OCR can be unavailable to the
                // legacy VSTO reflection loader. Keep the sidebar usable for all other types.
                return ex.Types.Where(type => type != null);
            }
        }

        public static FeatureDisplayInfo GetFeatureInfo(string id)
        {
            var record = _catalog.FirstOrDefault(item =>
                string.Equals(item.Id, id, StringComparison.Ordinal));
            return record == null ? new FeatureDisplayInfo() : ToDisplayInfo(record);
        }

        private static FeatureDisplayInfo GetLegacyFeatureInfo(string id)
        {
            var autoInfo = FeatureDiscovery.GetFeatureInfo(id);
            if (autoInfo != null) return autoInfo.Value;

            switch (id)
            {
                case "BtnTableFormatPainter": return new FeatureDisplayInfo { Tooltip = "Table Format Painter", Color = "#06B6D4", Description = "Paints fill or basic text formatting from one selected table cell along chosen row and column directions.", RequiresTable = true, DetailedHelpText = "### Table Format Painter\n\nSelect exactly one source cell, choose one or more directions, select Text Formatting or Cell Fill, and optionally set how many cells to skip between painted targets. Borders, cell text, rich-text runs, links, gradients, patterns, and effects are not copied." };
                case "BtnMoveTableRowColumn": return new FeatureDisplayInfo { Tooltip = "Reorder selected table rows or columns", Color = "#3B82F6", IconData = "M3,4H21V20H3V4M6,7H18M6,12H18M6,17H18M9,4V20M15,4V20M11,9L8,12L11,15M13,9L16,12L13,15", Description = "Moves, swaps, or duplicates selected table rows and columns, including contiguous blocks.", MinSelection = 0, RequiresTable = true, DetailedHelpText = "### Move Table Rows or Columns\n\nSelect one or more cells, or place the text cursor inside a cell. Move the corresponding row or column by one step, to an edge, or to a numbered position. Contiguous selections move as one block. **Keep table design** moves formatted content while retaining destination fills, borders, dimensions, and cell layout. Other modes can move all formatting, retain only destination borders, or move plain text only. The context menu also supports swapping two selected rows or columns, duplicating a selected block, repeating the last move, and protecting the first or last row." };
                case "BtnSearchBar": return new FeatureDisplayInfo { Tooltip = "Command Search", SearchAliases = new[] { "Search Command Palette", "Search features" }, Color = "#10B981", Description = "Find and execute any openEfficiency command by canonical or legacy name.", RequiredType = PowerPoint.PpSelectionType.ppSelectionNone, DetailedHelpText = "### Command Search\nType a command name, legacy name, or stable ID to filter the Sidebar. Press Enter to run the best match, or Escape to clear the search and return to the configured layout. Right-click the widget to focus the search field or clear the current query." };
                case "BtnAdjustSpacing": return new FeatureDisplayInfo { SearchAliases = new[] { "Advanced Spacing" }, Tooltip = "Adjust Spacing", Color = "#10B981", Description = "Opens anchored and center-preserving horizontal and vertical spacing actions for the selected shapes.", MinSelection = 2, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Adjust Spacing\nSelect at least two shapes, click the feature, then choose whether to increase or decrease horizontal or vertical spacing. The centered decrease actions move shapes equally toward the selection midpoint instead of anchoring the first spatial shape. The configured spacing increment is applied between each neighboring shape. A keyboard shortcut assigned to this command directly increases horizontal spacing." };
                case "BtnUpdateExcelCharts": return new FeatureDisplayInfo { SearchAliases = new[] { "Update Excel Data", "Update Excel Charts" }, Tooltip = "Refresh Linked Excel Charts", Color = "#F43F5E", IconData = "M3,3H21V21H3V3M5,5V19H19V5H5M7,7H10V10H7V7M11,7H14V10H11V7M15,7H18V10H15V7M7,11H10V14H7V11M11,11H14V14H11V11M15,11H18V14H15V11M7,15H10V18H7V15M11,15H14V18H11V15M15,15H18V18H15V15M19.14,17.88L20.5,19.25L18.31,21.44L16.94,20.08L19.14,17.88Z", Description = "Updates externally linked Excel charts on the current slide or across the presentation.", DetailedHelpText = "### Refresh Linked Excel Charts\nClick the feature and choose whether to update the current slide or the complete presentation. Only externally linked Excel charts are refreshed; embedded charts are left unchanged. A keyboard shortcut updates the current slide directly." };
                case "BtnApplyTextTool": return new FeatureDisplayInfo { SearchAliases = new[] { "Apply Text", "Apply Text to Selection", "Apply Text from First Selected", "Apply Text from Last Selected" }, Tooltip = "Replace Text", Color = "#FDE047", Description = "Opens a compact editor that replaces the active text selection, selected text shapes, or selected table cells.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Replace Text\n\nEnter replacement text and apply it to the active text selection, every selected text-bearing shape, or the selected cells of a table. Selecting the table itself applies the text to every cell. **Take text from Master** uses the configured first- or last-selected master shape and is available only when that master can supply text. Right-click can use the first or last selected shape explicitly. Text replacement creates one PowerPoint Undo entry; it does not copy font or paragraph formatting." };
                case "BtnArrangeCircle": return new FeatureDisplayInfo { SearchAliases = new[] { "Arrange in Circle" }, Tooltip = "Arrange Shapes in Circle", Color = "#FBBF24", Description = "Arranges selected shapes around a circular path. Right-click to choose another enclosing shape.", MinSelection = 2, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Arrange Shapes in Circle\nSelect at least two shapes and click the feature to distribute them around a circle. Right-click to arrange the selection in a triangle, square, star, octagon, or pentagon instead." };
                case "BtnIncreaseHorizontalSpacing": return new FeatureDisplayInfo { SearchAliases = new[] { "Incr Horiz Spacing" }, Tooltip = "Increase Horizontal Spacing", Color = "#10B981", Description = "Increases the horizontal gap between selected shapes by the configured spacing increment.", MinSelection = 2, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Increase Horizontal Spacing\nSelect two or more shapes, then increase the gaps between them horizontally. Right-click any incremental spacing button to increase or decrease horizontal or vertical spacing directly. Configure the increment in Settings to match your layout system." };
                case "BtnDecreaseHorizontalSpacing": return new FeatureDisplayInfo { SearchAliases = new[] { "Decr Horiz Spacing" }, Tooltip = "Decrease Horizontal Spacing", Color = "#10B981", Description = "Decreases the horizontal gap between selected shapes by the configured spacing increment.", MinSelection = 2, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Decrease Horizontal Spacing\nSelect two or more shapes, then reduce the gaps between them horizontally. Right-click any incremental spacing button to increase or decrease horizontal or vertical spacing directly. Configure the increment in Settings to match your layout system." };
                case "BtnIncreaseVerticalSpacing": return new FeatureDisplayInfo { SearchAliases = new[] { "Incr Vert Spacing" }, Tooltip = "Increase Vertical Spacing", Color = "#10B981", Description = "Increases the vertical gap between selected shapes by the configured spacing increment.", MinSelection = 2, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Increase Vertical Spacing\nSelect two or more shapes, then increase the gaps between them vertically. Right-click any incremental spacing button to increase or decrease horizontal or vertical spacing directly. Configure the increment in Settings to match your layout system." };
                case "BtnDecreaseVerticalSpacing": return new FeatureDisplayInfo { SearchAliases = new[] { "Decr Vert Spacing" }, Tooltip = "Decrease Vertical Spacing", Color = "#10B981", Description = "Decreases the vertical gap between selected shapes by the configured spacing increment.", MinSelection = 2, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Decrease Vertical Spacing\nSelect two or more shapes, then reduce the gaps between them vertically. Right-click any incremental spacing button to increase or decrease horizontal or vertical spacing directly. Configure the increment in Settings to match your layout system." };
                case "BtnFontPanel": return new FeatureDisplayInfo { SearchAliases = new[] { "Quick Font Panel" }, Tooltip = "Edit font family and size for the current selection", Color = "#FDE047", Description = "Shows embedded controls for changing the font family and size of the active text selection or selected text-bearing shapes.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Font Panel\n\nChoose or type a font family and font size for the active text selection or all selected text-bearing shapes. Mixed selections show an empty field. Enter and focus loss apply a typed value only once; successful changes create one PowerPoint Undo entry. Right-click to focus and select either field directly." };
                case "BtnSnapShape": return new FeatureDisplayInfo { Tooltip = "Snap Shape to Object", Color = "#FBBF24", Description = "Moves the first of two selected shapes flush to the nearest side of the second shape.", MinSelection = 2, MaxSelection = 2, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Snap Shape to Object\nSelect exactly two shapes. The first snaps flush to the nearest side of the second; right-click for an explicit direction." };
                case "BtnTableSort": return new FeatureDisplayInfo { SearchAliases = new[] { "Table Sorting" }, Tooltip = "Sort Table Cells", Color = "#3B82F6", Description = "Opens ascending and descending sorting for selected cells or complete row texts.", RequiresTable = true, DetailedHelpText = "### Sort Table Cells\n\nSelect at least two vertical cells in the sort column, then choose A-Z or Z-A. Selected-cell modes reorder only that column's plain text. Whole-row modes reorder every cell text in the selected row range using the leftmost selected column as key; destination formatting stays in its row. Numeric values sort numerically, text uses culture-aware case-insensitive comparison, and failures attempt to restore all affected texts." };
                case "BtnTranslate": return new FeatureDisplayInfo { Tooltip = "Translate Text", Color = "#FDE047", Description = "Opens PowerPoint's native translator or translates selected text through a configured external provider.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Translate Text\n\nClick to open PowerPoint's native translation interface. Use the context menu to send the active text selection or selected text-bearing shapes to the configured DeepL, OpenAI, Claude, LibreTranslate, or free MyMemory service. External providers use HTTPS; a custom LibreTranslate endpoint may use HTTP only for a loopback server such as localhost. API translation replaces plain text and reapplies the common base font properties; mixed rich-text runs cannot be preserved exactly. The result is applied only if the source presentation is still active and creates one PowerPoint Undo entry." };
                case "BtnExportWizard": return new FeatureDisplayInfo { SearchAliases = new[] { "Export Wizard" }, Tooltip = "Export selected or all slides to common delivery formats", Color = "#10B981", Description = "Exports all or selected slides as PPTX, PDF, both formats, or editable HTML, with optional cleanup and delivery actions.", MinSelection = 0, IsToggle = true, DetailedHelpText = "### Export Presentation\nChoose the entire presentation or selected slides, optionally exclude hidden slides, and export to PPTX, PDF, both formats, or editable HTML. Right-click opens the export window with a common format and scope combination already selected; export never starts until you confirm it. Cleanup applies only to exported copies; the open presentation is changed only when **Also apply these properties to the open presentation** is explicitly enabled. Existing output files are listed for confirmation before replacement. Clipboard and e-mail delivery use temporary files, while disk export opens the destination folder when complete." };
                case "BtnTemplateManager": return new FeatureDisplayInfo { SearchAliases = new[] { "PowerPoint Template Manager" }, Tooltip = "Find and open reusable PowerPoint templates", Color = "#3B82F6", Description = "Find and open PowerPoint templates from default locations, watched folders, and explicitly added files.", RequiredType = PowerPoint.PpSelectionType.ppSelectionNone, DetailedHelpText = "### Template Library\nSearch default Windows and Office template locations, add watched folders, or register individual POTX, POT, PPTX, and PPTM files. Right-click opens search, file registration, watched-folder management, or hidden-template restoration directly. Open creates an untitled presentation so the source template remains unchanged. Hide removes an entry from the library view without deleting its file. Permanent deletion is available only for files explicitly added to the library and always requires confirmation." };
                case "BtnParagraphDialog": return new FeatureDisplayInfo { Tooltip = "Paragraph & Line Spacing", Color = "#FDE047", Description = "Opens PowerPoint's native paragraph dialog or applies common spacing presets.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Paragraph & Line Spacing\n\nSelect text or one or more text-bearing shapes, then open PowerPoint's native paragraph dialog to adjust alignment, indentation, line spacing, and paragraph spacing. Right-click applies single, 1.15, or 1.5 line spacing immediately, or removes paragraph spacing. Each preset preserves unrelated paragraph properties and creates one Undo entry." };
                case "BtnSpecialChars": return new FeatureDisplayInfo { SearchAliases = new[] { "Insert Special Chars" }, Tooltip = "Insert Special Characters", Color = "#FDE047", Description = "Opens a persistent picker for inserting symbols, currencies, operators, punctuation, and language-specific characters.", MinSelection = 1, MaxSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Insert Special Characters\n\nPlace the cursor in text or select exactly one text-bearing shape, then choose a category and click a character. It is inserted after the active cursor or selection, or appended to the selected shape. The context menu opens the picker directly at a symbol or language category; each insertion creates one PowerPoint Undo entry." };
                case "BtnFlightMode": return new FeatureDisplayInfo { Tooltip = "Master Covers & Overlays", Color = "#D946EF", Description = "Quickly obscures or hides sensitive slide content, logos, and master backgrounds for clean presentation sharing and previews.", RequiredType = PowerPoint.PpSelectionType.ppSelectionNone, DetailedHelpText = "Quickly obscures sensitive content and branding elements." };
                case "BtnSizePanel": return new FeatureDisplayInfo { Tooltip = "Shape Size Panel (Width × Height)", Color = "#6366F1", Description = "Displays and edits exact width and height for the selected shapes using centimeter or point input.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Shape Size Panel\n\nThe embedded sidebar fields show a shared width or height for the selected shapes and show **Mixed** when that dimension differs. Enter a positive value with `cm`, with `pt`, or without a suffix for centimeters to apply one exact dimension to every selected shape. The arrow buttons adjust each shape by 0.1 cm. Existing aspect-ratio locks remain active, Enter and focus loss do not apply the same edit twice, and each change creates one Undo entry with rollback on errors. Right-click to focus and select the width or height field directly.", IsToggle = true };
                case "BtnStoryline": return new FeatureDisplayInfo { Tooltip = "Review the presentation storyline", Color = "#FDE047", Description = "Reviews, filters, edits, and exports section-aware slide headlines with consulting-focused quality checks.", RequiredType = PowerPoint.PpSelectionType.ppSelectionNone, DetailedHelpText = "### Storyline\nOpen a section-aware narrative view of the presentation. Search or filter headlines, navigate through stable slide IDs, edit detected title shapes, assign review status, and use the context menu for slide actions. The quality check flags missing, duplicate, generic, overly long, question-only, and weak topic headlines and provides local action-title suggestions without sending presentation content to an external service. Copy or export the complete deck, the current section, headlines only, or issues only." };
                case "BtnStickyNote": return new FeatureDisplayInfo { Tooltip = "Sticky Note", Color = "#10B981", Description = "Inserts a tagged sticky note for internal comments using the configured default style or a right-click style preset.", RequiredType = PowerPoint.PpSelectionType.ppSelectionNone, DetailedHelpText = "### Sticky Note\n\nLeft-click inserts the configured default note. Right-click inserts a Theme, Consulting, or Compact Banner note directly without changing the saved default. Each style keeps its own saved background and font colors; Compact Banner includes its offset shadow block. Sticky Note Manager handles moving, restoring, converting, and deleting notes." };
                case "BtnIllustrativeSticker": return new FeatureDisplayInfo { Tooltip = "Insert or manage slide status stickers", Color = "#F43F5E", IconData = "M12,16L12.01,16 M12,8L12,12 M15.31,2L20,6.69L20,13.31L15.31,18L8.69,18L4,13.31L4,8.69L8.69,4Z", Description = "Adds and manages consulting status stickers across one or many slides.", MinSelection = 0, DetailedHelpText = "### Status Sticker\nInsert presets or custom labels, combine workflow, confidentiality and audience states, update or remove stickers in bulk, choose a corner, use the presentation theme color, and add date or author metadata from the right-click menu." };
                case "BtnAssetManager": return new FeatureDisplayInfo { Tooltip = "Asset Library (Slides, Images, Shapes)", Color = "#8B5CF6", IconData = "M21.54,15L17,15L17,19.54 M7,3.34L7,5C7,6.1 7.9,7 9,7C9,5.9 9.9,5 11,5L14.17,5 M11,21.95L11,18L11,17L2.05,17 M12,2C17.522847498307936,2.0 22.0,6.477152501692064 22.0,12.0C22.0,17.522847498307936 17.522847498307936,22.0 12.0,22.0C6.477152501692064,22.0 2.0,17.522847498307936 2.0,12.0C2.0,6.477152501692064 6.477152501692064,2.0 12.0,2.0", Description = "Opens a searchable local library for reusable slides, shapes, images, folders, tags, and favorites.", MinSelection = 0, DetailedHelpText = "### Asset Library\nUse Slides, Images, Shapes, and My Assets to browse or search reusable local content. Right-click opens a specific library section or starts Save Current Selection directly. Save the current slide or shape selection into a chosen Favorites folder, optionally with tags and automatic tagging. PowerPoint export and import work remains on PowerPoint's UI thread; generated files use unique names, and rename or deletion operations are restricted to the Asset Library root. Permanent deletion requires confirmation in the relevant library tab.", IsToggle = true };
                case "BtnCopyXY": return new FeatureDisplayInfo { Tooltip = "Copy the selected shape's X/Y position", Color = "#F59E0B", IconData = "M6,20L10,20 M14,10L18,10 M6,14L8,14L8,20 M14,4L16,4L16,10 M14,16L14,18Q14,20 16,20L16,20Q18,20 18,18L18,16Q18,14 16,14L16,14Q14,14 14,16Z M6,6L6,8Q6,10 8,10L8,10Q10,10 10,8L10,6Q10,4 8,4L8,4Q6,4 6,6Z", Description = "Copies the exact top-left X/Y position from one selected shape for Paste XY-Coordinates.", MinSelection = 1, MaxSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Copy XY-Coordinates\nSelect exactly one shape, then copy its Left and Top position as a versioned openEfficiency X/Y payload for Paste XY-Coordinates. This command never copies Width or Height. Paste XY-Coordinates also accepts legacy two- or four-value X/Y payloads, but right-click Copy XY-Coordinates clears only a valid current openEfficiency payload and never unrelated numeric clipboard text. Use Copy/Paste Position & Size when you need to transfer dimensions." };
                case "BtnPasteXY": return new FeatureDisplayInfo { Tooltip = "Apply copied X/Y position to the selected shape", Color = "#F59E0B", Description = "Applies a copied top-left X/Y position to one selected shape without changing its size.", MinSelection = 1, MaxSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Paste XY-Coordinates\nSelect exactly one target shape, then paste the Left and Top position copied by Copy XY-Coordinates. Width, Height, rotation, text, and formatting stay unchanged. Current payloads are versioned and openEfficiency-owned; legacy two- and four-value X/Y payloads remain paste-compatible, with four-value Width and Height ignored. Right-click Paste XY-Coordinates clears only a valid current openEfficiency payload, never unrelated numeric clipboard text. Use Copy/Paste Position & Size when you need to transfer dimensions." };
                case "BtnConsolidateMasters": return new FeatureDisplayInfo { Tooltip = "Consolidate Masters", Color = "#10B981", IconData = "M2,2H11V6H9V4H4V9H6V11H2V2M22,22H13V18H15V20H20V15H18V13H22V22M8,8H16V16H8V8Z", Description = "Cleans up your presentation by merging duplicate slide masters and layouts. Reduces file size and improves template consistency.", RequiredType = PowerPoint.PpSelectionType.ppSelectionNone, DetailedHelpText = "Cleans up your presentation by merging duplicate slide masters and layouts. Reduces file size and improves template consistency." };
                case "BtnHarveyBall": return new FeatureDisplayInfo { Tooltip = "Harvey Ball", Color = "#F43F5E", IconData = "M21,12", Description = "Cycles through Harvey Ball states (0%, 25%, 50%, 75%, 100%) to represent progress or qualitative data. Right-click to set a specific value.", MinSelection = 0, DetailedHelpText = "Cycles through Harvey Ball states (0%, 25%, 50%, 75%, 100%) to represent progress or qualitative data.\n\nRight-click options:\n• Set specific percentage directly." };
                case "BtnHideSelected": return new FeatureDisplayInfo { Tooltip = "Hide Selected", Color = "#D946EF", IconData = "M10.73,5.08 M14.08,14.16 M17.48,17.5 M2,2L22,22", Description = "Hides selected shapes and marks them for safe restoration without taking ownership of already-hidden objects.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Hide Selected\nHides visible selected shapes and marks them with openEfficiency metadata. Right-click to hide the selection, safely restore openEfficiency-hidden shapes, or reveal every hidden shape after confirmation." };
                case "BtnShowHidden": return new FeatureDisplayInfo { Tooltip = "Show Hidden", Color = "#D946EF", IconData = "M2.06,12.35 M12,9C13.65685424949238,9.0 15.0,10.34314575050762 15.0,12.0C15.0,13.65685424949238 13.65685424949238,15.0 12.0,15.0C10.34314575050762,15.0 9.0,13.65685424949238 9.0,12.0C9.0,10.34314575050762 10.34314575050762,9.0 12.0,9.0", Description = "Restores shapes hidden by openEfficiency while preserving unrelated hidden shapes.", RequiresHiddenShapes = true, DetailedHelpText = "### Show Hidden\nRestores shapes hidden through Hide Selected. Right-click to choose safe openEfficiency-only restoration or the confirmed recovery action that reveals every hidden shape on the slide." };
                case "BtnCheckbox": return new FeatureDisplayInfo { Tooltip = "Checkbox", Color = "#F43F5E", IconData = "M9,12L11,14L15,10 M3,5L3,19Q3,21 5,21L19,21Q21,21 21,19L21,5Q21,3 19,3L5,3Q3,3 3,5Z", Description = "Inserts a vector-based checkmark or cross inside a box. Useful for to-do lists and project status slides.", MinSelection = 0, DetailedHelpText = "Inserts a vector-based checkmark or cross inside a box. Useful for to-do lists and project status slides." };
                case "BtnMagicResizer": return new FeatureDisplayInfo { SearchAliases = new[] { "Magic Resizer", "Scale Shapes & Spacing", "Scale Selection" }, Tooltip = "Scale selected shapes", Color = "#6366F1", IconData = "M136,112L48,112L48,200L136,200L136,120ZM128,200L56,200L56,128L128,128ZM216,184L216,200L176,200L200,200L200,184ZM216,112L216,144L216,112ZM216,56L216,72L216,56L184,56L200,56ZM152,48L112,48L144,48ZM40,80L40,56L72,56L56,56L56,80Z", Description = "Scales every selected shape by the same factor, with a shared anchor setting and optional text, line, margin, and bullet scaling.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "### Scale Selected Shapes\n\nSelect one or more shapes — a group is a valid target and scales as one unit — then scale each one by the same percentage or enter target dimensions derived from the first selected shape. Choose which point on every shape remains fixed. Font size, line weight, text margins, and bullet size can be scaled with the shapes. The command keeps each shape's chosen anchor in place rather than scaling the complete selection bounding box.", IsToggle = true };
                case "BtnDeleteText": return new FeatureDisplayInfo { Tooltip = "Delete Text", Color = "#EAB308", IconData = "M16,5L3,5 M11,12L3,12 M16,19L3,19 M15.5,9.5L20.5,14.5 M20.5,9.5L15.5,14.5", Description = "Clears text from selected text-bearing shapes, groups, or selected table cells while retaining the shapes.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "Clears selected table cells when cells are selected and clears every cell when the table itself is selected. Right-click to delete all text while keeping shapes, delete only highlighted text, delete selected paragraphs, or open Replace Selected Text. Highlight-dependent actions are disabled unless text is actively selected." };
                case "BtnSwapTextFormatted": return new FeatureDisplayInfo { Tooltip = "Swap Text (Formatted)", Color = "#EAB308", IconData = "M18,14V11H12V9H18V6L22,10L18,14M6,10V13H12V15H6V18L2,14L6,10Z", Description = "Exchanges text and its text formatting between exactly two shapes.", MinSelection = 2, MaxSelection = 2, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "The complete formatted text ranges travel between the two shapes; shape formatting and geometry remain in place. Right-click can additionally shrink overflowing text or resize both shapes to their swapped content." };
                case "BtnSwapTextPlain": return new FeatureDisplayInfo { Tooltip = "Swap plain text (right-click for more options)", Color = "#EAB308", IconData = "M21,9L17,5V8H10V10H17V13M7,11L3,15L7,19V16H14V14H7V11Z", Description = "Swaps plain text between two text shapes or selected table cells while retaining destination formatting.", MinSelection = 1, RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes, DetailedHelpText = "Click to swap two text shapes or two selected table cells. Right-click to shrink overflow, cycle text through several selected targets, or exchange chosen paragraphs." };
                case "BtnTrafficLight": return new FeatureDisplayInfo { Tooltip = "Traffic Light", Color = "#F43F5E", IconData = "M128,56ZM128,104ZM128,136ZM128,184ZM216,144L200,144L200,80L216,80L200,80L200,40L72,40L72,64L40,64L56,64L56,128L40,128L56,128L56,184L184,184L184,160L200,160ZM184,216L72,216L72,40L184,40L184,216Z", Description = "Inserts a traffic light symbol or changes a selected traffic light without resetting its appearance.", MinSelection = 0, DetailedHelpText = "### Traffic Light\n\nClick to insert or set Green. Right-click the Sidebar button to insert or set Red, Yellow, Green, or Off directly. Existing openEfficiency traffic lights retain their orientation and frame settings. Select the traffic light to change its color and appearance with the contextual controls." };
                case "BtnStarRating": return new FeatureDisplayInfo { Tooltip = "Star Rating", Color = "#F43F5E", IconData = "M11.53,2.29L13.84,6.97L19.0,7.73L15.27,11.37L16.15,16.51L11.53,14.08L6.4,21.01L7.28,15.87L2.16,9.79L7.33,9.04Z", Description = "Inserts a sequence of stars representing a rating from 1 to 5. Great for feedback or performance summaries.", MinSelection = 0, DetailedHelpText = "Inserts a sequence of stars representing a rating from 1 to 5. Great for feedback or performance summaries." };
                case "BtnExploreFeatures": return new FeatureDisplayInfo { Tooltip = "Feature Explorer", SearchAliases = new[] { "Explore Features" }, Color = "#6366F1", IconData = "M12,2A10,10 0 1,0 22,12A10,10 0 0,0 12,2M12,17A1,1 0 1,1 11,16A1,1 0 0,1 12,17M13,14H11V9H13V14Z", Description = "The Feature Explorer provides a searchable index of all commands available in openEfficiency.", DetailedHelpText = "### Feature Explorer\nSearch the openEfficiency command catalog by canonical name, legacy name, description, or stable ID, then review practical usage guidance and the current shortcut. Press Ctrl+F to focus search or Escape to close. Right-click to focus all-feature search or open common catalog categories directly." };
                case "BtnSettings": return new FeatureDisplayInfo { Tooltip = "Configure openEfficiency preferences", SearchAliases = new[] { "openEfficiency Settings", "oPenEfficiency Settings" }, Color = "#94A3B8", Description = "Configures Sidebar layout, appearance, sticky notes, translations, caching, local diagnostics, shortcuts, and portable configuration files.", DetailedHelpText = "### Settings\nEdit the Sidebar layout and appearance, sticky-note defaults, translation providers, cache folders, and shortcuts. Changes are applied only after Save & Apply succeeds. Translation API keys are stored locally using Windows user-scoped protection; existing plaintext settings are migrated automatically. Layout and configuration import skip unknown feature IDs. Portable exports omit API keys by default and require explicit confirmation before including them as plaintext. Under General, Create Diagnostic Bundle writes a local ZIP with runtime details, configured feature IDs, sanitized configuration metadata, and up to 512 KiB of the recent sanitized error log. It never uploads data and excludes presentation content, credential values, full paths, and machine/user identity; review the ZIP before sharing it. Restore All Defaults resets layout and appearance but intentionally leaves shortcuts unchanged. Right-click to open any settings section directly." };
                case "BtnGanttChart": return new FeatureDisplayInfo { SearchAliases = new[] { "Gantt Chart & Timeline Editor" }, Tooltip = "Create or edit Gantt charts and flexible timelines", Color = "#3B82F6", IconData = "M2,2H4V20H22V22H2Z M7,4H15V6H7Z M10,9H20V11H10Z M14,14H22V16H14Z", Description = "Creates rule-based Gantt charts or flexible project timelines directly in PowerPoint.", MinSelection = 0, IsToggle = true, DetailedHelpText = "### Gantt & Timeline Editor\n\nOpen the editor and click **Create New Chart** to insert a new chart with one editable starter task. Use **Update Active Chart** to rebuild the chart currently loaded in the editor. Use the editor in two complementary ways: build a rule-based Gantt chart with task hierarchies, consistent working-day durations, network-based critical-path calculation, and jointly evaluated FS/SS/FF predecessor constraints, or create a flexible Think-cell-style timeline with freely arranged lanes, events, sub-milestones, and sub-bars. You can switch between both modes; re-enabling Gantt rules recalculates summaries. Select an existing openEfficiency chart before opening the editor to continue editing that chart; multiple independent charts can coexist on one slide. Summary tasks move together with their descendants. New sibling rows are inserted after the complete subtree, deleting a summary retains and promotes its descendants, and Indent is available only when a valid parent row exists above. Each structural edit is restored as one complete Undo step. Long task names use 0.1 cm internal margins, wrap inside the fixed-width name column, and grow their lanes automatically. Select a regular task bar to show handles at both ends; drag the left handle to change its start or the right handle to change its end. You can also edit task names or schedules directly on the slide and use Sync From Slide to bring those changes back into the editor. Excel sync validates dates and preserves existing PowerPoint-only styling when rows can be matched by WBS or task name. Right-click chart elements for row, milestone, event-line, indentation, and styling actions." };
                default: return new FeatureDisplayInfo { Color = "#9CA3AF", IconData = "M12,2L14.5,9H22L16,13.5L18.5,21L12,16.5L5.5,21L8,13.5L2,9H9.5L12,2Z" };
            }
        }

        private static FeatureCatalogRecord CreateAttributedRecord(FeatureMetadataAttribute metadata)
        {
            var record = new FeatureCatalogRecord
            {
                Id = metadata.Id,
                Name = metadata.Name,
                Tooltip = metadata.Tooltip,
                Description = metadata.Description,
                DetailedHelpText = metadata.DetailedHelpText,
                IconData = metadata.IconData,
                Color = metadata.Color,
                HelpImagePath = metadata.HelpImagePath,
                Category = metadata.Category,
                SearchAliases = metadata.SearchAliases ?? Array.Empty<string>(),
                MinSelection = metadata.MinSelection,
                MaxSelection = metadata.MaxSelection,
                RequiredType = metadata.RequiredType,
                RequiresTable = metadata.RequiresTable,
                RequiresHiddenShapes = metadata.RequiresHiddenShapes,
                IsToggle = metadata.IsToggle,
                Maturity = metadata.Maturity,
                UserVisible = metadata.UserVisible,
                DefaultPlacement = metadata.DefaultPlacement,
                SupportsUserShortcut = metadata.SupportsUserShortcut,
                Variants = metadata.Variants ?? Array.Empty<string>()
            };
            FeatureCatalogPolicy.Apply(record);
            return record;
        }

        private static FeatureCatalogRecord CreateManualRecord(SidebarFeature feature)
        {
            var info = GetLegacyFeatureInfo(feature.Id);
            var record = new FeatureCatalogRecord
            {
                Id = feature.Id,
                Name = feature.Name,
                Tooltip = info.Tooltip,
                Description = info.Description,
                DetailedHelpText = info.DetailedHelpText,
                IconData = info.IconData,
                Color = info.Color,
                HelpImagePath = info.HelpImagePath,
                Category = info.Category,
                SearchAliases = info.SearchAliases ?? Array.Empty<string>(),
                MinSelection = info.MinSelection,
                MaxSelection = info.MaxSelection,
                RequiredType = ToOfficeSelectionType(info.RequiredType),
                RequiresTable = info.RequiresTable,
                RequiresHiddenShapes = info.RequiresHiddenShapes,
                IsToggle = info.IsToggle,
                Maturity = info.Maturity,
                UserVisible = true,
                DefaultPlacement = info.DefaultPlacement,
                SupportsUserShortcut = true,
                Variants = info.Variants ?? Array.Empty<string>(),
                IsManual = true
            };
            FeatureCatalogPolicy.Apply(record);
            return record;
        }

        private static FeatureDisplayInfo ToDisplayInfo(FeatureCatalogRecord record)
        {
            return new FeatureDisplayInfo
            {
                Name = record.Name,
                Tooltip = record.Tooltip,
                IconData = record.IconData,
                Color = record.Color,
                Description = record.Description,
                DetailedHelpText = record.DetailedHelpText,
                HelpImagePath = record.HelpImagePath,
                Category = record.Category,
                SearchAliases = record.SearchAliases,
                Maturity = record.Maturity,
                UserVisible = record.UserVisible,
                DefaultPlacement = record.DefaultPlacement,
                SupportsUserShortcut = record.SupportsUserShortcut,
                Variants = record.Variants,
                MinSelection = record.MinSelection,
                MaxSelection = record.MaxSelection,
                RequiredType = ToPowerPointSelectionType(record.RequiredType),
                RequiresTable = record.RequiresTable,
                RequiresHiddenShapes = record.RequiresHiddenShapes,
                IsToggle = record.IsToggle
            };
        }

        private static OfficeSelectionType ToOfficeSelectionType(PowerPoint.PpSelectionType type)
        {
            switch (type)
            {
                case PowerPoint.PpSelectionType.ppSelectionShapes: return OfficeSelectionType.Shapes;
                case PowerPoint.PpSelectionType.ppSelectionText: return OfficeSelectionType.Text;
                case PowerPoint.PpSelectionType.ppSelectionSlides: return OfficeSelectionType.Slides;
                default: return OfficeSelectionType.None;
            }
        }

        private static PowerPoint.PpSelectionType ToPowerPointSelectionType(OfficeSelectionType type)
        {
            switch (type)
            {
                case OfficeSelectionType.Shapes: return PowerPoint.PpSelectionType.ppSelectionShapes;
                case OfficeSelectionType.Text: return PowerPoint.PpSelectionType.ppSelectionText;
                case OfficeSelectionType.Slides: return PowerPoint.PpSelectionType.ppSelectionSlides;
                default: return PowerPoint.PpSelectionType.ppSelectionNone;
            }
        }
    }
}
