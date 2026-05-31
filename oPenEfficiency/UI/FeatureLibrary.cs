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
        private static List<SidebarFeature> _allFeatures;
        public static List<SidebarFeature> AllFeatures
        {
            get
            {
                if (_allFeatures == null)
                {
                    _allFeatures = new List<SidebarFeature>();
                    
                    // 1. Add all unmigrated legacy features explicitly
                    string[] legacyIds = { "BtnStickyNote" };
                    foreach (var id in legacyIds)
                    {
                        var info = GetFeatureInfo(id);
                        _allFeatures.Add(new SidebarFeature { Id = id, Name = info.Tooltip ?? id.Replace("Btn", "") });
                    }

                    // 2. Auto-discover all other features
                    var autoFeatures = FeatureDiscovery.AllFeatures;
                    if (autoFeatures != null)
                    {
                        foreach (var f in autoFeatures)
                        {
                            if (!_allFeatures.Any(x => x.Id == f.Id))
                            {
                                _allFeatures.Add(new SidebarFeature { Id = f.Id, Name = f.Name });
                            }
                        }
                    }

                    // 3. Sort alphabetically for UI consistency
                    _allFeatures = _allFeatures.OrderBy(f => f.Name).ToList();
                }
                return _allFeatures;
            }
        }

        public static SidebarConfig GetEmergencyDefault() => GetStandardLayout();

        private static void AddById(SidebarSection section, string id)
        {
            var feature = AllFeatures.Find(f => f.Id == id);
            if (feature != null)
                section.Features.Add(feature);
#if DEBUG
            else
                System.Diagnostics.Debug.WriteLine($"[oPenEfficiency] WARNING: Layout references unregistered feature ID '{id}' in section '{section.Name}'");
#endif
        }

        /// <summary>
        /// Curated, focused layout for daily professional use.
        /// </summary>
        public static SidebarConfig GetStandardLayout()
        {
            var config = new SidebarConfig();

            var sizeSection = new SidebarSection { Name = "PROPERTIES & SIZE", Color = "#6366F1" };
            string[] sizeIds = { "BtnSearchBar", "BtnSizePanel", "BtnMatchSize", "BtnMatchWidth", "BtnMatchHeight", "BtnLockDimensions", "BtnShapeLock", "BtnSlideSize" };
            foreach (var id in sizeIds) AddById(sizeSection, id);
            config.Sections.Add(sizeSection);

            var alignSection = new SidebarSection { Name = "ALIGN", Color = "#10B981" };
            string[] alignIds = { "BtnAlignLeft", "BtnAlignCenter", "BtnAlignRight", "BtnAlignTop", "BtnAlignCenterVertical", "BtnAlignBottom", "BtnDiagonalAlign", "BtnDistributeHorizontal", "BtnDistributeVertical", "BtnRemoveHorizontalGaps", "BtnRemoveVerticalGaps", "BtnIncreaseHorizontalSpacing", "BtnDecreaseHorizontalSpacing", "BtnIncreaseVerticalSpacing", "BtnDecreaseVerticalSpacing", "BtnAdjustSpacing" };
            foreach (var id in alignIds) AddById(alignSection, id);
            config.Sections.Add(alignSection);

            var stretchSection = new SidebarSection { Name = "DOCK", Color = "#F59E0B" };
            string[] stretchIds = { "BtnStretchLeft", "BtnStretchRight", "BtnStretchTop", "BtnStretchBottom", "BtnStretchLeftEdge", "BtnStretchRightEdge", "BtnStretchTopEdge", "BtnStretchBottomEdge", "BtnDockLeft", "BtnDockRight", "BtnDockTop", "BtnDockBottom" };
            foreach (var id in stretchIds) AddById(stretchSection, id);
            config.Sections.Add(stretchSection);

            var arrangeSection = new SidebarSection { Name = "ARRANGE & ORGANIZE", Color = "#FBBF24" };
            string[] arrangeIds = { "BtnSwapPositions", "BtnCopyXY", "BtnPasteXY", "BtnSelectSameType", "BtnSyncObjects", "BtnArrangeInShape", "BtnArrangePro", "BtnRectifyRotation", "BtnAlignShapeAdjustments", "BtnHideSelected", "BtnShowHidden", "BtnHideMasterObjects", "BtnCreateMotionPath", "BtnBringToFront", "BtnBringForward", "BtnSendBackward", "BtnSendToBack", "BtnSnapShape", "BtnSnapToGrid", "BtnSnapToObjects" };
            foreach (var id in arrangeIds) AddById(arrangeSection, id);
            config.Sections.Add(arrangeSection);

            var textSection = new SidebarSection { Name = "TEXT & FONTS", Color = "#FDE047" };
            string[] textIds = { "BtnFontPanel", "BtnApplyTextTool", "BtnFitFormToText", "BtnTranslate", "BtnReplaceFont", "BtnParagraphDialog", "BtnSpecialChars", "BtnFormatNumbers", "BtnAlignText", "BtnTextDirection", "BtnChangeCase", "BtnCharacterSpacing", "BtnSpellCheck", "BtnDeleteText", "BtnSwapTextFormatted", "BtnSwapTextPlain" };
            foreach (var id in textIds) AddById(textSection, id);
            config.Sections.Add(textSection);

            var formatSection = new SidebarSection { Name = "FORMAT & STYLE", Color = "#06B6D4" };
            string[] formatIds = { "BtnPickColor", "BtnPickFillColor", "BtnPickLineColor", "BtnPickTextColor", "BtnThemeColor", "BtnTransparentColor", "BtnGlassHide", "BtnFormatShapeDialog", "BtnOptimizeFreeForm", "BtnStyleCheck", "BtnCleaner", "BtnFormatBold", "BtnFormatItalic", "BtnFormatUnderline", "BtnFormatStrikethrough", "BtnFormatSuperscript", "BtnFormatSubscript" };
            foreach (var id in formatIds) AddById(formatSection, id);
            config.Sections.Add(formatSection);

            var tablesSection = new SidebarSection { Name = "TABLE TOOLS", Color = "#3B82F6" };
            string[] tablesIds = { "BtnInsertColumnLeft", "BtnInsertColumnRight", "BtnInsertRowTop", "BtnInsertRowBottom", "BtnTableTranspose", "BtnTableSplit", "BtnConvertToShapes", "BtnTableDimensions", "BtnTableFormatPainter", "BtnTableSortAZ", "BtnTableSum", "BtnTableBranding", "BtnAlignToTableCell" };
            foreach (var id in tablesIds) AddById(tablesSection, id);
            config.Sections.Add(tablesSection);

            var visualsSection = new SidebarSection { Name = "VISUAL ELEMENTS", Color = "#EF4444" };
            string[] visualsIds = { "BtnIllustrativeSticker", "BtnAddSticker", "BtnHarveyBall", "BtnTrafficLight", "BtnThermometer", "BtnStarRating", "BtnCheckbox", "BtnNumeration", "BtnProgressSeries", "BtnGlassHide", "BtnRepeatShape", "BtnSplitShape", "BtnSplitByParagraphs" };
            foreach (var id in visualsIds) AddById(visualsSection, id);
            config.Sections.Add(visualsSection);

            var storySection = new SidebarSection { Name = "STORYLINING & REVIEW", Color = "#FDE047" };
            string[] storyIds = { "BtnAgendaWizard", "BtnSaveAgendaLayout", "BtnStoryline", "BtnStyleCheck", "BtnSlideGuidelines", "BtnCleaner", "BtnFlightMode", "BtnStickyNote", "BtnSlideNotes" };
            foreach (var id in storyIds) AddById(storySection, id);
            config.Sections.Add(storySection);

            var utilSection = new SidebarSection { Name = "UTILITIES & WIZARDS", Color = "#10B981" };
            string[] utilIds = { "BtnAssetManager", "BtnWinnerPicker", "BtnUpdateExcelCharts", "BtnQRCode", "BtnExportWizard", "BtnSettings", "BtnTemplateManager", "BtnStickyNote", "BtnSlidePaste" };
            foreach (var id in utilIds) AddById(utilSection, id);
            config.Sections.Add(utilSection);

            return config;
        }

        /// <summary>
        /// IMPORTANT: Comprehensive layout containing every discovered feature in the toolkit.
        /// ALL new features should be added here during development to ensure they are accessible.
        /// </summary>
        public static SidebarConfig GetAllFeaturesLayout()
        {
            var config = new SidebarConfig();

            var sizeSection = new SidebarSection { Name = "PROPERTIES & SIZE", Color = "#6366F1" };
            string[] sizeIds = { "BtnSearchBar", "BtnSizePanel", "BtnMatchSize", "BtnMatchWidth", "BtnMatchHeight", "BtnMatchAngles", "BtnMagicResizer", "BtnLockDimensions", "BtnShapeLock", "BtnSlideSize" };
            foreach (var id in sizeIds) AddById(sizeSection, id);
            config.Sections.Add(sizeSection);

            var alignSection = new SidebarSection { Name = "ALIGN", Color = "#10B981" };
            string[] alignIds = { "BtnAlignLeft", "BtnAlignCenter", "BtnAlignRight", "BtnAlignTop", "BtnAlignCenterVertical", "BtnAlignBottom", "BtnDiagonalAlign", "BtnLearnMargin", "BtnSetMargin", "BtnDistributeHorizontal", "BtnDistributeVertical", "BtnRemoveHorizontalGaps", "BtnRemoveVerticalGaps", "BtnIncreaseHorizontalSpacing", "BtnDecreaseHorizontalSpacing", "BtnIncreaseVerticalSpacing", "BtnDecreaseVerticalSpacing", "BtnAdjustSpacing" };
            foreach (var id in alignIds) AddById(alignSection, id);
            config.Sections.Add(alignSection);

            var stretchSection = new SidebarSection { Name = "DOCK", Color = "#F59E0B" };
            string[] stretchIds = { "BtnStretchLeft", "BtnStretchRight", "BtnStretchTop", "BtnStretchBottom", "BtnStretchLeftEdge", "BtnStretchRightEdge", "BtnStretchTopEdge", "BtnStretchBottomEdge", "BtnDockLeft", "BtnDockRight", "BtnDockTop", "BtnDockBottom" };
            foreach (var id in stretchIds) AddById(stretchSection, id);
            config.Sections.Add(stretchSection);

            var arrangeSection = new SidebarSection { Name = "ARRANGE & ORGANIZE", Color = "#FBBF24" };
            string[] arrangeIds = { "BtnSwapPositions", "BtnCopyXY", "BtnPasteXY", "BtnSelectSameType", "BtnSyncObjects", "BtnArrangeGrid", "BtnArrangeInShape", "BtnArrangePro", "BtnMultiSwap", "BtnCloneBottom", "BtnCloneLeft", "BtnCloneRight", "BtnCloneTop", "BtnObjectConnector", "BtnRectifyRotation", "BtnAlignShapeAdjustments", "BtnFlipHorizontal", "BtnFlipVertical", "BtnRotateShape", "BtnPositionPainter", "BtnHideSelected", "BtnShowHidden", "BtnHideMasterObjects", "BtnCreateMotionPath", "BtnBringToFront", "BtnBringForward", "BtnSendBackward", "BtnSendToBack", "BtnSnapShape", "BtnSnapToGrid", "BtnSnapToObjects" };
            foreach (var id in arrangeIds) AddById(arrangeSection, id);
            config.Sections.Add(arrangeSection);

            var textSection = new SidebarSection { Name = "TEXT & FONTS", Color = "#FDE047" };
            string[] textIds = { "BtnFontPanel", "BtnApplyTextTool", "BtnInsertText", "BtnFitFormToText", "BtnTranslate", "BtnReplaceFont", "BtnParagraphDialog", "BtnSpecialChars", "BtnFormatNumbers", "BtnAlignText", "BtnTextDirection", "BtnChangeCase", "BtnIncreaseFontSize", "BtnDecreaseFontSize", "BtnCharacterSpacing", "BtnSpellCheck", "BtnDeleteText", "BtnSwapTextFormatted", "BtnSwapTextPlain" };
            foreach (var id in textIds) AddById(textSection, id);
            config.Sections.Add(textSection);

            var formatSection = new SidebarSection { Name = "FORMAT & STYLE", Color = "#06B6D4" };
            string[] formatIds = { "BtnPickColor", "BtnPickFillColor", "BtnPickLineColor", "BtnPickTextColor", "BtnThemeColor", "BtnTransparentColor", "BtnColorOverlay", "BtnSmartCorners", "BtnTransparency", "BtnGlassHide", "BtnFormatShapeDialog", "BtnOptimizeFreeForm", "BtnStyleCheck", "BtnCleaner", "BtnFormatBold", "BtnFormatItalic", "BtnFormatUnderline", "BtnFormatStrikethrough", "BtnFormatSuperscript", "BtnFormatSubscript" };
            foreach (var id in formatIds) AddById(formatSection, id);
            config.Sections.Add(formatSection);

            var tablesSection = new SidebarSection { Name = "TABLE TOOLS", Color = "#3B82F6" };
            string[] tablesIds = { "BtnInsertColumnLeft", "BtnInsertColumnRight", "BtnInsertRowTop", "BtnInsertRowBottom", "BtnTableTranspose", "BtnTableSplit", "BtnConvertToShapes", "BtnConvertToTable", "BtnTableColumnInsertion", "BtnTableColumnWidth", "BtnTableHeatmap", "BtnTableRowHeight", "BtnTableRowInsertion", "BtnTableDimensions", "BtnTableFormatPainter", "BtnTableSortAZ", "BtnTableSum", "BtnTableBranding", "BtnAlignToTableCell" };
            foreach (var id in tablesIds) AddById(tablesSection, id);
            config.Sections.Add(tablesSection);

            var visualsSection = new SidebarSection { Name = "VISUAL ELEMENTS", Color = "#EF4444" };
            string[] visualsIds = { "BtnIllustrativeSticker", "BtnAddSticker", "BtnHarveyBall", "BtnTrafficLight", "BtnThermometer", "BtnStarRating", "BtnCheckbox", "BtnNumeration", "BtnProgressSeries", "BtnChartOverlay", "BtnMapWizard", "BtnSeriesGenerator", "BtnGlassHide", "BtnRepeatShape", "BtnSplitShape", "BtnSplitByParagraphs" };
            foreach (var id in visualsIds) AddById(visualsSection, id);
            config.Sections.Add(visualsSection);

            var storySection = new SidebarSection { Name = "STORYLINING & REVIEW", Color = "#FDE047" };
            string[] storyIds = { "BtnAgendaWizard", "BtnSaveAgendaLayout", "BtnStoryline", "BtnAnonymize", "BtnConsolidateMasters", "BtnStyleCheck", "BtnSlideGuidelines", "BtnCleaner", "BtnFlightMode", "BtnStickyNote", "BtnStickyNoteManager", "BtnSlideNotes" };
            foreach (var id in storyIds) AddById(storySection, id);
            config.Sections.Add(storySection);

            var utilSection = new SidebarSection { Name = "UTILITIES & WIZARDS", Color = "#10B981" };
            string[] utilIds = { "BtnAssetManager", "BtnWinnerPicker", "BtnUpdateExcelCharts", "BtnQRCode", "BtnDocumentAutomation", "BtnExcelLinkManager", "BtnExploreFeatures", "BtnMoveToBackup", "BtnPropertyExtraction", "BtnTagInspector", "BtnNewSlide", "BtnExportWizard", "BtnSettings", "BtnTemplateManager", "BtnStickyNote", "BtnSlidePaste" };
            foreach (var id in utilIds) AddById(utilSection, id);
            config.Sections.Add(utilSection);

            return config;
        }

        public static SidebarConfig GetMinimalistLayout()
        {
            var config = new SidebarConfig();
            var essentialSection = new SidebarSection { Name = "ESSENTIALS", Color = "#3B82F6" };
            string[] essentialIds = { "BtnSearchBar", "BtnSizePanel", "BtnAlignLeft", "BtnAlignCenter", "BtnAlignRight", "BtnAlignTop", "BtnAlignCenterVertical", "BtnAlignBottom", "BtnFontPanel" };
            foreach (var id in essentialIds) AddById(essentialSection, id);
            config.Sections.Add(essentialSection);
            return config;
        }

        public static FeatureDisplayInfo GetFeatureInfo(string id)
        {
            var autoInfo = FeatureDiscovery.GetFeatureInfo(id);
            if (autoInfo != null) return autoInfo.Value;

            switch (id)
            {
                case "BtnStickyNote": return new FeatureDisplayInfo { Tooltip = "Sticky Note", Color = "#10B981", Description = "Inserts a sticky note with a random professional color to the side of your slide for internal comments. Right-click for bulk management (hide, move, delete).", RequiredType = PowerPoint.PpSelectionType.ppSelectionNone, DetailedHelpText = "Inserts a sticky note for internal comments.\n\nRight-click options:\n• Move off slide/selected slides\n• Move back on slide\n• Delete notes\n• Convert comments/shapes to notes" };
                default: return new FeatureDisplayInfo { Tooltip = id, Color = "#9CA3AF", IconData = "M12,2L14.5,9H22L16,13.5L18.5,21L12,16.5L5.5,21L8,13.5L2,9H9.5L12,2Z" };
            }
        }
    }
}
