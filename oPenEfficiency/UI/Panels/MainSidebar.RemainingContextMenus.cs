using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Features;
using oPenEfficiency.Features.Charts;
using oPenEfficiency.Features.Text;
using oPenEfficiency.Features.Utilities;
using oPenEfficiency.Features.Visuals;
using oPenEfficiency.Utils;

namespace oPenEfficiency
{
    public partial class MainSidebar
    {
        private static void AddRemainingMenuItem(ButtonBase button, int index, string fallback, Action action)
        {
            var item = new MenuItem { Header = GetCatalogVariantLabel(button.Name, index, fallback) };
            item.Click += (sender, args) => action();
            button.ContextMenu.Items.Add(item);
        }

        private bool TryAttachRemainingFeatureContextMenu(ButtonBase button)
        {
            string id = button.Name;
            var supported = new HashSet<string>(StringComparer.Ordinal)
            {
                "BtnSyncObjects", "BtnArrangePro", "BtnParagraphDialog", "BtnSwapTextFormatted",
                "BtnFormatShapeDialog", "BtnInsertExcelValue", "BtnConnectedTextLabel", "BtnAssetManager", "BtnExcelLinkManager",
                "BtnQRCode", "BtnExportWizard", "BtnGanttChart", "BtnTemplateManager", "BtnSlidePaste",
                "BtnHtmlSlideImport", "BtnAnonymize", "BtnChartOverlay", "BtnConvertToTable",
                "BtnDocumentAutomation", "BtnLocalSlideRebuild", "BtnMapWizard", "BtnMoveToBackup",
                "BtnPropertyExtraction", "BtnSeriesGenerator", "BtnQuickContextMenu", "BtnShapeFill", "BtnPerformanceMode", "BtnCountryFlags"
            };
            if (!supported.Contains(id)) return false;
            if (button.ContextMenu.Items.Count > 0) button.ContextMenu.Items.Add(new Separator());

            switch (id)
            {
                case "BtnQuickContextMenu":
                    AddRemainingMenuItem(button, 0, "Configure Quick Menu...", () =>
                    {
                        var window = new UI.Dialogs.QuickContextMenuConfigWindow();
                        var app = GetManager()?.GetApplication();
                        if (app != null) new System.Windows.Interop.WindowInteropHelper(window).Owner = (IntPtr)app.HWND;
                        window.ShowDialog();
                    });
                    break;
                case "BtnShapeFill":
                    AddRemainingMenuItem(button, 0, "Open Shape Fill Palette", () =>
                        ExecuteFeature(button.Name, button, Services.Commands.CommandInvocationSource.ContextMenu));
                    AddRemainingMenuItem(button, 1, "No Fill", () =>
                        new Services.ShapeFillService(GetManager()).ApplyNoFill());
                    AddRemainingMenuItem(button, 2, "Eyedropper", () =>
                        Features.PickFillColorFeature.Execute(GetManager()));
                    break;
                case "BtnPerformanceMode":
                    AddRemainingMenuItem(button, 0, "Configure Performance Mode...", () =>
                        _host?.ConfigurePerformanceMode());
                    break;
                case "BtnCountryFlags":
                    AddRemainingMenuItem(button, 0, "Open Country Flag Library", () =>
                        ExecuteFeature(button.Name, button, Services.Commands.CommandInvocationSource.ContextMenu));
                    break;
                case "BtnSyncObjects":
                    AddRemainingMenuItem(button, 0, "Open Sync Objects", () => SyncObjectsFeature.ExecuteDialog(GetManager()));
                    AddRemainingMenuItem(button, 1, "Sync Format", () => SyncObjectsFeature.SyncShapeProperties(GetManager(), true, false, false));
                    AddRemainingMenuItem(button, 2, "Sync Position", () => SyncObjectsFeature.SyncShapeProperties(GetManager(), false, true, false));
                    AddRemainingMenuItem(button, 3, "Sync Content", () => SyncObjectsFeature.SyncShapeProperties(GetManager(), false, false, true));
                    AddRemainingMenuItem(button, 4, "Sync All Properties", () => SyncObjectsFeature.SyncShapeProperties(GetManager(), true, true, true));
                    break;
                case "BtnArrangePro":
                    AddRemainingMenuItem(button, 0, "Circle Around Selection", () => ArrangeProFeature.Execute(GetManager(), ArrangeProFeature.Preset.CircleAroundSelection));
                    AddRemainingMenuItem(button, 1, "Circle Around Slide", () => ArrangeProFeature.Execute(GetManager(), ArrangeProFeature.Preset.CircleAroundSlide));
                    AddRemainingMenuItem(button, 2, "Arc Around Selection", () => ArrangeProFeature.Execute(GetManager(), ArrangeProFeature.Preset.ArcAroundSelection));
                    AddRemainingMenuItem(button, 3, "Rotate Shapes to Center", () => ArrangeProFeature.Execute(GetManager(), ArrangeProFeature.Preset.RotateToCenter));
                    AddRemainingMenuItem(button, 4, "Open Custom Arrangement", () => ArrangeProFeature.Execute(GetManager()));
                    break;
                case "BtnParagraphDialog":
                    AddRemainingMenuItem(button, 0, "Open Paragraph Dialog", () => ParagraphSpacingContextActions.OpenDialog(GetManager()));
                    AddRemainingMenuItem(button, 1, "Single Line Spacing", () => ParagraphSpacingContextActions.ApplyLineSpacing(GetManager(), 1f));
                    AddRemainingMenuItem(button, 2, "1.15 Line Spacing", () => ParagraphSpacingContextActions.ApplyLineSpacing(GetManager(), 1.15f));
                    AddRemainingMenuItem(button, 3, "1.5 Line Spacing", () => ParagraphSpacingContextActions.ApplyLineSpacing(GetManager(), 1.5f));
                    AddRemainingMenuItem(button, 4, "Remove Paragraph Spacing", () => ParagraphSpacingContextActions.RemoveParagraphSpacing(GetManager()));
                    break;
                case "BtnSwapTextFormatted":
                    AddRemainingMenuItem(button, 0, "Swap Formatted Text", () => SwapTextFormattedFeature.Execute(GetManager()));
                    AddRemainingMenuItem(button, 1, "Swap and Shrink Overflowing Text", () => SwapTextFormattedFeature.Execute(GetManager(), SwapTextFormattedFeature.FitMode.ShrinkOverflow));
                    AddRemainingMenuItem(button, 2, "Swap and Fit Shapes to Text", () => SwapTextFormattedFeature.Execute(GetManager(), SwapTextFormattedFeature.FitMode.FitShapesToText));
                    break;
                case "BtnFormatShapeDialog":
                    AddRemainingMenuItem(button, 0, "Open Format Shape", () => FormatShapeDialogFeature.Execute(GetManager()));
                    AddRemainingMenuItem(button, 1, "Reset to Default Shape Formatting", () => FormatShapeDialogFeature.Execute(GetManager(), FormatShapeDialogFeature.ContextAction.ResetToDefault));
                    AddRemainingMenuItem(button, 2, "Remove Shadow", () => FormatShapeDialogFeature.Execute(GetManager(), FormatShapeDialogFeature.ContextAction.RemoveShadow));
                    AddRemainingMenuItem(button, 3, "Remove Reflection", () => FormatShapeDialogFeature.Execute(GetManager(), FormatShapeDialogFeature.ContextAction.RemoveReflection));
                    AddRemainingMenuItem(button, 4, "Remove Glow and Soft Edges", () => FormatShapeDialogFeature.Execute(GetManager(), FormatShapeDialogFeature.ContextAction.RemoveGlowAndSoftEdges));
                    AddRemainingMenuItem(button, 5, "Remove 3-D Format and Rotation", () => FormatShapeDialogFeature.Execute(GetManager(), FormatShapeDialogFeature.ContextAction.RemoveThreeD));
                    AddRemainingMenuItem(button, 6, "Remove All Effects", () => FormatShapeDialogFeature.Execute(GetManager(), FormatShapeDialogFeature.ContextAction.RemoveAllEffects));
                    break;
                case "BtnInsertExcelValue":
                    AddRemainingMenuItem(button, 0, "Insert at Text Cursor", () => InsertExcelValueFeature.Execute(GetManager(), InsertExcelValueFeature.InsertMode.InlineAtCursor));
                    AddRemainingMenuItem(button, 1, "Insert as Linked Text Box", () => InsertExcelValueFeature.Execute(GetManager(), InsertExcelValueFeature.InsertMode.LinkedTextBox));
                    AddRemainingMenuItem(button, 2, "Insert as Unlinked Value", () => InsertExcelValueFeature.Execute(GetManager(), InsertExcelValueFeature.InsertMode.UnlinkedTextBox));
                    AddRemainingMenuItem(button, 3, "Refresh Selected Linked Value", () => InsertExcelValueFeature.RefreshSelected(GetManager()));
                    AddRemainingMenuItem(button, 4, "Unlink Selected Value", () => InsertExcelValueFeature.UnlinkSelected(GetManager()));
                    break;
                case "BtnConnectedTextLabel":
                    AddRemainingMenuItem(button, 0, "Send to Word", () => ConnectedTextLabelFeature.SendTo(GetManager(), ConnectedTextLabelFeature.TargetHost.Word));
                    AddRemainingMenuItem(button, 1, "Send to OneNote", () => ConnectedTextLabelFeature.SendTo(GetManager(), ConnectedTextLabelFeature.TargetHost.OneNote));
                    AddRemainingMenuItem(button, 2, "Refresh All Open Targets", () => ConnectedTextLabelFeature.RefreshAll(GetManager()));
                    AddRemainingMenuItem(button, 3, "Disconnect Selected Source", () => ConnectedTextLabelFeature.DisconnectSelected(GetManager()));
                    break;
                case "BtnAssetManager":
                    AddRemainingMenuItem(button, 0, "Open Slides", () => _host?.OpenAssetManager("Slides"));
                    AddRemainingMenuItem(button, 1, "Open Images", () => _host?.OpenAssetManager("Images"));
                    AddRemainingMenuItem(button, 2, "Open Shapes", () => _host?.OpenAssetManager("Shapes"));
                    AddRemainingMenuItem(button, 3, "Open My Assets", () => _host?.OpenAssetManager("MyAssets"));
                    AddRemainingMenuItem(button, 4, "Save Current Selection", () => _host?.OpenAssetManager("MyAssets", true));
                    break;
                case "BtnExcelLinkManager":
                    AddRemainingMenuItem(button, 0, "Show All Links", () => ExcelLinkManagerFeature.Execute(GetManager(), UI.ExcelLinkManagerStartAction.ShowAll));
                    AddRemainingMenuItem(button, 1, "Show Broken Links", () => ExcelLinkManagerFeature.Execute(GetManager(), UI.ExcelLinkManagerStartAction.ShowBroken));
                    AddRemainingMenuItem(button, 2, "Refresh All Links", () => ExcelLinkManagerFeature.Execute(GetManager(), UI.ExcelLinkManagerStartAction.RefreshAll));
                    AddRemainingMenuItem(button, 3, "Relink Source", () => ExcelLinkManagerFeature.Execute(GetManager(), UI.ExcelLinkManagerStartAction.RelinkSource));
                    AddRemainingMenuItem(button, 4, "Open Source Workbook", () => ExcelLinkManagerFeature.Execute(GetManager(), UI.ExcelLinkManagerStartAction.OpenSource));
                    break;
                case "BtnQRCode":
                    AddRemainingMenuItem(button, 0, "Enter URL or Text", () => ShowQrPreset(string.Empty));
                    AddRemainingMenuItem(button, 1, "Use Selected Text", () => ShowQrPreset(GetSelectedTextForRemainingMenu()));
                    AddRemainingMenuItem(button, 2, "Use Clipboard Text", () => ShowQrPreset(GetClipboardTextForRemainingMenu()));
                    break;
                case "BtnExportWizard":
                    AddRemainingMenuItem(button, 0, "PPTX — Entire Presentation", () => ShowExportPreset(UI.ExportWizardStartFormat.Pptx, UI.ExportWizardStartScope.EntirePresentation));
                    AddRemainingMenuItem(button, 1, "PDF — Entire Presentation", () => ShowExportPreset(UI.ExportWizardStartFormat.Pdf, UI.ExportWizardStartScope.EntirePresentation));
                    AddRemainingMenuItem(button, 2, "PPTX + PDF — Entire Presentation", () => ShowExportPreset(UI.ExportWizardStartFormat.PptxAndPdf, UI.ExportWizardStartScope.EntirePresentation));
                    AddRemainingMenuItem(button, 3, "PDF — Selected Slides", () => ShowExportPreset(UI.ExportWizardStartFormat.Pdf, UI.ExportWizardStartScope.SelectedSlides));
                    AddRemainingMenuItem(button, 4, "Editable HTML — Entire Presentation", () => ShowExportPreset(UI.ExportWizardStartFormat.Html, UI.ExportWizardStartScope.EntirePresentation));
                    break;
                case "BtnGanttChart":
                    AddRemainingMenuItem(button, 0, "Open Editor", () => OpenGanttPreset(button, UI.Panels.GanttStartAction.Open));
                    AddRemainingMenuItem(button, 1, "Create New Chart", () => OpenGanttPreset(button, UI.Panels.GanttStartAction.CreateNew));
                    AddRemainingMenuItem(button, 2, "Update Selected Chart", () => OpenGanttPreset(button, UI.Panels.GanttStartAction.UpdateSelected));
                    AddRemainingMenuItem(button, 3, "Import Excel", () => OpenGanttPreset(button, UI.Panels.GanttStartAction.ImportExcel));
                    AddRemainingMenuItem(button, 4, "Sync Linked Excel", () => OpenGanttPreset(button, UI.Panels.GanttStartAction.SyncExcel));
                    break;
                case "BtnTemplateManager":
                    AddRemainingMenuItem(button, 0, "Open Templates", () => OpenTemplateManagerPreset(UI.Dialogs.TemplateManagerStartAction.Templates));
                    AddRemainingMenuItem(button, 1, "Search Templates", () => OpenTemplateManagerPreset(UI.Dialogs.TemplateManagerStartAction.Search));
                    AddRemainingMenuItem(button, 2, "Add Template File", () => OpenTemplateManagerPreset(UI.Dialogs.TemplateManagerStartAction.AddTemplate));
                    AddRemainingMenuItem(button, 3, "Manage Watched Folders", () => OpenTemplateManagerPreset(UI.Dialogs.TemplateManagerStartAction.WatchedFolders));
                    AddRemainingMenuItem(button, 4, "Restore Hidden Templates", () => OpenTemplateManagerPreset(UI.Dialogs.TemplateManagerStartAction.RestoreHidden));
                    break;
                case "BtnSlidePaste":
                    AddRemainingMenuItem(button, 0, "Insert with Links and Shadows", () => SlidePasteFeature.Execute(GetManager(), true, true));
                    AddRemainingMenuItem(button, 1, "Insert with Links", () => SlidePasteFeature.Execute(GetManager(), true, false));
                    AddRemainingMenuItem(button, 2, "Insert with Shadows", () => SlidePasteFeature.Execute(GetManager(), false, true));
                    AddRemainingMenuItem(button, 3, "Insert without Links or Shadows", () => SlidePasteFeature.Execute(GetManager(), false, false));
                    break;
                case "BtnHtmlSlideImport":
                    AddRemainingMenuItem(button, 0, "Preserve HTML Styling", () => HtmlSlideImportFeature.Execute(GetManager(), HtmlSlideImportFeature.ImportPreset.PreserveHtmlStyling));
                    AddRemainingMenuItem(button, 1, "Apply Presentation Theme", () => HtmlSlideImportFeature.Execute(GetManager(), HtmlSlideImportFeature.ImportPreset.PresentationTheme));
                    AddRemainingMenuItem(button, 2, "Preserve Background Only", () => HtmlSlideImportFeature.Execute(GetManager(), HtmlSlideImportFeature.ImportPreset.PreserveBackgroundOnly));
                    AddRemainingMenuItem(button, 3, "Local Resources Only", () => HtmlSlideImportFeature.Execute(GetManager(), HtmlSlideImportFeature.ImportPreset.LocalResourcesOnly));
                    AddRemainingMenuItem(button, 4, "Review Import Options", () => HtmlSlideImportFeature.Execute(GetManager()));
                    break;
                case "BtnAnonymize":
                    AddRemainingMenuItem(button, 0, "Anonymize Current Slide", () => AnonymizeFeature.Execute(GetManager(), AnonymizeFeature.AnonymizeScope.CurrentSlide, false));
                    AddRemainingMenuItem(button, 1, "Anonymize Selected Slides", () => AnonymizeFeature.Execute(GetManager(), AnonymizeFeature.AnonymizeScope.SelectedSlides, false));
                    AddRemainingMenuItem(button, 2, "Anonymize Entire Presentation", () => AnonymizeFeature.Execute(GetManager(), AnonymizeFeature.AnonymizeScope.EntirePresentation, false));
                    AddRemainingMenuItem(button, 3, "Preview Affected Word Count", () => AnonymizeFeature.Execute(GetManager(), AnonymizeFeature.AnonymizeScope.SelectedSlides, true));
                    break;
                case "BtnChartOverlay":
                    AddRemainingMenuItem(button, 0, "Delta (Absolute)", () => ChartOverlayFeature.Execute(GetManager(), ChartOverlayFeature.OverlayPreset.Delta));
                    AddRemainingMenuItem(button, 1, "CAGR (%)", () => ChartOverlayFeature.Execute(GetManager(), ChartOverlayFeature.OverlayPreset.Cagr));
                    AddRemainingMenuItem(button, 2, "Growth (%)", () => ChartOverlayFeature.Execute(GetManager(), ChartOverlayFeature.OverlayPreset.Growth));
                    AddRemainingMenuItem(button, 3, "Average Level Line", () => ChartOverlayFeature.Execute(GetManager(), ChartOverlayFeature.OverlayPreset.AverageLevel));
                    AddRemainingMenuItem(button, 4, "Highlight Outliers", () => ChartOverlayFeature.Execute(GetManager(), ChartOverlayFeature.OverlayPreset.Outliers));
                    break;
                case "BtnConvertToTable":
                    AddRemainingMenuItem(button, 0, "Ask After Conversion", () => ConvertShapesToTableFeature.Execute(GetManager(), ConvertShapesToTableFeature.OriginalShapeDisposition.AskAfterConversion));
                    AddRemainingMenuItem(button, 1, "Keep Original Shapes", () => ConvertShapesToTableFeature.Execute(GetManager(), ConvertShapesToTableFeature.OriginalShapeDisposition.KeepOriginals));
                    AddRemainingMenuItem(button, 2, "Replace Original Shapes", () => ConvertShapesToTableFeature.Execute(GetManager(), ConvertShapesToTableFeature.OriginalShapeDisposition.ReplaceOriginals));
                    break;
                case "BtnDocumentAutomation":
                    AddRemainingMenuItem(button, 0, "Run Automation", () => DocumentAutomationFeature.Execute(GetManager(), DocumentAutomationFeature.Mode.Run));
                    AddRemainingMenuItem(button, 1, "Design Schema", () => DocumentAutomationFeature.Execute(GetManager(), DocumentAutomationFeature.Mode.Designer));
                    break;
                case "BtnLocalSlideRebuild":
                    AddRemainingMenuItem(button, 0, "Rebuild Image Locally", () => LocalSlideRebuildFeature.Execute(GetManager(), LocalSlideRebuildFeature.SourceMode.Image));
                    AddRemainingMenuItem(button, 1, "Rebuild Entire PDF Locally", () => LocalSlideRebuildFeature.Execute(GetManager(), LocalSlideRebuildFeature.SourceMode.EntirePdf));
                    AddRemainingMenuItem(button, 2, "Rebuild PDF Page Range", () => LocalSlideRebuildFeature.Execute(GetManager(), LocalSlideRebuildFeature.SourceMode.PdfPageRange));
                    break;
                case "BtnMapWizard":
                    AddRemainingMenuItem(button, 0, "New World Map", () => MapWizardFeature.Execute(GetManager(), MapWizardFeature.StartMode.NewWorld));
                    AddRemainingMenuItem(button, 1, "New Continent Map", () => MapWizardFeature.Execute(GetManager(), MapWizardFeature.StartMode.NewContinent));
                    AddRemainingMenuItem(button, 2, "New Country Map", () => MapWizardFeature.Execute(GetManager(), MapWizardFeature.StartMode.NewCountry));
                    AddRemainingMenuItem(button, 3, "Edit Selected Map", () => MapWizardFeature.Execute(GetManager(), MapWizardFeature.StartMode.EditSelected));
                    AddRemainingMenuItem(button, 4, "Open Recent Map Configuration", () => MapWizardFeature.Execute(GetManager(), MapWizardFeature.StartMode.Recent));
                    break;
                case "BtnMoveToBackup":
                    AddRemainingMenuItem(button, 0, "Move Selected Slides to Backup", () => MoveToBackupFeature.Execute(GetManager(), MoveToBackupFeature.BackupAction.Move));
                    AddRemainingMenuItem(button, 1, "Copy Selected Slides to Backup", () => MoveToBackupFeature.Execute(GetManager(), MoveToBackupFeature.BackupAction.Copy));
                    AddRemainingMenuItem(button, 2, "Locate Backup Slide", () => MoveToBackupFeature.Execute(GetManager(), MoveToBackupFeature.BackupAction.Locate));
                    break;
                case "BtnPropertyExtraction":
                    AddRemainingMenuItem(button, 0, "Inspect Selection", () => PropertyExtractionFeature.Execute(GetManager(), PropertyExtractionFeature.ExtractionScope.Selection, PropertyExtractionFeature.ExtractionOutput.Dialog));
                    AddRemainingMenuItem(button, 1, "Inspect Current Slide", () => PropertyExtractionFeature.Execute(GetManager(), PropertyExtractionFeature.ExtractionScope.CurrentSlide, PropertyExtractionFeature.ExtractionOutput.Dialog));
                    AddRemainingMenuItem(button, 2, "Copy Selection Report", () => PropertyExtractionFeature.Execute(GetManager(), PropertyExtractionFeature.ExtractionScope.Selection, PropertyExtractionFeature.ExtractionOutput.Clipboard));
                    AddRemainingMenuItem(button, 3, "Copy Slide Report", () => PropertyExtractionFeature.Execute(GetManager(), PropertyExtractionFeature.ExtractionScope.CurrentSlide, PropertyExtractionFeature.ExtractionOutput.Clipboard));
                    AddRemainingMenuItem(button, 4, "Copy Color List", () => PropertyExtractionFeature.Execute(GetManager(), PropertyExtractionFeature.ExtractionScope.Automatic, PropertyExtractionFeature.ExtractionOutput.ColorListClipboard));
                    break;
                case "BtnSeriesGenerator":
                    AddRemainingMenuItem(button, 0, "1–5 Numbers — Horizontal", () => SeriesGeneratorFeature.ExecutePreset(GetManager(), "Numbers", "Horizontal", 5));
                    AddRemainingMenuItem(button, 1, "A–E Letters — Horizontal", () => SeriesGeneratorFeature.ExecutePreset(GetManager(), "Letters", "Horizontal", 5));
                    AddRemainingMenuItem(button, 2, "1–5 Numbers — Vertical", () => SeriesGeneratorFeature.ExecutePreset(GetManager(), "Numbers", "Vertical", 5));
                    AddRemainingMenuItem(button, 3, "1–9 Numbers — 3×3 Matrix", () => SeriesGeneratorFeature.ExecutePreset(GetManager(), "Numbers", "Matrix", 9, 3));
                    AddRemainingMenuItem(button, 4, "Open Custom Series", () => SeriesGeneratorFeature.Execute(GetManager()));
                    break;
            }
            return true;
        }

        private void ShowQrPreset(string text)
        {
            const string key = "BtnQRCode";
            Window existing;
            if (_floatingWindows.TryGetValue(key, out existing) && existing.IsLoaded)
            {
                var qrWindow = existing as UI.QRCodeWindow;
                qrWindow?.SetInitialText(text);
                existing.Activate();
                return;
            }
            ToggleFloatingWindow(key, () => new UI.QRCodeWindow(GetManager(), text), null);
        }

        private void ShowExportPreset(UI.ExportWizardStartFormat format, UI.ExportWizardStartScope scope)
        {
            const string key = "BtnExportWizard";
            Window existing;
            if (_floatingWindows.TryGetValue(key, out existing) && existing.IsLoaded)
            {
                var wizard = existing as UI.ExportWizard;
                wizard?.ApplyStartSelection(format, scope);
                existing.Activate();
                return;
            }
            ToggleFloatingWindow(key, () => new UI.ExportWizard(format, scope, GetManager()), null);
        }

        private void OpenGanttPreset(ButtonBase button, UI.Panels.GanttStartAction action)
        {
            var toggle = button as ToggleButton;
            if (toggle != null) toggle.IsChecked = true;
            ShowOrCenterGanttWindow(GetManager(), toggle, action);
        }

        private void OpenTemplateManagerPreset(UI.Dialogs.TemplateManagerStartAction action)
        {
            var manager = GetManager();
            if (manager == null) return;
            var window = new UI.Dialogs.TemplateManagerWindow(manager, action);
            var app = manager.GetApplication();
            if (app != null) new System.Windows.Interop.WindowInteropHelper(window).Owner = (IntPtr)app.HWND;
            window.ShowDialog();
        }

        private string GetSelectedTextForRemainingMenu()
        {
            try
            {
                var selection = GetManager()?.GetApplication()?.ActiveWindow?.Selection;
                if (selection == null) return string.Empty;
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    return selection.TextRange.Text ?? string.Empty;
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var values = new List<string>();
                    foreach (PowerPoint.Shape shape in selection.ShapeRange)
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                            values.Add(shape.TextFrame.TextRange.Text);
                    }
                    return string.Join(Environment.NewLine, values);
                }
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "MainSidebar.GetSelectedTextForRemainingMenu"); }
            return string.Empty;
        }

        private static string GetClipboardTextForRemainingMenu()
        {
            try { return Clipboard.ContainsText() ? Clipboard.GetText() : string.Empty; }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MainSidebar.GetClipboardTextForRemainingMenu");
                return string.Empty;
            }
        }
    }
}
