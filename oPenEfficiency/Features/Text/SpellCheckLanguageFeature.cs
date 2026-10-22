using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using oPenEfficiency.UI.Dialogs;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Text
{
    [FeatureMetadata(
        Id = "BtnSpellCheck", 
        Name = "Set Spell Check Language", 
        Tooltip = "Spell Check", 
        IconData = "M21 11H6.83l3.58-3.59L9 6l-6 6 6 6 1.41-1.41L6.83 13H21z", 
        Color = "#FDE047", 
        MinSelection = 0, 
        RequiredType = OfficeSelectionType.None,
        Description = "Sets the proofing language for selected shapes, selected slides, or the entire presentation.",
        DetailedHelpText = "### Set Spell Check Language\n\nChoose a proofing language and apply it to selected shapes, selected slides, or the entire presentation. Presentation scope also updates slide masters, custom layouts, and the presentation's default language; narrower scopes do not change that global default. Groups and table cells are included, and successful changes create one PowerPoint Undo entry."
    )]
    public static class SpellCheckLanguageFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, null);
        }

        public static bool Execute(PowerPointManager manager, SpellCheckScope? initialScope)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app == null || app.ActivePresentation == null) return false;
                var pres = app.ActivePresentation;

                MsoLanguageID? detectedLang = null;
                try {
                    var sel = app.ActiveWindow.Selection;
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText) {
                        detectedLang = sel.TextRange.LanguageID;
                    } else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0) {
                        var shape = sel.ShapeRange[1];
                        if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue) {
                            detectedLang = shape.TextFrame.TextRange.LanguageID;
                        } else if (shape.HasTable == MsoTriState.msoTrue && shape.Table.Rows.Count > 0 && shape.Table.Columns.Count > 0) {
                            var cell = shape.Table.Cell(1, 1);
                            if (cell.Shape.HasTextFrame == MsoTriState.msoTrue && cell.Shape.TextFrame.HasText == MsoTriState.msoTrue) {
                                detectedLang = cell.Shape.TextFrame.TextRange.LanguageID;
                            }
                        }
                    }
                    if (detectedLang == null && pres != null) {
                        detectedLang = pres.DefaultLanguageID;
                    }
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "SpellCheckLanguageFeature.DetectLanguage");
                }

                var dialog = new SpellCheckLanguageWindow(detectedLang, initialScope);
                if (app.ActiveWindow != null)
                    new System.Windows.Interop.WindowInteropHelper(dialog).Owner = new IntPtr(app.HWND);

                if (dialog.ShowDialog() != true) return true;

                MsoLanguageID langId = dialog.SelectedLanguage;
                SpellCheckScope scope = dialog.SelectedScope;
                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "SpellCheckLanguageFeature.StartNewUndoEntry"); }

                int changed;
                switch (scope)
                {
                    case SpellCheckScope.Presentation:
                        changed = SetLanguageForPresentation(pres, langId);
                        if (pres.DefaultLanguageID != langId)
                        {
                            pres.DefaultLanguageID = langId;
                            changed++;
                        }
                        break;
                    case SpellCheckScope.SelectedSlides:
                        changed = SetLanguageForSelectedSlides(app.ActiveWindow?.Selection, langId);
                        break;
                    default:
                        changed = SetLanguageForSelectedShapes(app.ActiveWindow?.Selection, langId);
                        break;
                }

                if (changed == 0)
                {
                    System.Windows.MessageBox.Show(
                        "No text needed a language change in the selected scope.",
                        "Set Spell Check Language",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex.Message, "SpellCheckLanguageFeature");
                return false;
            }
        }

        private static int SetLanguageForPresentation(PowerPoint.Presentation pres, MsoLanguageID langId)
            => Services.ProofingLanguageService.SetLanguageForPresentation(pres, langId);

        private static int SetLanguageForSelectedSlides(PowerPoint.Selection selection, MsoLanguageID langId)
            => Services.ProofingLanguageService.SetLanguageForSelectedSlides(selection, langId);

        private static int SetLanguageForSelectedShapes(PowerPoint.Selection selection, MsoLanguageID langId)
            => Services.ProofingLanguageService.SetLanguageForSelectedShapes(selection, langId);

        private static int SetLanguageForShapes(PowerPoint.Shapes shapes, MsoLanguageID langId)
            => Services.ProofingLanguageService.SetLanguageForShapes(shapes, langId);

        private static int SetLanguageForShape(PowerPoint.Shape shape, MsoLanguageID langId)
            => Services.ProofingLanguageService.SetLanguageForShape(shape, langId);
    }
}
