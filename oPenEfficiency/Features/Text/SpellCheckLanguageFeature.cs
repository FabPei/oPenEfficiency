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
        RequiredType = PowerPoint.PpSelectionType.ppSelectionNone,
        Description = "Sets the proofing dictionary language globally across all slides and master slides. Essential for keeping spell checks clean.",
        DetailedHelpText = "### Set Spell Check Language\n\nTraverses every text range across all slides and forcefully sets the proofreading language to a single chosen locale, fixing mixed-language spell check issues from copy-pasting."
    )]
    public static class SpellCheckLanguageFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var app = manager.GetApplication();
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
                } catch { }

                var dialog = new SpellCheckLanguageWindow(detectedLang);
                if (dialog.ShowDialog() == true)
                {
                    MsoLanguageID langId = dialog.SelectedLanguage;
                    SpellCheckScope scope = dialog.SelectedScope;

                    switch (scope)
                    {
                        case SpellCheckScope.Presentation:
                            SetLanguageForPresentation(pres, langId);
                            break;
                        case SpellCheckScope.SelectedSlides:
                            SetLanguageForSelectedSlides(app.ActiveWindow.Selection, langId);
                            break;
                        case SpellCheckScope.SelectedShapes:
                            SetLanguageForSelectedShapes(app.ActiveWindow.Selection, langId);
                            break;
                    }

                    // Fallback for default text language
                    pres.DefaultLanguageID = langId;
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex.Message, "SpellCheckLanguageFeature");
                return false;
            }
        }

        private static void SetLanguageForPresentation(PowerPoint.Presentation pres, MsoLanguageID langId)
        {
            if (pres == null) return;
            
            // Loop through all slides safely
            try {
                for (int i = 1; i <= pres.Slides.Count; i++)
                {
                    try {
                        SetLanguageForShapes(pres.Slides[i].Shapes, langId);
                    } catch { }
                }
            } catch { }

            // Loop through master slides and layouts to ensure placeholders get updated
            try {
                if (pres.Designs != null) {
                    for (int i = 1; i <= pres.Designs.Count; i++) {
                        try {
                            var design = pres.Designs[i];
                            if (design.SlideMaster != null) {
                                SetLanguageForShapes(design.SlideMaster.Shapes, langId);
                                if (design.SlideMaster.CustomLayouts != null) {
                                    for (int j = 1; j <= design.SlideMaster.CustomLayouts.Count; j++) {
                                        try {
                                            SetLanguageForShapes(design.SlideMaster.CustomLayouts[j].Shapes, langId);
                                        } catch { }
                                    }
                                }
                            }
                        } catch { }
                    }
                }
            } catch { }
        }

        private static void SetLanguageForSelectedSlides(PowerPoint.Selection selection, MsoLanguageID langId)
        {
            if (selection == null) return;
            try {
                for (int i = 1; i <= selection.SlideRange.Count; i++)
                {
                    try {
                        SetLanguageForShapes(selection.SlideRange[i].Shapes, langId);
                    } catch { }
                }
            } catch { }
        }

        private static void SetLanguageForSelectedShapes(PowerPoint.Selection selection, MsoLanguageID langId)
        {
            if (selection == null || selection.Type == PowerPoint.PpSelectionType.ppSelectionNone) return;
            try {
                for (int i = 1; i <= selection.ShapeRange.Count; i++)
                {
                    try {
                        SetLanguageForShape(selection.ShapeRange[i], langId);
                    } catch { }
                }
            } catch { }
        }

        private static void SetLanguageForShapes(PowerPoint.Shapes shapes, MsoLanguageID langId)
        {
            if (shapes == null) return;
            try {
                for (int i = 1; i <= shapes.Count; i++)
                {
                    try {
                        SetLanguageForShape(shapes[i], langId);
                    } catch { }
                }
            } catch { }
        }

        private static void SetLanguageForShape(PowerPoint.Shape shape, MsoLanguageID langId)
        {
            try
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
                {
                    shape.TextFrame.TextRange.LanguageID = langId;
                }
                
                if (shape.HasTable == MsoTriState.msoTrue)
                {
                    for (int r = 1; r <= shape.Table.Rows.Count; r++)
                    {
                        for (int c = 1; c <= shape.Table.Columns.Count; c++)
                        {
                            try {
                                var cell = shape.Table.Cell(r, c);
                                if (cell.Shape.HasTextFrame == MsoTriState.msoTrue && cell.Shape.TextFrame.HasText == MsoTriState.msoTrue)
                                {
                                    cell.Shape.TextFrame.TextRange.LanguageID = langId;
                                }
                            } catch { }
                        }
                    }
                }

                if (shape.Type == MsoShapeType.msoGroup)
                {
                    for (int i = 1; i <= shape.GroupItems.Count; i++)
                    {
                        try {
                            SetLanguageForShape(shape.GroupItems[i], langId);
                        } catch { }
                    }
                }
            }
            catch { }
        }
    }
}
