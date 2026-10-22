using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Modal-free proofing-language mechanics (QA-S9 wave 3): applies a MsoLanguageID across the
    /// scopes the Set Spell Check Language feature offers (selected shapes/text, selected slides,
    /// the whole presentation including masters, layouts and the presentation default lives in the
    /// feature), recursing into groups and table cells. Extracted from SpellCheckLanguageFeature so
    /// the acceptance seams can drive the production apply path without the feature's modal
    /// language dialog; the feature keeps the dialog and scope wiring on top.
    /// </summary>
    internal static class ProofingLanguageService
    {
        public static int SetLanguageForPresentation(Presentation pres, Office.MsoLanguageID langId)
        {
            if (pres == null) return 0;
            int changed = 0;

            try
            {
                for (int i = 1; i <= pres.Slides.Count; i++)
                {
                    try { changed += SetLanguageForShapes(pres.Slides[i].Shapes, langId); }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "ProofingLanguageService.PresentationSlide"); }
                }
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ProofingLanguageService.PresentationSlides"); }

            try
            {
                if (pres.Designs != null)
                {
                    for (int i = 1; i <= pres.Designs.Count; i++)
                    {
                        try
                        {
                            var design = pres.Designs[i];
                            if (design.SlideMaster != null)
                            {
                                changed += SetLanguageForShapes(design.SlideMaster.Shapes, langId);
                                if (design.SlideMaster.CustomLayouts != null)
                                {
                                    for (int j = 1; j <= design.SlideMaster.CustomLayouts.Count; j++)
                                    {
                                        try { changed += SetLanguageForShapes(design.SlideMaster.CustomLayouts[j].Shapes, langId); }
                                        catch (Exception ex) { ExceptionLogger.Log(ex, "ProofingLanguageService.CustomLayout"); }
                                    }
                                }
                            }
                        }
                        catch (Exception ex) { ExceptionLogger.Log(ex, "ProofingLanguageService.Design"); }
                    }
                }
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ProofingLanguageService.Designs"); }
            return changed;
        }

        public static int SetLanguageForSelectedSlides(Selection selection, Office.MsoLanguageID langId)
        {
            if (selection == null || selection.Type != PpSelectionType.ppSelectionSlides) return 0;
            int changed = 0;
            try
            {
                for (int i = 1; i <= selection.SlideRange.Count; i++)
                {
                    try { changed += SetLanguageForShapes(selection.SlideRange[i].Shapes, langId); }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "ProofingLanguageService.SelectedSlide"); }
                }
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ProofingLanguageService.SelectedSlides"); }
            return changed;
        }

        public static int SetLanguageForSelectedShapes(Selection selection, Office.MsoLanguageID langId)
        {
            if (selection == null ||
                (selection.Type != PpSelectionType.ppSelectionShapes &&
                 selection.Type != PpSelectionType.ppSelectionText))
                return 0;

            int changed = 0;
            try
            {
                for (int i = 1; i <= selection.ShapeRange.Count; i++)
                {
                    try { changed += SetLanguageForShape(selection.ShapeRange[i], langId); }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "ProofingLanguageService.SelectedShape"); }
                }
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ProofingLanguageService.SelectedShapes"); }
            return changed;
        }

        public static int SetLanguageForShapes(Shapes shapes, Office.MsoLanguageID langId)
        {
            if (shapes == null) return 0;
            int changed = 0;
            try
            {
                for (int i = 1; i <= shapes.Count; i++)
                {
                    try { changed += SetLanguageForShape(shapes[i], langId); }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "ProofingLanguageService.Shape"); }
                }
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ProofingLanguageService.Shapes"); }
            return changed;
        }

        public static int SetLanguageForShape(Shape shape, Office.MsoLanguageID langId)
        {
            int changed = 0;
            try
            {
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    for (int i = 1; i <= shape.GroupItems.Count; i++)
                        changed += SetLanguageForShape(shape.GroupItems[i], langId);
                    return changed;
                }

                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    for (int r = 1; r <= shape.Table.Rows.Count; r++)
                    {
                        for (int c = 1; c <= shape.Table.Columns.Count; c++)
                        {
                            try
                            {
                                var cell = shape.Table.Cell(r, c);
                                if (cell.Shape.HasTextFrame == Office.MsoTriState.msoTrue && cell.Shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                {
                                    var range = cell.Shape.TextFrame.TextRange;
                                    if (range.LanguageID != langId)
                                    {
                                        range.LanguageID = langId;
                                        changed++;
                                    }
                                }
                            }
                            catch (Exception ex) { ExceptionLogger.Log(ex, "ProofingLanguageService.TableCell"); }
                        }
                    }
                    return changed;
                }

                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    var range = shape.TextFrame.TextRange;
                    if (range.LanguageID != langId)
                    {
                        range.LanguageID = langId;
                        changed++;
                    }
                }
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ProofingLanguageService.SetLanguageForShape"); }
            return changed;
        }
    }
}
