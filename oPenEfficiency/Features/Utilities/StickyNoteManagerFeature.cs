using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Sticky Note Manager - manage sticky notes (move, delete, convert comments).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStickyNoteManager",
        Name = "Sticky Note Manager",
        Tooltip = "Sticky note manager",
        IconData = "M19,3H5C3.89,3 3,3.89 3,5V19C3,20.11 3.89,21 5,21H19C20.11,21 21,20.11 21,19V5C21,3.89 20.11,3 19,3M19,19H5V5H19V19M7,7H17V9H7V7M7,11H17V13H7V11M7,15H17V17H7V15Z",
        Color = "#F43F5E",
        Description = "Manages sticky notes - move, delete, or convert to comments.",
        DetailedHelpText = "### Sticky Note Manager\nDisplays a consolidated list of all Sticky Notes across the entire presentation, allowing bulk review, editing, and removal.",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class StickyNoteManagerFeature
    {
        public enum ActionType
        {
            MoveOff,
            MoveOn,
            Delete,
            ConvertComments,
            PrefixName
        }

        public enum Scope
        {
            ThisSlide,
            SelectedSlides
        }

        /// <summary>
        /// Wrapper for auto-discovery - deletes sticky notes on this slide.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, ActionType.Delete, Scope.ThisSlide);
        }

        public static bool Execute(PowerPointManager manager, ActionType action, Scope scope)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                bool selectedOnly = (scope == Scope.SelectedSlides);

                if (action == ActionType.PrefixName)
                {
                    try
                    {
                        var selection = app.ActiveWindow.Selection;
                        if (selection.Type == PpSelectionType.ppSelectionShapes)
                        {
                            foreach (Shape shape in selection.ShapeRange)
                            {
                                if (!shape.Name.StartsWith("StickyNote", StringComparison.OrdinalIgnoreCase))
                                {
                                    shape.Name = "StickyNote_" + shape.Name;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "StickyNoteManagerFeature.PrefixName");
                    }
                    return true;
                }

                IEnumerable<Slide> slides;
                if (!selectedOnly)
                {
                    var activeSlide = app.ActiveWindow.View.Slide as Slide;
                    slides = activeSlide != null ? new List<Slide> { activeSlide } : new List<Slide>();
                }
                else
                {
                    slides = manager.GetSelectedSlidesForOperation();
                }

                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;

                foreach (Slide slide in slides)
                {
                    if (action == ActionType.Delete)
                    {
                        for (int i = slide.Shapes.Count; i >= 1; i--)
                        {
                            var shape = slide.Shapes[i];
                            if (shape.Name.StartsWith("StickyNote", StringComparison.OrdinalIgnoreCase))
                                shape.Delete();
                        }
                    }
                    else if (action == ActionType.MoveOff)
                    {
                        for (int i = 1; i <= slide.Shapes.Count; i++)
                        {
                            var shape = slide.Shapes[i];
                            if (shape.Name.StartsWith("StickyNote", StringComparison.OrdinalIgnoreCase))
                            {
                                if (shape.Left < slideWidth && shape.Left + shape.Width > 0 && shape.Top < slideHeight && shape.Top + shape.Height > 0)
                                {
                                    try 
                                    {
                                        shape.Tags.Add("OrigLeft", shape.Left.ToString(System.Globalization.CultureInfo.InvariantCulture));
                                        shape.Tags.Add("OrigTop", shape.Top.ToString(System.Globalization.CultureInfo.InvariantCulture));
                                    } 
                                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "StickyNoteManagerFeature.PrefixName");
                    }

                                    float dLeft = shape.Left;
                                    float dRight = slideWidth - (shape.Left + shape.Width);
                                    float dTop = shape.Top;
                                    float dBottom = slideHeight - (shape.Top + shape.Height);

                                    float min = Math.Min(Math.Min(dLeft, dRight), Math.Min(dTop, dBottom));

                                    if (min == dLeft) shape.Left = -shape.Width - 50;
                                    else if (min == dRight) shape.Left = slideWidth + 50;
                                    else if (min == dTop) shape.Top = -shape.Height - 50;
                                    else shape.Top = slideHeight + 50;
                                }
                            }
                        }
                    }
                    else if (action == ActionType.MoveOn)
                    {
                        for (int i = 1; i <= slide.Shapes.Count; i++)
                        {
                            var shape = slide.Shapes[i];
                            if (shape.Name.StartsWith("StickyNote", StringComparison.OrdinalIgnoreCase))
                            {
                                if (shape.Left >= slideWidth || shape.Left + shape.Width <= 0 || shape.Top >= slideHeight || shape.Top + shape.Height <= 0)
                                {
                                    bool moved = false;
                                    try 
                                    {
                                        string origLeftStr = "";
                                        string origTopStr = "";
                                        for (int t = 1; t <= shape.Tags.Count; t++)
                                        {
                                            if (shape.Tags.Name(t) == "OrigLeft") origLeftStr = shape.Tags.Value(t);
                                            else if (shape.Tags.Name(t) == "OrigTop") origTopStr = shape.Tags.Value(t);
                                        }

                                        if (!string.IsNullOrEmpty(origLeftStr) && !string.IsNullOrEmpty(origTopStr))
                                        {
                                            if (float.TryParse(origLeftStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float ol) &&
                                                float.TryParse(origTopStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float ot))
                                            {
                                                shape.Left = ol;
                                                shape.Top = ot;
                                                moved = true;
                                            }
                                        }
                                    } 
                                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "StickyNoteManagerFeature.PrefixName");
                    }

                                    if (!moved)
                                    {
                                        shape.Left = slideWidth - shape.Width - 50;
                                        shape.Top = 50;
                                    }
                                }
                            }
                        }
                    }
                    else if (action == ActionType.ConvertComments)
                    {
                        var comments = slide.Comments;
                        if (comments != null && comments.Count > 0)
                        {
                            float startTop = 50;
                            for (int i = comments.Count; i >= 1; i--)
                            {
                                var comment = comments[i];
                                string text = $"{comment.Author}:\n{comment.Text}";
                                
                                float cmToPt = 28.346f;
                                float width = 6.5f * cmToPt;
                                float height = 2.2f * cmToPt;
                                float marginCm = 0.2f * cmToPt;
                
                                float left = slideWidth - width - 50;
                                float top = startTop;
                
                                var shape = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeFoldedCorner, left, top, width, height);
                                shape.Name = "StickyNote_Comment_" + Guid.NewGuid().ToString().Substring(0, 8);
                                if (shape.Adjustments.Count > 0) shape.Adjustments[1] = 0.12f;
                                shape.TextFrame.MarginLeft = marginCm;
                                shape.TextFrame.MarginRight = marginCm;
                                shape.TextFrame.MarginTop = marginCm;
                                shape.TextFrame.MarginBottom = marginCm;
                                
                                shape.Fill.ForeColor.RGB = 0x88D2F8;
                                shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                                shape.Shadow.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

                                var tr = shape.TextFrame.TextRange;
                                tr.Text = text;
                                string themeFont = "Segoe UI";
                                try { themeFont = app.ActivePresentation.SlideMaster.Theme.ThemeFontScheme.MinorFont.Item(Microsoft.Office.Core.MsoFontLanguageIndex.msoThemeLatin).Name; } catch {}
                                tr.Font.Name = themeFont;
                                tr.Font.Size = 10;
                                tr.Font.Color.RGB = 0x000000;
                                startTop += height + 10;
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StickyNoteManagerFeature.Execute");
                return false;
            }
        }
    }
}
