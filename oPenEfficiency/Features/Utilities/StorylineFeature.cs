using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Consolidates slide titles into a narrative flow view (Storyline).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStoryline",
        Name = "Storyline",
        Tooltip = "Storyline",
        Color = "#FBBF24",
        Description = "Consolidates all slide titles into a single narrative flow view to check the 'red thread' of your story.",
        Keywords = "storyline, narrative, titles, flow, outline, overview",
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class StorylineFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var win = new UI.Dialogs.StorylineDialog(manager);
            win.Show();
            return true;
        }

        public static void GoToSlide(PowerPointManager manager, int slideNumber)
        {
            try
            {
                var app = manager.GetApplication();
                app.ActiveWindow.View.GotoSlide(slideNumber);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.GoToSlide");
            }
        }

        public static bool UpdateTitle(PowerPointManager manager, int slideNumber, string newTitle)
        {
            try
            {
                var app = manager.GetApplication();
                var slide = app.ActivePresentation.Slides[slideNumber];
                
                // Try to find the title shape again to be safe
                Shape titleShape = null;
                try { titleShape = slide.Shapes.Title; } catch { }

                if (titleShape == null)
                {
                    foreach (Shape s in slide.Shapes)
                    {
                        if (s.Type == Office.MsoShapeType.msoPlaceholder)
                        {
                            var pt = s.PlaceholderFormat.Type;
                            if (pt == PpPlaceholderType.ppPlaceholderTitle || pt == PpPlaceholderType.ppPlaceholderCenterTitle)
                            {
                                titleShape = s;
                                break;
                            }
                        }
                    }
                }

                if (titleShape != null)
                {
                    titleShape.TextFrame.TextRange.Text = newTitle;
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.UpdateTitle");
                return false;
            }
        }

        public class StorylineInfo
        {
            public int SlideNumber { get; set; }
            public string TitleText { get; set; }
            public Shape TitleShape { get; set; }
        }

        public static List<StorylineInfo> ExtractStoryline(PowerPointManager manager)
        {
            var result = new List<StorylineInfo>();
            try
            {
                var pres = manager.GetApplication().ActivePresentation;
                if (pres == null) return result;

                foreach (Slide slide in pres.Slides)
                {
                    result.Add(ExtractTitleFromSlide(slide));
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.ExtractStoryline");
            }
            return result;
        }

        private static StorylineInfo ExtractTitleFromSlide(Slide slide)
        {
            try
            {
                string[] emptyPrompts = { 
                    "Click to edit Master title style", 
                    "Click to edit title", 
                    "Klicken Sie, um den Titel zu bearbeiten", 
                    "Titel durch Klicken hinzufügen",
                    "Click to edit Master subtitle style",
                    "Untertitel durch Klicken hinzufügen",
                    "Title",
                    "Titel",
                    "Click to add title",
                    "Klicken Sie, um einen Titel hinzuzufügen"
                };

                Shape titleShape = null;
                try { titleShape = slide.Shapes.Title; } catch { }

                if (titleShape != null && HasActualContent(titleShape, emptyPrompts))
                {
                    return new StorylineInfo { 
                        SlideNumber = slide.SlideNumber, 
                        TitleText = titleShape.TextFrame.TextRange.Text.Trim(), 
                        TitleShape = titleShape 
                    };
                }

                Shape bestCandidate = null;
                float bestTop = float.MaxValue;
                bool foundExplicitTitlePlaceholder = false;

                foreach (Shape s in slide.Shapes)
                {
                    try
                    {
                        if (s.HasTextFrame == Office.MsoTriState.msoFalse) continue;
                        if (s.TextFrame.HasText == Office.MsoTriState.msoFalse) continue;
                        if (!HasActualContent(s, emptyPrompts)) continue;

                        bool isTitlePlaceholder = false;
                        if (s.Type == Office.MsoShapeType.msoPlaceholder)
                        {
                            var pt = s.PlaceholderFormat.Type;
                            if (pt == PpPlaceholderType.ppPlaceholderTitle || 
                                pt == PpPlaceholderType.ppPlaceholderCenterTitle ||
                                pt == PpPlaceholderType.ppPlaceholderHeader)
                            {
                                isTitlePlaceholder = true;
                            }
                        }

                        if (isTitlePlaceholder)
                        {
                            if (!foundExplicitTitlePlaceholder || s.Top < bestTop)
                            {
                                bestTop = s.Top;
                                bestCandidate = s;
                                foundExplicitTitlePlaceholder = true;
                            }
                        }
                        else if (!foundExplicitTitlePlaceholder)
                        {
                            float slideHeight = 540;
                            try { slideHeight = ((Presentation)slide.Parent).PageSetup.SlideHeight; } catch { }
                            
                            if (s.Top < slideHeight * 0.5f && s.Top < bestTop)
                            {
                                bestTop = s.Top;
                                bestCandidate = s;
                            }
                        }
                    }
                    catch { }
                }

                if (bestCandidate != null)
                {
                    return new StorylineInfo { 
                        SlideNumber = slide.SlideNumber, 
                        TitleText = bestCandidate.TextFrame.TextRange.Text.Trim(), 
                        TitleShape = bestCandidate 
                    };
                }

                return new StorylineInfo { SlideNumber = slide.SlideNumber, TitleText = "[No title found]", TitleShape = null };
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.ExtractTitleFromSlide");
                return new StorylineInfo { SlideNumber = slide.SlideNumber, TitleText = "[Error]", TitleShape = null };
            }
        }

        private static bool HasActualContent(Shape s, string[] prompts)
        {
            try
            {
                if (s.HasTextFrame == Office.MsoTriState.msoFalse || s.TextFrame.HasText == Office.MsoTriState.msoFalse) return false;
                string text = s.TextFrame.TextRange.Text.Trim();
                if (text.Length < 1) return false;
                
                foreach (var p in prompts)
                {
                    if (string.Equals(text, p, StringComparison.OrdinalIgnoreCase)) return false;
                }
                
                return true;
            }
            catch { return false; }
        }
    }
}
