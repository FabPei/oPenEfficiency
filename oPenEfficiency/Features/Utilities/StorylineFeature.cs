using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Interop;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Reviews slide headlines as a section-aware consulting storyline.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStoryline",
        Name = "Storyline",
        Tooltip = "Review the presentation storyline",
        Color = "#FBBF24",
        Description = "Reviews, filters, edits, and exports section-aware slide headlines with consulting-focused quality checks.",
        DetailedHelpText = "### Storyline\nOpen a section-aware narrative view of the presentation. Search or filter headlines, navigate through stable slide IDs, edit detected title shapes, assign review status, and use the context menu for slide actions. The quality check flags missing, duplicate, generic, overly long, question-only, and weak topic headlines and provides local action-title suggestions without sending presentation content to an external service. Copy or export the complete deck, the current section, headlines only, or issues only.",
        RequiredType = OfficeSelectionType.None)]
    public static class StorylineFeature
    {
        private const string ReviewStatusTag = "oPE_StorylineReviewStatus";
        private const string ExcludedTag = "oPE_StorylineExcluded";
        private static UI.Dialogs.StorylineDialog _window;

        private static readonly string[] GenericHeadlines =
        {
            "agenda", "analysis", "appendix", "background", "context", "discussion",
            "introduction", "next steps", "overview", "results", "summary", "update"
        };

        private static readonly string[] ActionWords =
        {
            "accelerate", "achieve", "can", "creates", "deliver", "drives", "enable",
            "improve", "increase", "is", "must", "need", "reduce", "requires", "should",
            "will", "are", "enables", "shows", "indicates", "leads", "supports"
        };

        public static bool Execute(PowerPointManager manager)
        {
            return Open(manager, false, false);
        }

        public static bool Open(PowerPointManager manager, bool currentSectionOnly, bool issuesOnly)
        {
            return OpenFiltered(manager, currentSectionOnly, issuesOnly ? "Issues" : "All");
        }

        public static bool OpenMissingTitles(PowerPointManager manager)
        {
            return OpenFiltered(manager, false, "Missing");
        }

        private static bool OpenFiltered(PowerPointManager manager, bool currentSectionOnly, string filter)
        {
            if (manager == null) return false;
            try
            {
                if (_window != null && _window.IsLoaded)
                {
                    _window.SetInitialView(currentSectionOnly, false);
                    _window.SetFilter(filter);
                    _window.RefreshStoryline();
                    _window.Activate();
                    return true;
                }

                _window = new UI.Dialogs.StorylineDialog(manager, currentSectionOnly, false);
                _window.SetFilter(filter);
                var app = manager.GetApplication();
                if (app != null) new WindowInteropHelper(_window).Owner = (IntPtr)app.HWND;
                _window.Closed += (s, e) => _window = null;
                _window.Show();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.OpenFiltered");
                var failedWindow = _window;
                _window = null;
                try { failedWindow?.Close(); }
                catch (Exception closeEx)
                {
                    ExceptionLogger.Log(closeEx, "StorylineFeature.OpenFiltered.FailureCleanup");
                }
                throw;
            }
        }

        public static void RefreshOpenWindow()
        {
            if (_window != null && _window.IsLoaded)
            {
                _window.RefreshStoryline();
                _window.Activate();
            }
        }

        public static void GoToSlide(PowerPointManager manager, int slideId)
        {
            try
            {
                var app = manager.GetApplication();
                var slide = app.ActivePresentation.Slides.FindBySlideID(slideId);
                if (slide != null) app.ActiveWindow.View.GotoSlide(slide.SlideIndex);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.GoToSlide");
            }
        }

        public static bool UpdateTitle(PowerPointManager manager, int slideId, int shapeId, string newTitle)
        {
            try
            {
                var app = manager.GetApplication();
                var slide = app.ActivePresentation.Slides.FindBySlideID(slideId);
                if (slide == null) return false;

                Shape titleShape = FindShapeById(slide, shapeId) ?? FindNativeTitle(slide);
                if (titleShape == null || titleShape.HasTextFrame == Office.MsoTriState.msoFalse) return false;

                TryStartUndo(app, "StorylineFeature.UpdateTitle.StartNewUndoEntry");
                titleShape.TextFrame.TextRange.Text = newTitle ?? string.Empty;
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.UpdateTitle");
                return false;
            }
        }

        public static bool SetReviewStatus(PowerPointManager manager, int slideId, string status)
        {
            try
            {
                var slide = manager.GetApplication().ActivePresentation.Slides.FindBySlideID(slideId);
                if (slide == null) return false;
                TryStartUndo(manager.GetApplication(), "StorylineFeature.SetReviewStatus.StartNewUndoEntry");
                if (string.IsNullOrWhiteSpace(status) || string.Equals(status, "None", StringComparison.OrdinalIgnoreCase))
                    TryDeleteTag(slide, ReviewStatusTag);
                else
                    slide.Tags.Add(ReviewStatusTag, status);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.SetReviewStatus");
                return false;
            }
        }

        public static bool ToggleExcluded(PowerPointManager manager, int slideId)
        {
            try
            {
                var slide = manager.GetApplication().ActivePresentation.Slides.FindBySlideID(slideId);
                if (slide == null) return false;
                bool excluded = IsTrue(SafeTag(slide, ExcludedTag));
                TryStartUndo(manager.GetApplication(), "StorylineFeature.ToggleExcluded.StartNewUndoEntry");
                if (excluded) TryDeleteTag(slide, ExcludedTag);
                else slide.Tags.Add(ExcludedTag, "True");
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.ToggleExcluded");
                return false;
            }
        }

        public static bool ToggleHidden(PowerPointManager manager, int slideId)
        {
            try
            {
                var app = manager.GetApplication();
                var slide = app.ActivePresentation.Slides.FindBySlideID(slideId);
                if (slide == null) return false;
                TryStartUndo(app, "StorylineFeature.ToggleHidden.StartNewUndoEntry");
                slide.SlideShowTransition.Hidden =
                    slide.SlideShowTransition.Hidden == Office.MsoTriState.msoTrue
                        ? Office.MsoTriState.msoFalse
                        : Office.MsoTriState.msoTrue;
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.ToggleHidden");
                return false;
            }
        }

        public static bool InsertSlideAfter(PowerPointManager manager, int slideId)
        {
            try
            {
                var app = manager.GetApplication();
                var presentation = app.ActivePresentation;
                var source = presentation.Slides.FindBySlideID(slideId);
                if (source == null) return false;
                TryStartUndo(app, "StorylineFeature.InsertSlideAfter.StartNewUndoEntry");
                Slide inserted = presentation.Slides.Add(source.SlideIndex + 1, PpSlideLayout.ppLayoutText);
                try { inserted.Shapes.Title.TextFrame.TextRange.Text = "New action title"; }
                catch (Exception ex) { ReportOptionalStorylineFailure(ex, "StorylineFeature.InsertSlideAfter.Title"); }
                app.ActiveWindow.View.GotoSlide(inserted.SlideIndex);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.InsertSlideAfter");
                return false;
            }
        }

        public static bool MoveSlide(PowerPointManager manager, int slideId, int delta)
        {
            try
            {
                var app = manager.GetApplication();
                var presentation = app.ActivePresentation;
                var slide = presentation.Slides.FindBySlideID(slideId);
                if (slide == null) return false;
                int target = Math.Max(1, Math.Min(presentation.Slides.Count, slide.SlideIndex + delta));
                if (target == slide.SlideIndex) return false;
                TryStartUndo(app, "StorylineFeature.MoveSlide.StartNewUndoEntry");
                slide.MoveTo(target);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.MoveSlide");
                return false;
            }
        }

        public static bool AssignToSection(PowerPointManager manager, int slideId, string sectionName)
        {
            try
            {
                var app = manager.GetApplication();
                var presentation = app.ActivePresentation;
                var slide = presentation.Slides.FindBySlideID(slideId);
                if (slide == null || string.IsNullOrWhiteSpace(sectionName)) return false;

                int sectionIndex = FindSectionIndex(presentation, sectionName);
                if (sectionIndex < 1) return false;
                dynamic sections = presentation.SectionProperties;
                int firstIndex = sections.FirstSlideIndex(sectionIndex);
                int count = sections.SlidesCount(sectionIndex);
                int target = Math.Min(presentation.Slides.Count, firstIndex + Math.Max(0, count - 1));
                TryStartUndo(app, "StorylineFeature.AssignToSection.StartNewUndoEntry");
                slide.MoveTo(target);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.AssignToSection");
                return false;
            }
        }

        public static IList<string> GetSectionNames(PowerPointManager manager)
        {
            var result = new List<string>();
            try
            {
                var presentation = manager.GetApplication().ActivePresentation;
                for (int i = 1; i <= presentation.SectionProperties.Count; i++)
                {
                    string name = presentation.SectionProperties.Name(i);
                    if (!string.IsNullOrWhiteSpace(name)) result.Add(name.Trim());
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.GetSectionNames");
            }
            return result;
        }

        public sealed class StorylineInfo
        {
            public int SlideNumber { get; set; }
            public int SlideId { get; set; }
            public int TitleShapeId { get; set; }
            public string SectionName { get; set; }
            public string TitleText { get; set; }
            public string ReviewStatus { get; set; }
            public string FrameworkRole { get; set; }
            public string SuggestedTitle { get; set; }
            public bool CanEdit { get; set; }
            public bool IsHidden { get; set; }
            public bool IsExcluded { get; set; }
            public List<string> Issues { get; } = new List<string>();
            public bool HasIssues => Issues.Count > 0;
            public bool IsMissing => string.Equals(TitleText, "[No title found]", StringComparison.Ordinal);
            public string IssueSummary => Issues.Count == 0 ? "No issues detected" : string.Join(Environment.NewLine, Issues);
            public string StatusLabel => string.IsNullOrWhiteSpace(ReviewStatus) ? "Unreviewed" : ReviewStatus;
        }

        public static List<StorylineInfo> ExtractStoryline(PowerPointManager manager)
        {
            var result = new List<StorylineInfo>();
            try
            {
                var presentation = manager.GetApplication().ActivePresentation;
                if (presentation == null) return result;

                foreach (Slide slide in presentation.Slides)
                    result.Add(ExtractTitleFromSlide(presentation, slide));

                AnalyzeStoryline(result);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.ExtractStoryline");
            }
            return result;
        }

        public static string BuildOutline(IEnumerable<StorylineInfo> items, bool includeNumbers, bool groupBySection, bool issuesOnly)
        {
            var selected = (items ?? Enumerable.Empty<StorylineInfo>())
                .Where(item => !item.IsExcluded && (!issuesOnly || item.HasIssues))
                .ToList();
            var builder = new StringBuilder();
            string lastSection = null;
            foreach (var item in selected)
            {
                if (groupBySection && !string.Equals(lastSection, item.SectionName, StringComparison.Ordinal))
                {
                    if (builder.Length > 0) builder.AppendLine();
                    builder.AppendLine(item.SectionName);
                    lastSection = item.SectionName;
                }
                builder.Append(includeNumbers ? item.SlideNumber + ": " : string.Empty);
                builder.Append(item.TitleText);
                if (issuesOnly && item.HasIssues) builder.Append(" — " + string.Join("; ", item.Issues));
                builder.AppendLine();
            }
            return builder.ToString();
        }

        public static string SuggestActionTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title) || title.StartsWith("[", StringComparison.Ordinal))
                return "Add a concise action title that states the slide's main conclusion";

            string cleaned = Regex.Replace(title.Trim(), @"\s+", " ").TrimEnd(':', '.', ';');
            if (cleaned.EndsWith("?", StringComparison.Ordinal))
                return "Answer the question directly: " + cleaned.TrimEnd('?');
            if (GenericHeadlines.Contains(cleaned, StringComparer.OrdinalIgnoreCase))
                return cleaned + " — state the key implication for the client";
            if (!ContainsActionWord(cleaned))
                return cleaned + " — clarify the key implication";
            return cleaned;
        }

        private static StorylineInfo ExtractTitleFromSlide(Presentation presentation, Slide slide)
        {
            try
            {
                string[] emptyPrompts =
                {
                    "Click to edit Master title style", "Click to edit title",
                    "Klicken Sie, um den Titel zu bearbeiten", "Titel durch Klicken hinzufügen",
                    "Click to edit Master subtitle style", "Untertitel durch Klicken hinzufügen",
                    "Title", "Titel", "Click to add title", "Klicken Sie, um einen Titel hinzuzufügen"
                };

                Shape candidate = FindNativeTitle(slide);
                if (candidate == null || !HasActualContent(candidate, emptyPrompts))
                    candidate = FindBestTitleCandidate(presentation, slide, emptyPrompts);

                string title = candidate == null
                    ? "[No title found]"
                    : candidate.TextFrame.TextRange.Text.Trim();

                return new StorylineInfo
                {
                    SlideNumber = slide.SlideNumber,
                    SlideId = slide.SlideID,
                    TitleShapeId = candidate?.Id ?? 0,
                    SectionName = GetSectionName(presentation, slide.SlideIndex),
                    TitleText = title,
                    CanEdit = candidate != null,
                    IsHidden = slide.SlideShowTransition.Hidden == Office.MsoTriState.msoTrue,
                    IsExcluded = IsTrue(SafeTag(slide, ExcludedTag)),
                    ReviewStatus = SafeTag(slide, ReviewStatusTag),
                    FrameworkRole = InferFrameworkRole(title),
                    SuggestedTitle = SuggestActionTitle(title)
                };
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.ExtractTitleFromSlide");
                return new StorylineInfo
                {
                    SlideNumber = slide.SlideNumber,
                    SlideId = slide.SlideID,
                    SectionName = "Unsectioned",
                    TitleText = "[Error]",
                    ReviewStatus = string.Empty,
                    FrameworkRole = "Unclassified",
                    SuggestedTitle = string.Empty
                };
            }
        }

        private static Shape FindNativeTitle(Slide slide)
        {
            Shape titleShape = null;
            try { titleShape = slide.Shapes.Title; }
            catch (Exception ex) { ReportOptionalStorylineFailure(ex, "StorylineFeature.FindTitleShape.BuiltInTitle"); }
            if (titleShape != null) return titleShape;

            foreach (Shape shape in slide.Shapes)
            {
                try
                {
                    if (shape.Type != Office.MsoShapeType.msoPlaceholder) continue;
                    var type = shape.PlaceholderFormat.Type;
                    if (type == PpPlaceholderType.ppPlaceholderTitle ||
                        type == PpPlaceholderType.ppPlaceholderCenterTitle ||
                        type == PpPlaceholderType.ppPlaceholderHeader)
                        return shape;
                }
                catch (Exception ex) { ReportOptionalStorylineFailure(ex, "StorylineFeature.FindTitleShape.Candidate"); }
            }
            return null;
        }

        private static Shape FindBestTitleCandidate(Presentation presentation, Slide slide, string[] emptyPrompts)
        {
            Shape bestCandidate = null;
            float bestTop = float.MaxValue;
            float slideHeight = 540;
            try { slideHeight = presentation.PageSetup.SlideHeight; }
            catch (Exception ex) { ReportOptionalStorylineFailure(ex, "StorylineFeature.FindHeadlineCandidate.SlideHeight"); }

            foreach (Shape shape in slide.Shapes)
            {
                try
                {
                    if (!HasActualContent(shape, emptyPrompts)) continue;
                    if (shape.Top < slideHeight * 0.5f && shape.Top < bestTop)
                    {
                        bestTop = shape.Top;
                        bestCandidate = shape;
                    }
                }
                catch (Exception ex) { ReportOptionalStorylineFailure(ex, "StorylineFeature.FindHeadlineCandidate.Shape"); }
            }
            return bestCandidate;
        }

        private static void AnalyzeStoryline(List<StorylineInfo> items)
        {
            var duplicateTitles = items
                .Where(item => !item.IsMissing && !item.IsExcluded)
                .GroupBy(item => Normalize(item.TitleText), StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            bool hasExecutiveSummary = items.Any(item =>
                item.TitleText.IndexOf("executive summary", StringComparison.OrdinalIgnoreCase) >= 0 ||
                item.TitleText.IndexOf("management summary", StringComparison.OrdinalIgnoreCase) >= 0);

            foreach (var item in items)
            {
                if (item.IsExcluded) continue;
                string normalized = Normalize(item.TitleText);
                if (item.IsMissing) item.Issues.Add("Missing headline");
                else
                {
                    if (duplicateTitles.Contains(normalized)) item.Issues.Add("Duplicate headline");
                    if (GenericHeadlines.Contains(normalized, StringComparer.OrdinalIgnoreCase)) item.Issues.Add("Generic topic headline");
                    if (item.TitleText.Length > 120) item.Issues.Add("Headline exceeds 120 characters");
                    if (item.TitleText.TrimEnd().EndsWith("?", StringComparison.Ordinal)) item.Issues.Add("Question headline does not state the answer");
                    if (!ContainsActionWord(item.TitleText) && item.TitleText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length <= 7)
                        item.Issues.Add("Likely topic headline; state the key implication");
                }
                if (item.IsHidden) item.Issues.Add("Slide is hidden");
            }

            if (items.Count >= 12 && !hasExecutiveSummary && items.Count > 0)
                items[0].Issues.Add("No Executive Summary headline detected");
        }

        private static string InferFrameworkRole(string title)
        {
            string value = Normalize(title);
            if (value.Contains("challenge") || value.Contains("however") || value.Contains("gap") || value.Contains("risk"))
                return "Complication";
            if (value.EndsWith("?") || value.Contains("question"))
                return "Question";
            if (value.Contains("recommend") || value.Contains("should") || value.Contains("will") || value.Contains("action"))
                return "Answer / Action";
            return "Situation / Evidence";
        }

        private static bool ContainsActionWord(string title)
        {
            string padded = " " + Normalize(title) + " ";
            return ActionWords.Any(word => padded.Contains(" " + word + " "));
        }

        private static string Normalize(string value)
        {
            return Regex.Replace((value ?? string.Empty).Trim().ToLowerInvariant(), @"[^\p{L}\p{Nd}\?]+", " ").Trim();
        }

        private static string GetSectionName(Presentation presentation, int slideIndex)
        {
            try
            {
                dynamic sections = presentation.SectionProperties;
                for (int i = 1; i <= presentation.SectionProperties.Count; i++)
                {
                    int first = sections.FirstSlideIndex(i);
                    int count = sections.SlidesCount(i);
                    if (slideIndex >= first && slideIndex < first + count)
                        return presentation.SectionProperties.Name(i);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineFeature.GetSectionName");
            }
            return "Unsectioned";
        }

        private static int FindSectionIndex(Presentation presentation, string name)
        {
            for (int i = 1; i <= presentation.SectionProperties.Count; i++)
            {
                if (string.Equals(presentation.SectionProperties.Name(i), name, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return -1;
        }

        private static Shape FindShapeById(Slide slide, int shapeId)
        {
            if (shapeId <= 0) return null;
            foreach (Shape shape in slide.Shapes)
            {
                try { if (shape.Id == shapeId) return shape; }
                catch (Exception ex) { ReportOptionalStorylineFailure(ex, "StorylineFeature.FindShapeById"); }
            }
            return null;
        }

        private static bool HasActualContent(Shape shape, IEnumerable<string> prompts)
        {
            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoFalse ||
                    shape.TextFrame.HasText == Office.MsoTriState.msoFalse) return false;
                string text = shape.TextFrame.TextRange.Text.Trim();
                return text.Length > 0 && !prompts.Any(prompt =>
                    string.Equals(text, prompt, StringComparison.OrdinalIgnoreCase));
            }
            catch { return false; }
        }

        private static string SafeTag(Slide slide, string name)
        {
            try { return slide.Tags[name] ?? string.Empty; } catch { return string.Empty; }
        }

        private static void TryDeleteTag(Slide slide, string name)
        {
            try { slide.Tags.Delete(name); }
            catch (Exception ex) { ReportOptionalStorylineFailure(ex, "StorylineFeature.TryDeleteTag"); }
        }

        private static void ReportOptionalStorylineFailure(Exception exception, string context)
        {
            if (!ExceptionPolicy.IsExpectedComFailure(exception))
                ExceptionLogger.Log(exception, context);
            // Missing optional title, tag, or stale shape state falls back to the next candidate.
        }

        private static bool IsTrue(string value)
        {
            return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "1", StringComparison.OrdinalIgnoreCase);
        }

        private static void TryStartUndo(Microsoft.Office.Interop.PowerPoint.Application app, string context)
        {
            try { app?.StartNewUndoEntry(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, context); }
        }
    }
}
