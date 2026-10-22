using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnSlideGuidelines",
        Name = "Create Slide Guide Grid",
        SearchAliases = new[] { "Slide Guidelines" },
        Tooltip = "Create responsive layout guides",
        IconData = "M3,3H21V21H3V3M7,7V19H19V7H7M9,9H17V11H9V9M9,13H17V15H9V13M9,17H17V19H9V17Z",
        Color = "#F43F5E",
        Description = "Creates responsive tagged guide shapes on the current slide and replaces only guides owned by openEfficiency.",
        DetailedHelpText = "### Create Slide Guide Grid\nCreates responsive line-shape guides for the current slide size. A normal click applies a two-column layout; right-click offers three, four, five, 1:2, and rule-of-thirds layouts plus explicit removal. Applying a layout replaces only shapes tagged as openEfficiency guidelines, preserving similarly named user shapes. These are slide shapes rather than PowerPoint's native guides.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class SlideGuidelinesFeature
    {
        private const double PointsPerCm = 28.3464567;
        private const string GuidelineTag = "oPE_Guideline";
        private const string GuidelineSetTag = "oPE_GuidelineSet";

        public enum GuidelineLayout { Default, ThreeColumn, FourColumn, FiveColumn, OneTwoColumn, RuleOfThirds }

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, GuidelineLayout.Default);
        }

        public static bool Execute(PowerPointManager manager, GuidelineLayout layout)
        {
            if (!TryGetCurrentSlide(manager, out var app, out var slide, out var presentation)) return false;
            var created = new List<Shape>();
            string setId = Guid.NewGuid().ToString("N");
            try
            {
                try { app.StartNewUndoEntry(); }
                catch (Exception undoEx) { ExceptionPolicy.ReportOptionalComFailure(undoEx, "SlideGuidelinesFeature.Execute.StartUndoEntry"); }
                CreateGuidelines(slide, presentation, layout, setId, created);
                RemoveExistingGuidelines(slide, setId);
                return true;
            }
            catch (Exception ex)
            {
                for (int i = created.Count - 1; i >= 0; i--)
                {
                    try { created[i].Delete(); }
                    catch (Exception rollbackEx) { ExceptionPolicy.ReportOptionalComFailure(rollbackEx, "SlideGuidelinesFeature.Execute.Rollback"); }
                }
                ExceptionLogger.Log(ex, "SlideGuidelinesFeature.Execute");
                System.Windows.MessageBox.Show("The guideline layout could not be created. Existing guidelines were preserved.",
                    "Create Slide Guide Grid", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return false;
            }
        }

        public static bool RemoveGuidelines(PowerPointManager manager)
        {
            if (!TryGetCurrentSlide(manager, out var app, out var slide, out _)) return false;
            try
            {
                try { app.StartNewUndoEntry(); }
                catch (Exception undoEx) { ExceptionPolicy.ReportOptionalComFailure(undoEx, "SlideGuidelinesFeature.Remove.StartUndoEntry"); }
                return RemoveExistingGuidelines(slide, null) > 0;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SlideGuidelinesFeature.RemoveGuidelines");
                return false;
            }
        }

        private static bool TryGetCurrentSlide(PowerPointManager manager, out Application app, out Slide slide, out Presentation presentation)
        {
            app = null;
            slide = null;
            presentation = null;
            try
            {
                app = manager?.GetApplication();
                presentation = app?.ActivePresentation;
                if (app?.ActiveWindow == null || presentation == null) return false;
                slide = app.ActiveWindow.View.Slide as Slide;
                return slide != null;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SlideGuidelinesFeature.TryGetCurrentSlide");
                return false;
            }
        }

        private static void CreateGuidelines(Slide slide, Presentation presentation, GuidelineLayout layout, string setId, List<Shape> created)
        {
            float widthCm = (float)(presentation.PageSetup.SlideWidth / PointsPerCm);
            float heightCm = (float)(presentation.PageSetup.SlideHeight / PointsPerCm);
            if (layout == GuidelineLayout.RuleOfThirds)
            {
                AddVertical(slide, presentation, widthCm / 3f, "Third Left", setId, created);
                AddVertical(slide, presentation, widthCm * 2f / 3f, "Third Right", setId, created);
                AddHorizontal(slide, presentation, heightCm / 3f, "Third Top", setId, created);
                AddHorizontal(slide, presentation, heightCm * 2f / 3f, "Third Bottom", setId, created);
                return;
            }

            float outerHorizontalMargin = Math.Min(1.25f, heightCm * 0.1f);
            float titleBaseline = Math.Min(3.75f, heightCm * 0.25f);
            AddHorizontal(slide, presentation, outerHorizontalMargin, "Top Margin", setId, created);
            AddHorizontal(slide, presentation, titleBaseline, "Title Baseline", setId, created);
            AddHorizontal(slide, presentation, heightCm / 2f, "Vertical Center", setId, created);
            AddHorizontal(slide, presentation, heightCm - outerHorizontalMargin, "Bottom Margin", setId, created);

            if (layout == GuidelineLayout.OneTwoColumn) AddRatioColumns(slide, presentation, widthCm, setId, created);
            else AddEqualColumns(slide, presentation, widthCm,
                layout == GuidelineLayout.ThreeColumn ? 3 : layout == GuidelineLayout.FourColumn ? 4 : layout == GuidelineLayout.FiveColumn ? 5 : 2,
                setId, created);
        }

        private static void AddEqualColumns(Slide slide, Presentation presentation, float widthCm, int columns, string setId, List<Shape> created)
        {
            float margin = Math.Min(0.6f, widthCm * 0.04f);
            float contentWidth = widthCm - 2f * margin;
            float gutter = Math.Min(1.2f, contentWidth / (columns * 4f));
            AddVertical(slide, presentation, margin, "Left Margin", setId, created);
            for (int i = 1; i < columns; i++)
            {
                float boundary = margin + contentWidth * i / columns;
                AddVertical(slide, presentation, boundary - gutter / 2f, $"Column {i} Right", setId, created);
                AddVertical(slide, presentation, boundary + gutter / 2f, $"Column {i + 1} Left", setId, created);
            }
            AddVertical(slide, presentation, widthCm - margin, "Right Margin", setId, created);
        }

        private static void AddRatioColumns(Slide slide, Presentation presentation, float widthCm, string setId, List<Shape> created)
        {
            float margin = Math.Min(0.6f, widthCm * 0.04f);
            float contentWidth = widthCm - 2f * margin;
            float gutter = Math.Min(1.2f, contentWidth / 12f);
            float boundary = margin + contentWidth / 3f;
            AddVertical(slide, presentation, margin, "Left Margin", setId, created);
            AddVertical(slide, presentation, boundary - gutter / 2f, "Narrow Column Right", setId, created);
            AddVertical(slide, presentation, boundary + gutter / 2f, "Wide Column Left", setId, created);
            AddVertical(slide, presentation, widthCm - margin, "Right Margin", setId, created);
        }

        private static void AddHorizontal(Slide slide, Presentation presentation, float positionCm, string label, string setId, List<Shape> created)
        {
            float position = (float)(positionCm * PointsPerCm);
            var line = slide.Shapes.AddLine(0, position, presentation.PageSetup.SlideWidth, position);
            created.Add(line);
            Configure(line, "H", label, setId, 0x0000FF);
        }

        private static void AddVertical(Slide slide, Presentation presentation, float positionCm, string label, string setId, List<Shape> created)
        {
            float position = (float)(positionCm * PointsPerCm);
            var line = slide.Shapes.AddLine(position, 0, position, presentation.PageSetup.SlideHeight);
            created.Add(line);
            Configure(line, "V", label, setId, 0xFF0000);
        }

        private static void Configure(Shape line, string axis, string label, string setId, int color)
        {
            line.Name = $"oPE_Guide_{axis}_{label}_{setId.Substring(0, 8)}";
            line.AlternativeText = GuidelineTag;
            line.Tags.Add(GuidelineTag, "true");
            line.Tags.Add(GuidelineSetTag, setId);
            line.Line.ForeColor.RGB = color;
            line.Line.Weight = 1f;
            line.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;
            line.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
            try
            {
                ((dynamic)line).Locked = Office.MsoTriState.msoTrue;
                line.Tags.Add("oPE_Locked", "true");
            }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "SlideGuidelinesFeature.Configure.Lock"); }
        }

        private static int RemoveExistingGuidelines(Slide slide, string exceptSetId)
        {
            int removed = 0;
            for (int i = slide.Shapes.Count; i >= 1; i--)
            {
                var shape = slide.Shapes[i];
                if (!IsGuideline(shape)) continue;
                string setId = GetTag(shape, GuidelineSetTag);
                if (!string.IsNullOrEmpty(exceptSetId) && string.Equals(setId, exceptSetId, StringComparison.OrdinalIgnoreCase)) continue;
                shape.Delete();
                removed++;
            }
            return removed;
        }

        private static bool IsGuideline(Shape shape)
        {
            try
            {
                return string.Equals(shape.AlternativeText, GuidelineTag, StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(GetTag(shape, GuidelineTag), "true", StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        private static string GetTag(Shape shape, string name)
        {
            try { return shape.Tags[name]; }
            catch { return null; }
        }
    }
}
