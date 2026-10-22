using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnReplaceFont",
        Name = "Replace Formatting",
        SearchAliases = new[] { "Bulk Format Replacer" },
        Tooltip = "Replace matching formatting",
        IconData = "M13,12H15V16H19V18H15V22H13V18H9V16H13V12M5,2H19C20.1,2 21,2.9 21,4V10.1C20.39,10.04 19.77,10 19.12,10C17.43,10 15.89,10.65 14.72,11.71L13.72,10.71L12.72,11.71L11,10V7H13V5H11V3H9V5H7V3H5V20H11.1C11.04,20.61 11,21.23 11,21.88C11,22.25 11.03,22.63 11.08,23H5C3.9,23 3,22.1 3,21V4C3,2.9 3.9,2 5,2Z",
        Color = "#FBBF24",
        Description = "Replaces matching fonts, sizes, text colors, fills, and lines within a chosen slide scope.",
        DetailedHelpText = "### Replace Formatting\n\nChoose one or more format properties, define each find and replacement value, and apply them to the current slide, selected slides, or the entire presentation. Groups are traversed recursively. Matching is exact except for a 0.1 pt font-size tolerance. Invalid criteria are rejected before mutation; successful changes form one Undo entry, and an apply error triggers rollback of changes made by that run.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class BulkFormatReplacerFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, null);
        }

        public static bool Execute(PowerPointManager manager, ReplacementScope? initialScope)
        {
            if (manager == null) return false;
            UI.BulkFormatReplacerWindow window = null;
            try
            {
                window = new UI.BulkFormatReplacerWindow(manager, initialScope);
                var app = manager.GetApplication();
                if (app?.ActiveWindow != null)
                    new System.Windows.Interop.WindowInteropHelper(window).Owner = new IntPtr(app.HWND);
                window.Show();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "BulkFormatReplacerFeature.Execute");
                try { window?.Close(); }
                catch (Exception closeEx)
                {
                    ExceptionLogger.Log(
                        closeEx,
                        "BulkFormatReplacerFeature.Execute.FailureCleanup");
                }
                throw;
            }
        }

        public class FormatProps
        {
            public HashSet<string> Fonts = new HashSet<string>();
            public HashSet<float> Sizes = new HashSet<float>();
            public HashSet<int> TextColors = new HashSet<int>();
            public HashSet<int> FillColors = new HashSet<int>();
            public HashSet<int> LineColors = new HashSet<int>();
        }

        public static FormatProps GetPropsInSelectedSlides(PowerPointManager manager)
        {
            var props = new FormatProps();
            if (manager == null) return props;

            var slides = manager.GetSelectedSlidesForOperation();
            foreach (Slide slide in slides)
            {
                foreach (Shape shape in slide.Shapes)
                {
                    ProcessShapeProps(shape, props);
                }
            }
            return props;
        }

        private static void ProcessShapeProps(Shape shape, FormatProps props)
        {
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (Shape child in shape.GroupItems) ProcessShapeProps(child, props);
                return;
            }

            // Text props
            if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
            {
                var range = shape.TextFrame.TextRange;
                for (int i = 1; i <= range.Runs().Count; i++)
                {
                    var run = range.Runs(i);
                    try { props.Fonts.Add(run.Font.Name); }
                    catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "BulkFormatReplacerFeature.Scan.Font"); }
                    try { props.Sizes.Add(run.Font.Size); }
                    catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "BulkFormatReplacerFeature.Scan.FontSize"); }
                    try { props.TextColors.Add(run.Font.Color.RGB); }
                    catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "BulkFormatReplacerFeature.Scan.TextColor"); }
                }
            }

            // Shape props
            try { if (shape.Fill.Visible == Office.MsoTriState.msoTrue) props.FillColors.Add(shape.Fill.ForeColor.RGB); }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "BulkFormatReplacerFeature.Scan.FillColor"); }
            try { if (shape.Line.Visible == Office.MsoTriState.msoTrue) props.LineColors.Add(shape.Line.ForeColor.RGB); }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "BulkFormatReplacerFeature.Scan.LineColor"); }
        }

        public enum ReplacementScope
        {
            CurrentSlide,
            SelectedSlides,
            EntirePresentation
        }

        public static int ReplaceFormats(PowerPointManager manager,
            string findFont, string replaceFont,
            float? findSize, float? replaceSize,
            int? findTextColor, int? replaceTextColor,
            int? findFillColor, int? replaceFillColor,
            int? findLineColor, int? replaceLineColor,
            ReplacementScope scope = ReplacementScope.SelectedSlides)
        {
            if (manager == null) return -1;

            try
            {
                var slides = GetTargetSlides(manager, scope);
                var changes = new List<FormatChange>();

                foreach (var slide in slides)
                {
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        BuildChangesForShape(
                            slide.Shapes[i],
                            changes,
                            findFont,
                            replaceFont,
                            findSize,
                            replaceSize,
                            findTextColor,
                            replaceTextColor,
                            findFillColor,
                            replaceFillColor,
                            findLineColor,
                            replaceLineColor);
                    }
                }

                if (changes.Count == 0) return 0;
                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "BulkFormatReplacerFeature.StartNewUndoEntry"); }

                int applied = 0;
                try
                {
                    for (; applied < changes.Count; applied++)
                        changes[applied].Apply();
                    return changes.Count;
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "BulkFormatReplacerFeature.Apply");
                    for (int i = applied - 1; i >= 0; i--)
                    {
                        try { changes[i].Rollback(); }
                        catch (Exception rollbackEx)
                        {
                            ExceptionLogger.Log(rollbackEx, "BulkFormatReplacerFeature.Rollback");
                        }
                    }
                    return -1;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "BulkFormatReplacerFeature.ReplaceFormats");
                return -1;
            }
        }

        private static List<Slide> GetTargetSlides(PowerPointManager manager, ReplacementScope scope)
        {
            var slides = new List<Slide>();
            switch (scope)
            {
                case ReplacementScope.CurrentSlide:
                    var current = manager.GetCurrentSlide();
                    if (current != null) slides.Add(current);
                    break;
                case ReplacementScope.EntirePresentation:
                    var presentation = manager.GetApplication()?.ActivePresentation;
                    if (presentation != null)
                    {
                        for (int i = 1; i <= presentation.Slides.Count; i++)
                            slides.Add(presentation.Slides[i]);
                    }
                    break;
                default:
                    slides.AddRange(manager.GetSelectedSlidesForOperation());
                    break;
            }
            return slides;
        }

        private static void BuildChangesForShape(
            Shape shape,
            List<FormatChange> changes,
            string findFont, string replaceFont,
            float? findSize, float? replaceSize,
            int? findTextColor, int? replaceTextColor,
            int? findFillColor, int? replaceFillColor,
            int? findLineColor, int? replaceLineColor)
        {
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                for (int i = 1; i <= shape.GroupItems.Count; i++)
                    BuildChangesForShape(shape.GroupItems[i], changes, findFont, replaceFont, findSize, replaceSize, findTextColor, replaceTextColor, findFillColor, replaceFillColor, findLineColor, replaceLineColor);
                return;
            }

            if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
            {
                var range = shape.TextFrame.TextRange;
                for (int i = 1; i <= range.Runs().Count; i++)
                {
                    var run = range.Runs(i);
                    if (BulkFormatMatchPolicy.MatchesFont(run.Font.Name, findFont, replaceFont))
                    {
                        string oldValue = run.Font.Name;
                        var target = run;
                        changes.Add(new FormatChange(
                            () => target.Font.Name = replaceFont,
                            () => target.Font.Name = oldValue));
                    }
                    if (findSize.HasValue && replaceSize.HasValue &&
                        BulkFormatMatchPolicy.MatchesSize(run.Font.Size, findSize, replaceSize))
                    {
                        float oldValue = run.Font.Size;
                        var target = run;
                        changes.Add(new FormatChange(
                            () => target.Font.Size = replaceSize.Value,
                            () => target.Font.Size = oldValue));
                    }
                    if (findTextColor.HasValue && replaceTextColor.HasValue &&
                        BulkFormatMatchPolicy.MatchesColor(run.Font.Color.RGB, findTextColor, replaceTextColor))
                    {
                        int oldValue = run.Font.Color.RGB;
                        var target = run;
                        changes.Add(new FormatChange(
                            () => target.Font.Color.RGB = replaceTextColor.Value,
                            () => target.Font.Color.RGB = oldValue));
                    }
                }
            }

            if (findFillColor.HasValue && replaceFillColor.HasValue)
            {
                if (BulkFormatMatchPolicy.MatchesColor(
                        shape.Fill.ForeColor.RGB,
                        findFillColor,
                        replaceFillColor,
                        shape.Fill.Visible == Office.MsoTriState.msoTrue))
                {
                    int oldValue = shape.Fill.ForeColor.RGB;
                    changes.Add(new FormatChange(
                        () => shape.Fill.ForeColor.RGB = replaceFillColor.Value,
                        () => shape.Fill.ForeColor.RGB = oldValue));
                }
            }

            if (findLineColor.HasValue && replaceLineColor.HasValue)
            {
                if (BulkFormatMatchPolicy.MatchesColor(
                        shape.Line.ForeColor.RGB,
                        findLineColor,
                        replaceLineColor,
                        shape.Line.Visible == Office.MsoTriState.msoTrue))
                {
                    int oldValue = shape.Line.ForeColor.RGB;
                    changes.Add(new FormatChange(
                        () => shape.Line.ForeColor.RGB = replaceLineColor.Value,
                        () => shape.Line.ForeColor.RGB = oldValue));
                }
            }
        }

        private sealed class FormatChange
        {
            public FormatChange(Action apply, Action rollback)
            {
                Apply = apply;
                Rollback = rollback;
            }

            public Action Apply { get; }
            public Action Rollback { get; }
        }
    }
}
