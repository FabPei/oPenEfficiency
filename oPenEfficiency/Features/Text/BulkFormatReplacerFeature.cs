using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnReplaceFont",
        Name = "Bulk Format Replacer",
        Tooltip = "Bulk format replacer",
        IconData = "M13,12H15V16H19V18H15V22H13V18H9V16H13V12M5,2H19C20.1,2 21,2.9 21,4V10.1C20.39,10.04 19.77,10 19.12,10C17.43,10 15.89,10.65 14.72,11.71L13.72,10.71L12.72,11.71L11,10V7H13V5H11V3H9V5H7V3H5V20H11.1C11.04,20.61 11,21.23 11,21.88C11,22.25 11.03,22.63 11.08,23H5C3.9,23 3,22.1 3,21V4C3,2.9 3.9,2 5,2Z",
        Color = "#FBBF24",
        Description = "Batch replace fonts, sizes, and colors (Text, Fill, Line) across all selected slides.",
        DetailedHelpText = "### Bulk Format Replacer\nFinds all text runs in the presentation matching a set of font and color criteria and replaces them with a new format in one batch operation.",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class BulkFormatReplacerFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var win = new UI.BulkFormatReplacerWindow(manager);
            win.Show();
            return true;
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
                    try { props.Fonts.Add(run.Font.Name); } catch { }
                    try { props.Sizes.Add(run.Font.Size); } catch { }
                    try { props.TextColors.Add(run.Font.Color.RGB); } catch { }
                }
            }

            // Shape props
            try { if (shape.Fill.Visible == Office.MsoTriState.msoTrue) props.FillColors.Add(shape.Fill.ForeColor.RGB); } catch { }
            try { if (shape.Line.Visible == Office.MsoTriState.msoTrue) props.LineColors.Add(shape.Line.ForeColor.RGB); } catch { }
        }

        public enum ReplacementScope
        {
            CurrentSlide,
            SelectedSlides,
            EntirePresentation
        }

        public static void ReplaceFormats(PowerPointManager manager, 
            string findFont, string replaceFont,
            float? findSize, float? replaceSize,
            int? findTextColor, int? replaceTextColor,
            int? findFillColor, int? replaceFillColor,
            int? findLineColor, int? replaceLineColor,
            ReplacementScope scope = ReplacementScope.SelectedSlides)
        {
            if (manager == null) return;
            
            IEnumerable<Slide> slides;
            var app = manager.GetApplication();

            switch (scope)
            {
                case ReplacementScope.CurrentSlide:
                    var current = manager.GetCurrentSlide();
                    slides = current != null ? new List<Slide> { current } : new List<Slide>();
                    break;
                case ReplacementScope.EntirePresentation:
                    var pres = app.ActivePresentation;
                    var allSlides = new List<Slide>();
                    if (pres != null)
                    {
                        foreach (Slide s in pres.Slides) allSlides.Add(s);
                    }
                    slides = allSlides;
                    break;
                case ReplacementScope.SelectedSlides:
                default:
                    slides = manager.GetSelectedSlidesForOperation();
                    break;
            }

            foreach (Slide slide in slides)
            {
                foreach (Shape shape in slide.Shapes)
                {
                    ApplyReplacementToShape(shape, 
                        findFont, replaceFont, 
                        findSize, replaceSize, 
                        findTextColor, replaceTextColor,
                        findFillColor, replaceFillColor,
                        findLineColor, replaceLineColor);
                }
            }
        }

        private static void ApplyReplacementToShape(Shape shape,
            string findFont, string replaceFont,
            float? findSize, float? replaceSize,
            int? findTextColor, int? replaceTextColor,
            int? findFillColor, int? replaceFillColor,
            int? findLineColor, int? replaceLineColor)
        {
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (Shape child in shape.GroupItems) 
                    ApplyReplacementToShape(child, findFont, replaceFont, findSize, replaceSize, findTextColor, replaceTextColor, findFillColor, replaceFillColor, findLineColor, replaceLineColor);
                return;
            }

            // Replace Text Formats
            if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
            {
                var range = shape.TextFrame.TextRange;
                for (int i = 1; i <= range.Runs().Count; i++)
                {
                    var run = range.Runs(i);
                    if (!string.IsNullOrEmpty(findFont) && run.Font.Name == findFont) run.Font.Name = replaceFont;
                    if (findSize.HasValue && Math.Abs(run.Font.Size - findSize.Value) < 0.1) run.Font.Size = replaceSize.Value;
                    if (findTextColor.HasValue && run.Font.Color.RGB == findTextColor.Value) run.Font.Color.RGB = replaceTextColor.Value;
                }
            }

            // Replace Fill Color
            if (findFillColor.HasValue && replaceFillColor.HasValue)
            {
                if (shape.Fill.Visible == Office.MsoTriState.msoTrue && shape.Fill.ForeColor.RGB == findFillColor.Value)
                    shape.Fill.ForeColor.RGB = replaceFillColor.Value;
            }

            // Replace Line Color
            if (findLineColor.HasValue && replaceLineColor.HasValue)
            {
                if (shape.Line.Visible == Office.MsoTriState.msoTrue && shape.Line.ForeColor.RGB == findLineColor.Value)
                    shape.Line.ForeColor.RGB = replaceLineColor.Value;
            }
        }
    }
}
