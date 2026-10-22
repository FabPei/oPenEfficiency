using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnPropertyExtraction",
        Name = "Inspect Fonts & Colors",
        SearchAliases = new[] { "Property Extraction" },
        Tooltip = "Inspect fonts, sizes, and colors",
        IconData = "M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M19,19H5V5H19V19M14,17H7V15H14V17M17,13H7V11H17V13M17,9H7V7H17V9Z",
        Color = "#6B7280",
        Description = "Displays unique fonts, sizes, text colors, fill colors, and line colors in the selection or current slide.",
        DetailedHelpText = "### Inspect Fonts & Colors\nInspects selected shapes or the current slide. Right-click chooses the scope and can copy the full report or a deduplicated hexadecimal color list directly. Reports include grouped shapes, mixed-format text runs, and table-cell text, plus unique fonts, sizes, text colors, fills, and lines. Picture fills are skipped because they have no single fill color.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class PropertyExtractionFeature
    {
        public enum ExtractionScope { Automatic, Selection, CurrentSlide }
        public enum ExtractionOutput { Dialog, Clipboard, ColorListClipboard }

        public struct TextProps
        {
            public List<string> Fonts;
            public List<float> Sizes;
            public List<int> Colors;
        }

        public struct ShapeProps
        {
            public List<int> FillColors;
            public List<int> LineColors;
        }

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, ExtractionScope.Automatic, ExtractionOutput.Dialog);
        }

        public static bool Execute(
            PowerPointManager manager,
            ExtractionScope scope,
            ExtractionOutput output)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app?.ActiveWindow == null || app.ActiveWindow.View.Slide == null) return false;
                bool hasSelection = manager.GetSelectedShapes()?.Count > 0;
                if (scope == ExtractionScope.Selection && !hasSelection)
                {
                    MessageBox.Show("Select one or more shapes to inspect.", "Inspect Fonts & Colors",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return false;
                }
                bool useSelection = scope == ExtractionScope.Selection ||
                                    (scope == ExtractionScope.Automatic && hasSelection);
                var textProps = GetTextProps(manager, !useSelection);
                var shapeProps = GetShapeProps(manager, !useSelection);
                string report = BuildReport(useSelection ? "Selected shapes" : "Current slide", textProps, shapeProps);
                if (output == ExtractionOutput.ColorListClipboard)
                {
                    var colors = textProps.Colors.Concat(shapeProps.FillColors).Concat(shapeProps.LineColors)
                        .Distinct().OrderBy(value => value).Select(ToHex).ToArray();
                    if (colors.Length == 0)
                    {
                        MessageBox.Show("No text, fill, or line colors were found in this scope.",
                            "Inspect Fonts & Colors", MessageBoxButton.OK, MessageBoxImage.Information);
                        return true;
                    }
                    Clipboard.SetText(string.Join(Environment.NewLine, colors));
                    return true;
                }
                if (output == ExtractionOutput.Clipboard)
                {
                    Clipboard.SetText(report);
                    return true;
                }
                var result = MessageBox.Show(report + "\n\nCopy this report to the clipboard?", "Inspect Fonts & Colors",
                    MessageBoxButton.YesNo, MessageBoxImage.Information);
                if (result == MessageBoxResult.Yes)
                {
                    try { Clipboard.SetText(report); }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "PropertyExtractionFeature.CopyReport");
                        MessageBox.Show("The report was created, but could not be copied to the clipboard.", "Clipboard Unavailable",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PropertyExtractionFeature.Execute");
                return false;
            }
        }

        public static TextProps GetTextPropsInSelection(PowerPointManager manager)
        {
            return GetTextProps(manager, false);
        }

        private static TextProps GetTextProps(PowerPointManager manager, bool forceCurrentSlide)
        {
            var props = new TextProps { Fonts = new List<string>(), Sizes = new List<float>(), Colors = new List<int>() };
            foreach (var shape in GetTargetShapes(manager, forceCurrentSlide))
            {
                try
                {
                    if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        for (int row = 1; row <= shape.Table.Rows.Count; row++)
                            for (int column = 1; column <= shape.Table.Columns.Count; column++)
                            {
                                var cellShape = shape.Table.Cell(row, column).Shape;
                                if (cellShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                    SetTextProps(cellShape.TextFrame.TextRange, props);
                            }
                    }
                    else if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        SetTextProps(shape.TextFrame.TextRange, props);
                    }
                }
                catch (Exception ex) { ExceptionLogger.Log(ex, "PropertyExtractionFeature.TextProperties"); }
            }
            props.Fonts = oPenEfficiency.Core.Services.PropertyReportPolicy.NormalizeFonts(props.Fonts);
            props.Sizes = oPenEfficiency.Core.Services.PropertyReportPolicy.NormalizeSizes(props.Sizes);
            props.Colors = oPenEfficiency.Core.Services.PropertyReportPolicy.NormalizeColors(props.Colors);
            return props;
        }

        public static ShapeProps GetShapePropsInSelection(PowerPointManager manager)
        {
            return GetShapeProps(manager, false);
        }

        private static ShapeProps GetShapeProps(PowerPointManager manager, bool forceCurrentSlide)
        {
            var props = new ShapeProps { FillColors = new List<int>(), LineColors = new List<int>() };
            foreach (var shape in GetTargetShapes(manager, forceCurrentSlide))
            {
                try
                {
                    if (shape.Fill.Visible == Office.MsoTriState.msoTrue && shape.Fill.Type != Office.MsoFillType.msoFillPicture)
                        props.FillColors.Add(shape.Fill.ForeColor.RGB);
                }
                catch (Exception ex) { ExceptionLogger.Log(ex, "PropertyExtractionFeature.FillColor"); }
                try
                {
                    if (shape.Line.Visible == Office.MsoTriState.msoTrue) props.LineColors.Add(shape.Line.ForeColor.RGB);
                }
                catch (Exception ex) { ExceptionLogger.Log(ex, "PropertyExtractionFeature.LineColor"); }
            }
            props.FillColors = oPenEfficiency.Core.Services.PropertyReportPolicy.NormalizeColors(props.FillColors);
            props.LineColors = oPenEfficiency.Core.Services.PropertyReportPolicy.NormalizeColors(props.LineColors);
            return props;
        }

        private static List<Shape> GetTargetShapes(PowerPointManager manager, bool forceCurrentSlide)
        {
            var result = new List<Shape>();
            if (manager == null) return result;
            var selected = manager.GetSelectedShapes();
            if (!forceCurrentSlide && selected != null && selected.Count > 0)
            {
                foreach (Shape shape in selected) AddRecursive(shape, result);
                return result;
            }

            var app = manager.GetApplication();
            if (app?.ActiveWindow == null) return result;
            var slide = app.ActiveWindow.View.Slide as Slide;
            if (slide == null) return result;
            foreach (Shape shape in slide.Shapes) AddRecursive(shape, result);
            return result;
        }

        private static void AddRecursive(Shape shape, List<Shape> target)
        {
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (Shape child in shape.GroupItems) AddRecursive(child, target);
            }
            else target.Add(shape);
        }

        private static void SetTextProps(TextRange range, TextProps props)
        {
            if (range == null) return;
            bool foundRun = false;
            try
            {
                foreach (TextRange run in range.Runs())
                {
                    foundRun = true;
                    AddTextProperties(run, props);
                }
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "PropertyExtractionFeature.TextRuns"); }
            if (!foundRun) AddTextProperties(range, props);
        }

        private static void AddTextProperties(TextRange range, TextProps props)
        {
            try { if (!string.IsNullOrWhiteSpace(range.Font.Name)) props.Fonts.Add(range.Font.Name); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "PropertyExtractionFeature.FontName"); }
            try { if (range.Font.Size > 0) props.Sizes.Add(range.Font.Size); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "PropertyExtractionFeature.FontSize"); }
            try { if (range.Font.Color.RGB >= 0) props.Colors.Add(range.Font.Color.RGB); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "PropertyExtractionFeature.FontColor"); }
        }

        private static string BuildReport(string scope, TextProps text, ShapeProps shapes)
        {
            return oPenEfficiency.Core.Services.PropertyReportPolicy.BuildReport(
                scope, text.Fonts, text.Sizes, text.Colors, shapes.FillColors, shapes.LineColors);
        }

        private static string ToHex(int colorRef)
        {
            return oPenEfficiency.Core.Services.PropertyReportPolicy.ToHex(colorRef);
        }
    }
}
