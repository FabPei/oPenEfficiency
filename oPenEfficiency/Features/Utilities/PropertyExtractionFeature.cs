using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Extract text and shape properties from the current selection or slide.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnPropertyExtraction",
        Name = "Property Extraction",
        Tooltip = "Extract properties",
        IconData = "M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M19,19H5V5H19V19M14,17H7V15H14V17M17,13H7V11H17V13M17,9H7V7H17V9Z",
        Color = "#6B7280",
        Description = "Extracts and displays unique fonts, sizes, and colors used in the selected shapes or slide.",
        DetailedHelpText = "### Property Extraction\nExtracts all formatting properties from the selected shape and displays them in a readable panel, useful for debugging or documenting style libraries.",
        Keywords = "extract, fonts, colors, formats, style guide, debug",
        MinSelection = 0)]
    public static class PropertyExtractionFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - extracts properties from current selection.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            var textProps = GetTextPropsInSelection(manager);
            var shapeProps = GetShapePropsInSelection(manager);
            // Properties are extracted - further display handling is done via manual switch
            return textProps.Fonts.Count > 0 || shapeProps.FillColors.Count > 0;
        }

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

        /// <summary>
        /// Gets unique text properties (fonts, sizes, colors) in the current selection or slide.
        /// </summary>
        public static TextProps GetTextPropsInSelection(PowerPointManager manager)
        {
            var props = new TextProps { Fonts = new List<string>(), Sizes = new List<float>(), Colors = new List<int>() };
            if (manager == null) return props;

            var shapes = manager.GetSelectedShapes();
            List<Shape> targetShapes = new List<Shape>();

            if (shapes != null && shapes.Count > 0)
            {
                foreach (Shape s in shapes) GetAllShapesRecursive(s, targetShapes);
            }
            else
            {
                var app = manager.GetApplication();
                var slide = app.ActiveWindow.View.Slide as Slide;
                if (app.ActiveWindow != null && slide != null)
                {
                    foreach (Shape sh in slide.Shapes)
                        GetAllShapesRecursive(sh, targetShapes);
                }
            }

            foreach (var shape in targetShapes)
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    SetTextProps(shape.TextFrame.TextRange, props);
                }
            }

            props.Fonts = props.Fonts.Distinct().OrderBy(f => f).ToList();
            props.Sizes = props.Sizes.Distinct().OrderBy(s => s).ToList();
            props.Colors = props.Colors.Distinct().ToList();
            return props;
        }

        /// <summary>
        /// Gets unique shape properties (fill and line colors) in the current selection or slide.
        /// </summary>
        public static ShapeProps GetShapePropsInSelection(PowerPointManager manager)
        {
            var props = new ShapeProps { FillColors = new List<int>(), LineColors = new List<int>() };
            if (manager == null) return props;

            var shapes = manager.GetSelectedShapes();
            List<Shape> targetShapes = new List<Shape>();

            if (shapes != null && shapes.Count > 0)
            {
                foreach (Shape s in shapes) GetAllShapesRecursive(s, targetShapes);
            }
            else
            {
                var app = manager.GetApplication();
                var slide = app.ActiveWindow.View.Slide as Slide;
                if (app.ActiveWindow != null && slide != null)
                {
                    foreach (Shape sh in slide.Shapes)
                        GetAllShapesRecursive(sh, targetShapes);
                }
            }

            foreach (var shape in targetShapes)
            {
                try
                {
                    if (shape.Fill.Visible == Office.MsoTriState.msoTrue && shape.Fill.Type != Office.MsoFillType.msoFillPicture)
                    {
                        int c = shape.Fill.ForeColor.RGB;
                        if (!props.FillColors.Contains(c)) props.FillColors.Add(c);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"PropertyExtractionFeature.FillColor error: {ex.Message}");
                }

                try
                {
                    if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                        int c = shape.Line.ForeColor.RGB;
                        if (!props.LineColors.Contains(c)) props.LineColors.Add(c);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"PropertyExtractionFeature.LineColor error: {ex.Message}");
                }
            }

            return props;
        }

        private static void GetAllShapesRecursive(Shape shape, List<Shape> allShapes)
        {
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (Shape subShape in shape.GroupItems)
                {
                    GetAllShapesRecursive(subShape, allShapes);
                }
            }
            else
            {
                allShapes.Add(shape);
            }
        }

        private static void SetTextProps(TextRange range, TextProps props)
        {
            if (range == null) return;
            try
            {
                props.Fonts.Add(range.Font.Name);
                props.Sizes.Add(range.Font.Size);
                props.Colors.Add(range.Font.Color.RGB);
            }
            catch
            {
                // Multi-style range, iterate characters or runs
                foreach (TextRange run in range.Runs())
                {
                    try
                    {
                        props.Fonts.Add(run.Font.Name);
                        props.Sizes.Add(run.Font.Size);
                        props.Colors.Add(run.Font.Color.RGB);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"PropertyExtractionFeature.RunProperties error: {ex.Message}");
                    }
                }
            }
        }
    }
}
