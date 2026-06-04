using System;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Splits the currently selected shape into a grid of smaller shapes.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnSplitShape",
        Name = "Split to Grid",
        Tooltip = "Split to grid",
        IconData = "M2,2H10V10H2V2M14,2H22V10H14V2M2,14H10V22H2V14M14,14H22V22H14V14Z",
        Color = "#F43F5E",
        Description = "Splits the selected shape into a grid of smaller shapes (default 2x2 with no spacing).",
        DetailedHelpText = "### Split to Grid\nDivides the selected shape into an N x M grid of equally-sized non-overlapping shapes, with configurable row and column counts.",
        Keywords = "split, grid, divide, chop, partition",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class SplitShapeToGridFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - opens the Split Shape dialog.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var window = new UI.SplitShapeWindow(manager);
            window.Show();
            return true;
        }

        public static System.Collections.Generic.List<PowerPoint.Shape> Execute(PowerPointManager manager, PowerPoint.Shape original, int rows, int columns, double spacingPoints, bool addHorizontalLines, bool addVerticalLines)
        {
            if (manager == null || original == null) return null;
            var createdShapes = new System.Collections.Generic.List<PowerPoint.Shape>();

            try
            {
                var app = manager.GetApplication();
                float totalWidth = original.Width;
                float totalHeight = original.Height;
                float origLeft = original.Left;
                float origTop = original.Top;

                // Calculate dimensions
                float newWidth = (totalWidth - ((float)spacingPoints * (columns - 1))) / (float)columns;
                float newHeight = (totalHeight - ((float)spacingPoints * (rows - 1))) / (float)rows;

                if (newWidth <= 0 || newHeight <= 0) return null;

                // Create grid
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < columns; c++)
                    {
                        var dupRange = original.Duplicate();
                        var dup = dupRange[1];
                        
                        try
                        {
                            if (dup.HasTextFrame == Office.MsoTriState.msoTrue)
                            {
                                dup.TextFrame.DeleteText();
                            }
                        }
                        catch { }

                        // Position
                        dup.Left = origLeft + (c * (newWidth + (float)spacingPoints));
                        dup.Top = origTop + (r * (newHeight + (float)spacingPoints));
                        dup.Width = newWidth;
                        dup.Height = newHeight;
                        
                        dup.Visible = Office.MsoTriState.msoTrue;
                        createdShapes.Add(dup);
                    }
                }

                var slide = app.ActiveWindow.View.Slide as Slide;

                // Vertical lines (between columns)
                if (addVerticalLines && columns > 1 && slide != null)
                {
                    for (int c = 1; c < columns; c++)
                    {
                        float bottomOfPrev = origLeft + ((c - 1) * (newWidth + (float)spacingPoints)) + newWidth;
                        float topOfCurr = origLeft + (c * (newWidth + (float)spacingPoints));
                        float midX = (bottomOfPrev + topOfCurr) / 2f;

                        var line = slide.Shapes.AddLine(midX, origTop, midX, origTop + totalHeight);
                        line.Line.ForeColor.RGB = 0x000000;
                        line.Line.Weight = 1f;
                        createdShapes.Add(line);
                    }
                }

                // Horizontal lines (between rows)
                if (addHorizontalLines && rows > 1 && slide != null)
                {
                    for (int r = 1; r < rows; r++)
                    {
                        float bottomOfPrev = origTop + ((r - 1) * (newHeight + (float)spacingPoints)) + newHeight;
                        float topOfCurr = origTop + (r * (newHeight + (float)spacingPoints));
                        float midY = (bottomOfPrev + topOfCurr) / 2f;
                        
                        var line = slide.Shapes.AddLine(origLeft, midY, origLeft + totalWidth, midY);
                        line.Line.ForeColor.RGB = 0x000000;
                        line.Line.Weight = 1f;
                        createdShapes.Add(line);
                    }
                }

                return createdShapes;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SplitShapeToGridFeature.Execute");
                return null;
            }
        }
    }
}
