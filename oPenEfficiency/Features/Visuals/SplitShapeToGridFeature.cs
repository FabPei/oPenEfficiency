using System;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Splits the currently selected shape into a grid of smaller shapes.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnSplitShape",
        Name = "Split Shape into Grid",
        SearchAliases = new[] { "Split to Grid" },
        Tooltip = "Split shape into grid",
        IconData = "M2,2H10V10H2V2M14,2H22V10H14V2M2,14H10V22H2V14M14,14H22V22H14V14Z",
        Color = "#F43F5E",
        Description = "Splits the selected shape into a grid of smaller shapes (default 2x2 with no spacing).",
        DetailedHelpText = "### Split Shape into Grid\nDivides one selected shape into an N × M grid of equally-sized non-overlapping shapes. Right-click for direct 1×2, 2×1, 2×2, 3×2, or 3×3 presets with zero spacing. Direct presets replace the source only after every grid shape has been created successfully.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
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

                var plannedCells = ShapeGeometryPlanner.SplitGrid(
                    new ShapeGeometry
                    {
                        Id = original.Id,
                        Left = origLeft,
                        Top = origTop,
                        Width = totalWidth,
                        Height = totalHeight
                    },
                    rows,
                    columns,
                    (float)spacingPoints);

                // Create grid
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < columns; c++)
                    {
                        var plannedCell = plannedCells[r * columns + c];
                        var dupRange = original.Duplicate();
                        var dup = dupRange[1];
                        
                        if (dup.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            dup.TextFrame.DeleteText();
                        }

                        // Position
                        dup.Left = plannedCell.Left;
                        dup.Top = plannedCell.Top;
                        dup.Width = plannedCell.Width;
                        dup.Height = plannedCell.Height;
                        
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
                        float bottomOfPrev = plannedCells[c - 1].Left + plannedCells[c - 1].Width;
                        float topOfCurr = plannedCells[c].Left;
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
                        float bottomOfPrev = plannedCells[(r - 1) * columns].Top + plannedCells[(r - 1) * columns].Height;
                        float topOfCurr = plannedCells[r * columns].Top;
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

        public static bool ExecutePreset(PowerPointManager manager, int rows, int columns)
        {
            var selection = manager?.GetSelectedShapes();
            if (selection == null || selection.Count != 1) return false;
            var original = selection[1];
            System.Collections.Generic.List<PowerPoint.Shape> created = null;
            try
            {
                try { manager.GetApplication()?.StartNewUndoEntry(); } catch { /* Splitting remains valid when undo grouping is unavailable. */ }
                created = Execute(manager, original, rows, columns, 0, false, false);
                if (created == null || created.Count != rows * columns) throw new InvalidOperationException("The complete grid could not be created.");
                original.Delete();
                try
                {
                    var names = new string[created.Count];
                    for (int index = 0; index < created.Count; index++) names[index] = created[index].Name;
                    ((Slide)created[0].Parent).Shapes.Range(names).Select();
                }
                catch { /* Result selection is optional after a successful split. */ }
                return true;
            }
            catch (Exception ex)
            {
                if (created != null)
                    foreach (var shape in created) { try { shape.Delete(); } catch { /* Rollback shape may already have been removed. */ } }
                ExceptionLogger.Log(ex, "SplitShapeToGridFeature.ExecutePreset");
                return false;
            }
        }
    }
}
