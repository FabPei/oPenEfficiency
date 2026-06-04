using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.UI;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Align shapes to their containing table cell (Top, Middle, Bottom x Left, Center, Right).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnAlignToTableCell",
        Name = "Align to Cell",
        Tooltip = "Align to cell",
        IconData = "M3,3H21V21H3V3M5,5V10H10V5H5M11,5V10H19V5H11M5,11V19H10V11H5M11,11V19H19V11H11M13,13L15,15M15,13L13,15",
        Color = "#F43F5E",
        Description = "Opens dialog to align shapes to their containing table cell.",
        DetailedHelpText = "### Align to Cell\nAligns the selected shapes to be perfectly centered inside the underlying table cell directly beneath them. Detects the intersecting table automatically.",
        Keywords = "table cell, center in cell, align inside table, fit to cell",
        MinSelection = 1,
        RequiresTable = true)]
    public static class AlignToTableCellFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - returns false as Align to Cell requires dialog interaction.
        /// Manual switch in MainSidebar.xaml.cs handles opening the AlignToTableWindow dialog.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var win = new UI.AlignToTableWindow(manager);
            win.Show();
            return true;
        }

        public static void Align(PowerPointManager manager, string alignmentType)
        {
            if (manager == null || string.IsNullOrWhiteSpace(alignmentType)) return;

            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return;
                var selection = app.ActiveWindow.Selection;
                if (selection.Type != PpSelectionType.ppSelectionShapes) return;

                var shapes = selection.ShapeRange;
                var slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return;

                // Find all tables on the slide
                var tableShapes = new List<Shape>();
                foreach (Shape s in slide.Shapes)
                {
                    if (s.HasTable == Office.MsoTriState.msoTrue) tableShapes.Add(s);
                }

                if (tableShapes.Count == 0) return;

                foreach (Shape shape in shapes)
                {
                    if (shape.HasTable == Office.MsoTriState.msoTrue) continue;

                    float centerX = shape.Left + (shape.Width / 2);
                    float centerY = shape.Top + (shape.Height / 2);

                    Cell targetCell = null;
                    foreach (var tableShape in tableShapes)
                    {
                        var table = tableShape.Table;
                        for (int r = 1; r <= table.Rows.Count; r++)
                        {
                            for (int c = 1; c <= table.Columns.Count; c++)
                            {
                                var cell = table.Cell(r, c);
                                float cellLeft = cell.Shape.Left;
                                float cellTop = cell.Shape.Top;
                                float cellWidth = cell.Shape.Width;
                                float cellHeight = cell.Shape.Height;

                                if (centerX >= cellLeft && centerX <= cellLeft + cellWidth &&
                                    centerY >= cellTop && centerY <= cellTop + cellHeight)
                                {
                                    targetCell = cell;
                                    break;
                                }
                            }
                            if (targetCell != null) break;
                        }
                        if (targetCell != null) break;
                    }

                    if (targetCell != null)
                    {
                        float cellLeft = targetCell.Shape.Left;
                        float cellTop = targetCell.Shape.Top;
                        float cellWidth = targetCell.Shape.Width;
                        float cellHeight = targetCell.Shape.Height;

                        float targetLeft = shape.Left;
                        float targetTop = shape.Top;

                        // Parse alignmentType (e.g. "TopLeft")
                        string al = alignmentType.ToLower();

                        if (al.Contains("left")) targetLeft = cellLeft;
                        else if (al.Contains("center")) targetLeft = cellLeft + (cellWidth / 2) - (shape.Width / 2);
                        else if (al.Contains("right")) targetLeft = cellLeft + cellWidth - shape.Width;

                        if (al.Contains("top")) targetTop = cellTop;
                        else if (al.Contains("middle")) targetTop = cellTop + (cellHeight / 2) - (shape.Height / 2);
                        else if (al.Contains("bottom")) targetTop = cellTop + cellHeight - shape.Height;

                        shape.Left = targetLeft;
                        shape.Top = targetTop;
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AlignToTableCellFeature.Execute");
            }
        }
    }
}
