using System;
using System.Collections.Generic;
using System.Windows.Interop;
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
        Description = "Aligns selected non-table shapes to one of nine positions inside the table cell containing each shape's center.",
        DetailedHelpText = "### Align to Cell\n\nPlace one or more shapes over table cells, then choose Top/Center/Bottom combined with Left/Center/Right. The right-click menu exposes all nine positions directly in Top, Middle, and Bottom groups. Each shape is matched to the visible table cell containing its center and moved independently; table shapes themselves are ignored. The center action aligns both axes. Shapes whose centers are outside every table cell remain unchanged, and a successful action creates a PowerPoint Undo entry.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class AlignToTableCellFeature
    {
        private static UI.AlignToTableWindow _window;

        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                if (_window != null && _window.IsLoaded)
                {
                    _window.Activate();
                    return true;
                }
                _window = new UI.AlignToTableWindow(manager);
                var app = manager.GetApplication();
                if (app != null) new WindowInteropHelper(_window).Owner = (IntPtr)app.HWND;
                _window.Closed += (s, e) => _window = null;
                _window.Show();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AlignToTableCellFeature.Execute");
                var failedWindow = _window;
                _window = null;
                try { failedWindow?.Close(); }
                catch (Exception closeEx)
                {
                    ExceptionLogger.Log(
                        closeEx,
                        "AlignToTableCellFeature.Execute.FailureCleanup");
                }
                throw;
            }
        }

        public static bool Align(PowerPointManager manager, string alignmentType)
        {
            if (manager == null || string.IsNullOrWhiteSpace(alignmentType)) return false;

            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;
                var selection = app.ActiveWindow.Selection;
                if (selection.Type != PpSelectionType.ppSelectionShapes) return false;

                var shapes = selection.ShapeRange;
                var slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                // Find all tables on the slide
                var tableShapes = new List<Shape>();
                foreach (Shape s in slide.Shapes)
                {
                    if (s.HasTable == Office.MsoTriState.msoTrue && s.Visible != Office.MsoTriState.msoFalse) tableShapes.Add(s);
                }

                if (tableShapes.Count == 0) return false;

                try { app.StartNewUndoEntry(); } catch { /* Alignment remains valid when undo grouping is unavailable. */ }
                int alignedCount = 0;

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
                        else if (al.Contains("middle") || al == "center") targetTop = cellTop + (cellHeight / 2) - (shape.Height / 2);
                        else if (al.Contains("bottom")) targetTop = cellTop + cellHeight - shape.Height;

                        shape.Left = targetLeft;
                        shape.Top = targetTop;
                        alignedCount++;
                    }
                }
                return alignedCount > 0;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AlignToTableCellFeature.Align");
                return false;
            }
        }
    }
}
