using System;
using System.Linq;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Convert selected shapes to a PowerPoint table.
    /// Arranges shapes into a grid based on their positions and creates a table with matching structure.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnConvertToTable",
        Name = "Convert to Table",
        Tooltip = "Convert to table",
        IconData = "M3,3H21V21H3V3M5,5V19H19V5H5M7,7H17V9H7V7M7,11H17V13H7V11M7,15H13V17H7V15Z",
        Color = "#F43F5E",
        Description = "Converts selected shapes into a PowerPoint table, arranging them into a grid based on their positions.",
        DetailedHelpText = "### Convert to Table\nConverts a selection of aligned rectangle shapes into a native PowerPoint table, preserving fill colors, text content, and relative grid positions.",
        MinSelection = 2,
        RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes)]
    public static class ConvertShapesToTableFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;

            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;

                var selection = manager.GetSelectedShapes();
                if (selection == null || selection.Count < 2)
                {
                    System.Windows.Forms.MessageBox.Show("Please select at least 2 shapes to convert to a table.", "oPenEfficiency", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null) return false;

                // Get all selected shapes and sort them by position (top-to-bottom, left-to-right)
                var shapes = new System.Collections.Generic.List<PowerPoint.Shape>();
                for (int i = 1; i <= selection.Count; i++)
                {
                    shapes.Add(selection[i]);
                }

                // Group shapes into rows based on their vertical position
                var rows = new System.Collections.Generic.List<System.Collections.Generic.List<PowerPoint.Shape>>();
                float rowThreshold = 10f; // Shapes within 10 points are considered same row

                // Sort by top position first
                var sortedByTop = shapes.OrderBy(s => s.Top).ToList();

                var currentRow = new System.Collections.Generic.List<PowerPoint.Shape>();
                float? currentRowTop = null;

                foreach (var shape in sortedByTop)
                {
                    if (currentRowTop == null)
                    {
                        currentRowTop = shape.Top;
                        currentRow.Add(shape);
                    }
                    else if (Math.Abs(shape.Top - currentRowTop.Value) < rowThreshold)
                    {
                        currentRow.Add(shape);
                    }
                    else
                    {
                        // Sort current row by left position and add to rows
                        currentRow.Sort((a, b) => a.Left.CompareTo(b.Left));
                        rows.Add(currentRow);
                        currentRow = new System.Collections.Generic.List<PowerPoint.Shape> { shape };
                        currentRowTop = shape.Top;
                    }
                }

                // Add the last row
                if (currentRow.Count > 0)
                {
                    currentRow.Sort((a, b) => a.Left.CompareTo(b.Left));
                    rows.Add(currentRow);
                }

                // Validate that all rows have the same number of columns
                int columnCount = rows[0].Count;
                bool uniformColumns = rows.All(r => r.Count == columnCount);

                if (!uniformColumns)
                {
                    var result = System.Windows.Forms.MessageBox.Show(
                        "Shapes are not aligned in a uniform grid. Continue with maximum columns?",
                        "oPenEfficiency",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

                    if (result == System.Windows.Forms.DialogResult.No)
                        return false;

                    columnCount = rows.Max(r => r.Count);
                }

                // Calculate table bounds
                float minX = shapes.Min(s => s.Left);
                float minY = shapes.Min(s => s.Top);
                float maxX = shapes.Max(s => s.Left + s.Width);
                float maxY = shapes.Max(s => s.Top + s.Height);
                float tableWidth = maxX - minX;
                float tableHeight = maxY - minY;

                // Calculate column widths and row heights
                var colWidths = new float[columnCount];
                var rowHeights = new float[rows.Count];

                for (int col = 0; col < columnCount; col++)
                {
                    foreach (var row in rows)
                    {
                        if (col < row.Count)
                        {
                            colWidths[col] = Math.Max(colWidths[col], row[col].Width);
                        }
                    }
                }

                for (int row = 0; row < rows.Count; row++)
                {
                    foreach (var shape in rows[row])
                    {
                        rowHeights[row] = Math.Max(rowHeights[row], shape.Height);
                    }
                }

                // Create the table
                var table = slide.Shapes.AddTable(rows.Count, columnCount, minX, minY, tableWidth, tableHeight).Table;

                // Set column widths and row heights
                for (int col = 0; col < columnCount; col++)
                {
                    table.Columns[col + 1].Width = colWidths[col];
                }

                for (int row = 0; row < rows.Count; row++)
                {
                    table.Rows[row + 1].Height = rowHeights[row];
                }

                // Copy shape content to table cells
                for (int row = 0; row < rows.Count; row++)
                {
                    for (int col = 0; col < rows[row].Count; col++)
                    {
                        var shape = rows[row][col];
                        var cell = table.Cell(row + 1, col + 1);

                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            cell.Shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text;
                            cell.Shape.TextFrame.TextRange.Font.Size = shape.TextFrame.TextRange.Font.Size;
                        }
                    }
                }

                // Delete original shapes
                foreach (var shape in shapes)
                {
                    shape.Delete();
                }

                // Select the new table - cast to Shape to avoid LINQ confusion
                ((PowerPoint.Shape)table.Parent).Select();

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ConvertShapesToTableFeature.Execute");
                System.Windows.Forms.MessageBox.Show($"Error converting shapes to table: {ex.Message}", "oPenEfficiency Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
    }
}
