using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Convert a PowerPoint table into a set of grouped rectangular shapes, preserving text and basic formatting.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnConvertToShapes",
        Name = "Convert to Shapes",
        Tooltip = "Convert to shapes",
        IconData = "M4,4H20V20H4V4M6,6V10H10V6H6M12,6V10H16V6H12M18,6V10H18V6M6,12V16H10V12H6M12,12V16H16V12H12M18,12V16H18V12M6,18V18H10V18H6M12,18V18H16V18H12M18,18V18H18V18Z",
        Color = "#F43F5E",
        Description = "Converts a PowerPoint table into grouped rectangular shapes, preserving text and basic formatting.",
        DetailedHelpText = "### Convert to Shapes\nDecomposes a native PowerPoint table into individually editable rectangle shapes and text boxes, one per cell, allowing independent animation and styling.",
        Keywords = "explode table, break table, table to grid of shapes, ungroup table cells",
        MinSelection = 0,
        RequiresTable = true)]
    public static class ConvertTableToShapesFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;

            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;

                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                var slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                int targetZ = tableShape.ZOrderPosition;
                string originalName = tableShape.Name;
                var created = new List<string>();
                var seen = new HashSet<string>();

                for (int r = 1; r <= table.Rows.Count; r++)
                {
                    for (int c = 1; c <= table.Columns.Count; c++)
                    {
                        var cell = table.Cell(r, c);
                        var cs = cell.Shape;
                        float left = tableShape.Left + cs.Left;
                        float top = tableShape.Top + cs.Top;
                        float width = cs.Width;
                        float height = cs.Height;
                        string key = $"{Math.Round(left, 2)}|{Math.Round(top, 2)}|{Math.Round(width, 2)}|{Math.Round(height, 2)}";
                        if (seen.Contains(key)) continue;
                        seen.Add(key);

                        var rect = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
                        rect.Name = $"oPE_Cell_{r}_{c}_" + Guid.NewGuid().ToString().Substring(0, 8);

                        try
                        {
                            rect.Fill.Visible = cs.Fill.Visible;
                            rect.Fill.ForeColor.RGB = cs.Fill.ForeColor.RGB;
                            try { rect.Fill.Transparency = cs.Fill.Transparency; } catch (Exception ex) { ExceptionLogger.Log(ex, "ConvertTableToShapesFeature.cellFill"); }
                            rect.Fill.Solid();
                        }
                        catch (Exception ex)
                        {
                            ExceptionLogger.Log(ex, "ConvertTableToShapesFeature.cellFormat");
                        }

                        try
                        {
                            rect.Line.Visible = cs.Line.Visible;
                            rect.Line.ForeColor.RGB = cs.Line.ForeColor.RGB;
                            rect.Line.Weight = cs.Line.Weight;
                            try { rect.Line.DashStyle = cs.Line.DashStyle; } catch (Exception ex) { ExceptionLogger.Log(ex, "ConvertTableToShapesFeature.cellLine"); }
                        }
                        catch (Exception ex)
                        {
                            ExceptionLogger.Log(ex, "ConvertTableToShapesFeature.cellFormat");
                        }

                        try
                        {
                            rect.TextFrame.TextRange.Text = cs.TextFrame.TextRange.Text;
                            var sF = cs.TextFrame.TextRange.Font;
                            var tF = rect.TextFrame.TextRange.Font;
                            tF.Name = sF.Name;
                            tF.Size = sF.Size;
                            tF.Color.RGB = sF.Color.RGB;
                            tF.Bold = sF.Bold;
                            tF.Italic = sF.Italic;
                            try { rect.TextFrame.TextRange.ParagraphFormat.Alignment = cs.TextFrame.TextRange.ParagraphFormat.Alignment; } catch (Exception ex) { ExceptionLogger.Log(ex, "ConvertTableToShapesFeature.cellText"); }
                            try { rect.TextFrame.VerticalAnchor = cs.TextFrame.VerticalAnchor; } catch (Exception ex) { ExceptionLogger.Log(ex, "ConvertTableToShapesFeature.cellText"); }
                            try { rect.TextFrame.Orientation = cs.TextFrame.Orientation; } catch (Exception ex) { ExceptionLogger.Log(ex, "ConvertTableToShapesFeature.cellText"); }
                            rect.TextFrame.MarginLeft = cs.TextFrame.MarginLeft;
                            rect.TextFrame.MarginRight = cs.TextFrame.MarginRight;
                            rect.TextFrame.MarginTop = cs.TextFrame.MarginTop;
                            rect.TextFrame.MarginBottom = cs.TextFrame.MarginBottom;
                        }
                        catch (Exception ex)
                        {
                            ExceptionLogger.Log(ex, "ConvertTableToShapesFeature.cellFormat");
                        }
                        created.Add(rect.Name);
                    }
                }

                if (created.Count == 0) { return false; }
                var group = slide.Shapes.Range(created.ToArray()).Group();
                group.Name = "oPen_ConvertedTable_" + originalName;

                int guard = 0;
                while (group.ZOrderPosition < targetZ && guard < 1000)
                {
                    group.ZOrder(Office.MsoZOrderCmd.msoBringForward);
                    guard++;
                }
                while (group.ZOrderPosition > targetZ && guard < 2000)
                {
                    group.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                    guard++;
                }

                var res = System.Windows.Forms.MessageBox.Show("Delete original table? Click No to hide.", "Convert Table to Shapes", System.Windows.Forms.MessageBoxButtons.YesNoCancel, System.Windows.Forms.MessageBoxIcon.Question);
                if (res == System.Windows.Forms.DialogResult.Yes) { try { tableShape.Delete(); } catch (Exception ex) { ExceptionLogger.Log(ex, "ConvertTableToShapesFeature.deleteOriginal"); } }
                else if (res == System.Windows.Forms.DialogResult.No) { try { tableShape.Visible = Office.MsoTriState.msoFalse; } catch (Exception ex) { ExceptionLogger.Log(ex, "ConvertTableToShapesFeature.hideOriginal"); } }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ConvertTableToShapesFeature.Execute");
                return false;
            }
        }
    }
}
