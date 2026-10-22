using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Tables
{
    /// <summary>
    /// Options for a programmatic heatmap run. The dialog builds one from its controls; the
    /// seam-driven automation entry builds one directly. This is the complete input surface of
    /// <see cref="TableHeatmapApplier.Apply"/> - the applier never reads UI state.
    /// </summary>
    public sealed class TableHeatmapOptions
    {
        public UI.Dialogs.TableHeatmapScope Scope { get; set; } = UI.Dialogs.TableHeatmapScope.Row;
        public bool ApplyToFill { get; set; }
        public bool ApplyIcons { get; set; }
        public bool ExcludeHeader { get; set; }
        public bool IgnoreEmpty { get; set; }
        public int ColorPresetIndex { get; set; }
        public int IconPresetIndex { get; set; }
        // Custom color defaults mirror the dialog's picker defaults.
        public System.Drawing.Color CustomLow { get; set; } = System.Drawing.Color.FromArgb(248, 113, 113);
        public System.Drawing.Color? CustomMid { get; set; } = System.Drawing.Color.FromArgb(251, 191, 36);
        public System.Drawing.Color CustomHigh { get; set; } = System.Drawing.Color.FromArgb(52, 211, 153);
    }

    /// <summary>
    /// Outcome of a heatmap run. The dialog maps these flags back to its original message boxes;
    /// the programmatic entry maps them to a boolean. No message box is ever raised here.
    /// </summary>
    public sealed class TableHeatmapApplyResult
    {
        public bool NoDataRows { get; set; }
        public bool NoNumericValues { get; set; }
        public int Formatted { get; set; }
        public int Failed { get; set; }
    }

    /// <summary>
    /// The heatmap application core. This code previously lived inside
    /// <c>UI.Dialogs.TableHeatmapDialog.ApplyHeatmap</c> and its helpers and was moved verbatim so
    /// the exact same math, parsing, icon and formatting behavior runs for the dialog and for
    /// dialog-free automation. The dialog keeps ownership of every user-facing message.
    /// </summary>
    public static class TableHeatmapApplier
    {
        private sealed class HeatmapGroup
        {
            public List<Cell> Cells { get; } = new List<Cell>();
            public double Minimum { get; set; }
            public double Maximum { get; set; }
        }

        public static TableHeatmapApplyResult Apply(PowerPointManager manager, Table table, TableHeatmapOptions options)
        {
            var result = new TableHeatmapApplyResult();
            if (manager == null || table == null || options == null) return result;

            bool isTableScope = options.Scope == UI.Dialogs.TableHeatmapScope.EntireTable;
            bool isRowScope = options.Scope == UI.Dialogs.TableHeatmapScope.Row;
            bool isColScope = options.Scope == UI.Dialogs.TableHeatmapScope.Column;
            bool excludeHeader = options.ExcludeHeader;
            bool ignoreEmpty = options.IgnoreEmpty;
            int presetIndex = options.ColorPresetIndex;


            int rows = table.Rows.Count;
            int cols = table.Columns.Count;
            int startRow = excludeHeader ? 2 : 1;
            if (startRow > rows)
            {
                result.NoDataRows = true;
                return result;
            }

            var rawGroups = new List<List<Cell>>();
            if (isTableScope)
            {
                var cells = new List<Cell>();
                for (int row = startRow; row <= rows; row++)
                    for (int column = 1; column <= cols; column++) cells.Add(table.Cell(row, column));
                rawGroups.Add(cells);
            }
            else if (isRowScope)
            {
                for (int row = startRow; row <= rows; row++)
                {
                    var cells = new List<Cell>();
                    for (int column = 1; column <= cols; column++) cells.Add(table.Cell(row, column));
                    rawGroups.Add(cells);
                }
            }
            else if (isColScope)
            {
                for (int column = 1; column <= cols; column++)
                {
                    var cells = new List<Cell>();
                    for (int row = startRow; row <= rows; row++) cells.Add(table.Cell(row, column));
                    rawGroups.Add(cells);
                }
            }

            var groups = new List<HeatmapGroup>();
            foreach (var cells in rawGroups)
            {
                double minimum = double.MaxValue;
                double maximum = double.MinValue;
                bool found = false;
                foreach (var cell in cells)
                {
                    if (!TryParseCell(cell, out double value)) continue;
                    minimum = Math.Min(minimum, value);
                    maximum = Math.Max(maximum, value);
                    found = true;
                }
                if (!found) continue;
                var group = new HeatmapGroup { Minimum = minimum, Maximum = maximum };
                group.Cells.AddRange(cells);
                groups.Add(group);
            }

            if (groups.Count == 0)
            {
                result.NoNumericValues = true;
                return result;
            }

            try { manager.GetApplication().StartNewUndoEntry(); } catch { /* Heatmap formatting remains valid when undo grouping is unavailable. */ }
            int formatted = 0;
            int failed = 0;
            foreach (var group in groups)
            {
                foreach (var cell in group.Cells)
                {
                    bool numeric = TryParseCell(cell, out double value);
                    bool blank = string.IsNullOrWhiteSpace(cell.Shape.TextFrame.TextRange.Text);
                    if (!numeric && (!blank || ignoreEmpty)) continue;
                    if (!numeric) value = group.Minimum;
                    if (ApplyFormattingToCell(cell, value, group.Minimum, group.Maximum, options, presetIndex)) formatted++;
                    else failed++;
                }
            }

            result.Formatted = formatted;
            result.Failed = failed;
            return result;
        }

        private static bool TryParseCell(Cell cell, out double value)
        {
            value = 0;
            try
            {
                var shape = cell.Shape;
                if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    string text = shape.TextFrame.TextRange.Text.Trim();
                    string[] knownIcons = { "▲", "▶", "▼", "↗", "↘", "●", "✔", "❗", "✖", "★", "☆" };
                    foreach (var icon in knownIcons)
                    {
                        if (text.StartsWith(icon, StringComparison.Ordinal))
                        {
                            text = text.Substring(icon.Length).TrimStart();
                            break;
                        }
                    }
                    return oPenEfficiency.Features.TableHelper.TryParseDoubleFromCell(text, out value);
                }
            }
            catch { /* A non-numeric or inaccessible cell is excluded from heatmap interpolation. */ }
            return false;
        }

        private static bool ApplyFormattingToCell(Cell cell, double value, double min, double max, TableHeatmapOptions options, int presetIndex)
        {
            try
            {
                double normalized = 0;
                if (max > min)
                {
                    normalized = (value - min) / (max - min);
                }

                if (options.ApplyIcons)
                {
                    return ApplyIconToCell(cell, normalized, options.IconPresetIndex);
                }

                // REG-TABLEHEATMAP-PREVIEW-001: the dialog preview and applied cell colors share this path.
                System.Drawing.Color finalColor = ComputeColor(normalized, presetIndex, options);

                int bgr = System.Drawing.ColorTranslator.ToOle(finalColor);

                if (options.ApplyToFill)
                {
                    cell.Shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                    cell.Shape.Fill.Solid();
                    cell.Shape.Fill.ForeColor.RGB = bgr;
                }
                else
                {
                    if (cell.Shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && cell.Shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        cell.Shape.TextFrame.TextRange.Font.Color.RGB = bgr;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableHeatmapApplier.ApplyColorToCell");
                return false;
            }
        }

        /// <summary>
        /// Resolves the active color scale from the preset index and custom colors. Shared by the
        /// applier and the dialog preview so the preview can never drift from the applied colors.
        /// </summary>
        public static void ResolveColors(int presetIndex, TableHeatmapOptions options, out System.Drawing.Color low, out System.Drawing.Color? mid, out System.Drawing.Color high)
        {
            System.Drawing.Color green = System.Drawing.Color.FromArgb(99, 190, 123);
            System.Drawing.Color yellow = System.Drawing.Color.FromArgb(255, 235, 132);
            System.Drawing.Color red = System.Drawing.Color.FromArgb(248, 105, 107);
            System.Drawing.Color white = System.Drawing.Color.FromArgb(255, 255, 255);
            System.Drawing.Color blue = System.Drawing.Color.FromArgb(90, 138, 198);
            mid = null;

            switch (presetIndex)
            {
                case 0: // 3-Color: Red to Yellow to Green
                    low = red; mid = yellow; high = green; break;
                case 1: // 3-Color: Green to Yellow to Red
                    low = green; mid = yellow; high = red; break;
                case 2: // 3-Color: Red to White to Green
                    low = red; mid = white; high = green; break;
                case 3: // 3-Color: Green to White to Red
                    low = green; mid = white; high = red; break;
                case 4: // 3-Color: Red to White to Blue
                    low = red; mid = white; high = blue; break;
                case 5: // 3-Color: Blue to White to Red
                    low = blue; mid = white; high = red; break;
                case 6: // 2-Color: White to Red
                    low = white; high = red; break;
                case 7: // 2-Color: Red to White
                    low = red; high = white; break;
                case 8: // 2-Color: White to Green
                    low = white; high = green; break;
                case 9: // 2-Color: Green to White
                    low = green; high = white; break;
                case 10: // 2-Color: Yellow to Green
                    low = yellow; high = green; break;
                case 11: // 2-Color: Green to Yellow
                    low = green; high = yellow; break;
                case 12: // Custom
                default:
                    low = options.CustomLow;
                    mid = options.CustomMid;
                    high = options.CustomHigh;
                    break;
            }
        }

        /// <summary>
        /// Color for one normalized value (0..1) under the active preset - the exact math that
        /// applied cells receive (REG-TABLEHEATMAP-PREVIEW-001).
        /// </summary>
        public static System.Drawing.Color ComputeColor(double normalized, int presetIndex, TableHeatmapOptions options)
        {
            ResolveColors(presetIndex, options, out System.Drawing.Color low, out System.Drawing.Color? mid, out System.Drawing.Color high);
            if (mid.HasValue)
            {
                if (normalized < 0.5)
                {
                    double n = normalized * 2; // scale 0-0.5 to 0-1
                    return InterpolateColor(low, mid.Value, n);
                }
                else
                {
                    double n = (normalized - 0.5) * 2; // scale 0.5-1 to 0-1
                    return InterpolateColor(mid.Value, high, n);
                }
            }
            return InterpolateColor(low, high, normalized);
        }

        /// <summary>
        /// Icon glyph and glyph color for one normalized value under the active icon preset - the
        /// exact math that applied cells receive (REG-TABLEHEATMAP-PREVIEW-001).
        /// </summary>
        public static void ComputeIcon(double normalized, int presetIndex, out string icon, out System.Drawing.Color color)
        {
            icon = "";
            color = System.Drawing.Color.Black;

            switch (presetIndex)
            {
                case 0: // 3 Arrows (▲ ▶ ▼)
                    if (normalized >= 0.66) { icon = "▲"; color = System.Drawing.Color.FromArgb(99, 190, 123); } // Green
                    else if (normalized >= 0.33) { icon = "▶"; color = System.Drawing.Color.FromArgb(255, 235, 132); } // Yellow
                    else { icon = "▼"; color = System.Drawing.Color.FromArgb(248, 105, 107); } // Red
                    break;
                case 1: // 4 Arrows (▲ ↗ ↘ ▼)
                    if (normalized >= 0.75) { icon = "▲"; color = System.Drawing.Color.FromArgb(99, 190, 123); }
                    else if (normalized >= 0.50) { icon = "↗"; color = System.Drawing.Color.FromArgb(255, 235, 132); }
                    else if (normalized >= 0.25) { icon = "↘"; color = System.Drawing.Color.FromArgb(248, 105, 107); }
                    else { icon = "▼"; color = System.Drawing.Color.FromArgb(248, 105, 107); }
                    break;
                case 2: // 3 Traffic Lights (● ● ●)
                    if (normalized >= 0.66) { icon = "●"; color = System.Drawing.Color.FromArgb(99, 190, 123); }
                    else if (normalized >= 0.33) { icon = "●"; color = System.Drawing.Color.FromArgb(255, 235, 132); }
                    else { icon = "●"; color = System.Drawing.Color.FromArgb(248, 105, 107); }
                    break;
                case 3: // 3 Indicators (✔ ❗ ✖)
                    if (normalized >= 0.66) { icon = "✔"; color = System.Drawing.Color.FromArgb(99, 190, 123); }
                    else if (normalized >= 0.33) { icon = "❗"; color = System.Drawing.Color.FromArgb(255, 235, 132); }
                    else { icon = "✖"; color = System.Drawing.Color.FromArgb(248, 105, 107); }
                    break;
                case 4: // 3 Ratings (★)
                    if (normalized >= 0.66) { icon = "★"; color = System.Drawing.Color.FromArgb(255, 235, 132); }
                    else if (normalized >= 0.33) { icon = "★"; color = System.Drawing.Color.Silver; }
                    else { icon = "★"; color = System.Drawing.Color.SandyBrown; }
                    break;
            }
        }

        private static System.Drawing.Color InterpolateColor(System.Drawing.Color c1, System.Drawing.Color c2, double fraction)
        {
            fraction = Math.Max(0, Math.Min(1, fraction));
            int r = (int)(c1.R + (c2.R - c1.R) * fraction);
            int g = (int)(c1.G + (c2.G - c1.G) * fraction);
            int b = (int)(c1.B + (c2.B - c1.B) * fraction);
            return System.Drawing.Color.FromArgb(r, g, b);
        }

        private static bool ApplyIconToCell(Cell cell, double normalized, int presetIndex)
        {
            try
            {
                ComputeIcon(normalized, presetIndex, out string icon, out System.Drawing.Color color);

                if (cell.Shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    var textRange = cell.Shape.TextFrame.TextRange;
                    string oldText = textRange.Text;

                    // Clean existing icons
                    string[] knownIcons = { "▲", "▶", "▼", "↗", "↘", "●", "✔", "❗", "✖", "★", "☆" };
                    foreach (var k in knownIcons)
                    {
                        int removeLength = oldText.StartsWith(k + " ") ? k.Length + 1 : (oldText.StartsWith(k) ? k.Length : 0);
                        if (removeLength > 0)
                        {
                            textRange.Characters(1, removeLength).Delete();
                            break;
                        }
                    }

                    textRange.InsertBefore(icon + " ");
                    var firstChar = textRange.Characters(1, 1);
                    firstChar.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(color);
                    firstChar.Font.Name = "Segoe UI Symbol";
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableHeatmapApplier.ApplyIconToCell");
                return false;
            }
        }
    }
}
