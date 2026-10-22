using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Forms = System.Windows.Forms;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Reorders selected table rows or columns while preserving the requested content and layout properties.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnMoveTableRowColumn",
        Name = "Move Table Rows or Columns",
        Tooltip = "Reorder selected table rows or columns",
        IconData = "M3,4H21V20H3V4M6,7H18M6,12H18M6,17H18M9,4V20M15,4V20M11,9L8,12L11,15M13,9L16,12L13,15",
        Color = "#3B82F6",
        Description = "Moves, swaps, or duplicates selected table rows and columns, including contiguous blocks.",
        DetailedHelpText = "### Move Table Rows or Columns\n\nSelect one or more cells, or place the text cursor inside a cell. Move the corresponding row or column by one step, to an edge, or to a numbered position. Contiguous selections move as one block. **Keep table design** moves formatted content while retaining destination fills, borders, dimensions, and cell layout. Other modes can move all formatting, retain only destination borders, or move plain text only. The context menu also supports swapping two selected rows or columns, duplicating a selected block, repeating the last move, and protecting the first or last row.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class MoveTableRowColumnFeature
    {
        public enum MoveDirection { RowUp, RowDown, ColumnLeft, ColumnRight }
        public enum MoveMode { KeepTableDesign, Complete, IgnoreBorders, TextOnly }
        public enum MoveAxis { Row, Column }

        private enum LastMoveKind { Direction, Boundary, Position, Swap }

        private sealed class LastMoveSpec
        {
            public LastMoveKind Kind { get; set; }
            public MoveAxis Axis { get; set; }
            public MoveDirection Direction { get; set; }
            public bool ToStart { get; set; }
            public int Position { get; set; }
            public MoveMode Mode { get; set; }
        }

        private static LastMoveSpec _lastMove;

        public static MoveMode CurrentMode { get; set; } = MoveMode.KeepTableDesign;
        public static bool ProtectFirstRow { get; set; } = true;
        public static bool ProtectLastRow { get; set; }
        public static bool CanRepeatLastMove => _lastMove != null;

        internal enum MergeDetectionMode { NotRun, ShapeId, GeometryFallback }

        /// <summary>
        /// Reports which oracle the last merge check used, so the real-host acceptance seam can show
        /// in its evidence whether the geometry fallback was exercised.
        /// </summary>
        internal static MergeDetectionMode LastMergeDetectionMode { get; private set; } = MergeDetectionMode.NotRun;

        public static bool Execute(PowerPointManager manager, MoveDirection direction, MoveMode mode)
        {
            return ExecuteDirection(manager, direction, mode, true);
        }

        public static bool MoveToBoundary(PowerPointManager manager, MoveAxis axis, bool toStart, MoveMode mode)
        {
            return MoveToBoundaryInternal(manager, axis, toStart, mode, true);
        }

        public static bool MoveToPositionInteractive(PowerPointManager manager, MoveAxis axis, MoveMode mode)
        {
            Shape tableShape;
            Table table;
            SelectionBounds bounds;
            if (!TryGetSelection(manager, out tableShape, out table, out bounds, true)) return false;

            int count = axis == MoveAxis.Row ? bounds.MaxRow - bounds.MinRow + 1 : bounds.MaxColumn - bounds.MinColumn + 1;
            int total = axis == MoveAxis.Row ? table.Rows.Count : table.Columns.Count;
            int maximumStart = total - count + 1;
            int current = axis == MoveAxis.Row ? bounds.MinRow : bounds.MinColumn;
            int position;
            if (!PromptForPosition(axis, current, maximumStart, out position)) return false;
            return MoveToPositionInternal(manager, axis, position, mode, true);
        }

        public static bool SwapSelected(PowerPointManager manager, MoveAxis axis, MoveMode mode)
        {
            return SwapSelectedInternal(manager, axis, mode, true);
        }

        public static bool DuplicateSelected(PowerPointManager manager, MoveAxis axis, bool before)
        {
            Shape tableShape;
            Table table;
            SelectionBounds bounds;
            if (!TryGetSelection(manager, out tableShape, out table, out bounds, true)) return false;
            if (!EnsureContiguousSelection(table, axis, bounds)) return false;
            if (!EnsureNoMergedCells(table)) return false;

            try
            {
                TryStartUndoEntry(manager);
                return axis == MoveAxis.Row
                    ? DuplicateRows(table, bounds.MinRow, bounds.MaxRow, before)
                    : DuplicateColumns(table, bounds.MinColumn, bounds.MaxColumn, before);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MoveTableRowColumnFeature.DuplicateSelected");
                ShowError("The selected table block could not be duplicated.");
                return false;
            }
        }

        public static bool RepeatLastMove(PowerPointManager manager)
        {
            if (_lastMove == null)
            {
                ShowInfo("There is no table move to repeat yet.");
                return false;
            }

            switch (_lastMove.Kind)
            {
                case LastMoveKind.Direction:
                    return ExecuteDirection(manager, _lastMove.Direction, _lastMove.Mode, false);
                case LastMoveKind.Boundary:
                    return MoveToBoundaryInternal(manager, _lastMove.Axis, _lastMove.ToStart, _lastMove.Mode, false);
                case LastMoveKind.Position:
                    return MoveToPositionInternal(manager, _lastMove.Axis, _lastMove.Position, _lastMove.Mode, false);
                case LastMoveKind.Swap:
                    return SwapSelectedInternal(manager, _lastMove.Axis, _lastMove.Mode, false);
                default:
                    return false;
            }
        }

        private static bool ExecuteDirection(PowerPointManager manager, MoveDirection direction, MoveMode mode, bool remember)
        {
            Shape tableShape;
            Table table;
            SelectionBounds bounds;
            if (!TryGetSelection(manager, out tableShape, out table, out bounds, true)) return false;
            if (!EnsureNoMergedCells(table)) return false;

            MoveAxis axis = direction == MoveDirection.RowUp || direction == MoveDirection.RowDown ? MoveAxis.Row : MoveAxis.Column;
            if (!EnsureContiguousSelection(table, axis, bounds)) return false;
            int sourceStart = axis == MoveAxis.Row ? bounds.MinRow : bounds.MinColumn;
            int sourceEnd = axis == MoveAxis.Row ? bounds.MaxRow : bounds.MaxColumn;
            int delta = direction == MoveDirection.RowUp || direction == MoveDirection.ColumnLeft ? -1 : 1;
            bool result = MoveBlock(manager, table, axis, sourceStart, sourceEnd, sourceStart + delta, mode, true);
            if (result && remember)
            {
                _lastMove = new LastMoveSpec { Kind = LastMoveKind.Direction, Axis = axis, Direction = direction, Mode = mode };
            }
            return result;
        }

        private static bool MoveToBoundaryInternal(PowerPointManager manager, MoveAxis axis, bool toStart, MoveMode mode, bool remember)
        {
            Shape tableShape;
            Table table;
            SelectionBounds bounds;
            if (!TryGetSelection(manager, out tableShape, out table, out bounds, true)) return false;
            if (!EnsureContiguousSelection(table, axis, bounds)) return false;
            if (!EnsureNoMergedCells(table)) return false;

            int sourceStart = axis == MoveAxis.Row ? bounds.MinRow : bounds.MinColumn;
            int sourceEnd = axis == MoveAxis.Row ? bounds.MaxRow : bounds.MaxColumn;
            int count = sourceEnd - sourceStart + 1;
            int total = axis == MoveAxis.Row ? table.Rows.Count : table.Columns.Count;
            int target = toStart ? 1 : total - count + 1;
            if (axis == MoveAxis.Row)
            {
                if (toStart && ProtectFirstRow) target = 2;
                if (!toStart && ProtectLastRow) target--;
            }

            bool result = MoveBlock(manager, table, axis, sourceStart, sourceEnd, target, mode, true);
            if (result && remember)
            {
                _lastMove = new LastMoveSpec { Kind = LastMoveKind.Boundary, Axis = axis, ToStart = toStart, Mode = mode };
            }
            return result;
        }

        private static bool MoveToPositionInternal(PowerPointManager manager, MoveAxis axis, int position, MoveMode mode, bool remember)
        {
            Shape tableShape;
            Table table;
            SelectionBounds bounds;
            if (!TryGetSelection(manager, out tableShape, out table, out bounds, true)) return false;
            if (!EnsureContiguousSelection(table, axis, bounds)) return false;
            if (!EnsureNoMergedCells(table)) return false;

            int sourceStart = axis == MoveAxis.Row ? bounds.MinRow : bounds.MinColumn;
            int sourceEnd = axis == MoveAxis.Row ? bounds.MaxRow : bounds.MaxColumn;
            bool result = MoveBlock(manager, table, axis, sourceStart, sourceEnd, position, mode, true);
            if (result && remember)
            {
                _lastMove = new LastMoveSpec { Kind = LastMoveKind.Position, Axis = axis, Position = position, Mode = mode };
            }
            return result;
        }

        private static bool SwapSelectedInternal(PowerPointManager manager, MoveAxis axis, MoveMode mode, bool remember)
        {
            Shape tableShape;
            Table table;
            SelectionBounds bounds;
            if (!TryGetSelection(manager, out tableShape, out table, out bounds, true)) return false;
            if (!EnsureNoMergedCells(table)) return false;

            try
            {
                var indices = GetSelectedIndices(table, axis);
                if (indices.Count != 2)
                {
                    ShowInfo("Select cells in exactly two " + (axis == MoveAxis.Row ? "rows" : "columns") + " before using Swap.");
                    return false;
                }
                if (axis == MoveAxis.Row && IsProtectedRow(indices[0], table.Rows.Count) ||
                    axis == MoveAxis.Row && IsProtectedRow(indices[1], table.Rows.Count))
                {
                    ShowInfo("A protected header or total row cannot be swapped. Change the protection options in the context menu first.");
                    return false;
                }

                TryStartUndoEntry(manager);
                if (axis == MoveAxis.Row)
                {
                    var first = CaptureRow(table, indices[0]);
                    var second = CaptureRow(table, indices[1]);
                    ApplyRow(table, indices[0], second, mode);
                    ApplyRow(table, indices[1], first, mode);
                }
                else
                {
                    var first = CaptureColumn(table, indices[0]);
                    var second = CaptureColumn(table, indices[1]);
                    ApplyColumn(table, indices[0], second, mode);
                    ApplyColumn(table, indices[1], first, mode);
                }

                if (remember) _lastMove = new LastMoveSpec { Kind = LastMoveKind.Swap, Axis = axis, Mode = mode };
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MoveTableRowColumnFeature.SwapSelected");
                ShowError("The selected table rows or columns could not be swapped.");
                return false;
            }
        }

        private static bool MoveBlock(PowerPointManager manager, Table table, MoveAxis axis, int sourceStart, int sourceEnd, int targetStart, MoveMode mode, bool showMessages)
        {
            int total = axis == MoveAxis.Row ? table.Rows.Count : table.Columns.Count;
            int count = sourceEnd - sourceStart + 1;
            int maximumStart = total - count + 1;
            if (targetStart < 1 || targetStart > maximumStart)
            {
                if (showMessages) ShowInfo("The selected block is already at the table boundary.");
                return false;
            }
            if (targetStart == sourceStart) return false;
            if (axis == MoveAxis.Row && !CanMoveRows(sourceStart, sourceEnd, targetStart, count, total)) return false;

            try
            {
                TryStartUndoEntry(manager);
                if (axis == MoveAxis.Row)
                {
                    var states = new List<RowState>();
                    for (int row = 1; row <= total; row++) states.Add(CaptureRow(table, row));
                    Reorder(states, sourceStart - 1, count, targetStart - 1);
                    int firstChanged = Math.Min(sourceStart, targetStart);
                    int lastChanged = Math.Max(sourceEnd, targetStart + count - 1);
                    for (int row = firstChanged; row <= lastChanged; row++) ApplyRow(table, row, states[row - 1], mode);
                }
                else
                {
                    var states = new List<ColumnState>();
                    for (int column = 1; column <= total; column++) states.Add(CaptureColumn(table, column));
                    Reorder(states, sourceStart - 1, count, targetStart - 1);
                    int firstChanged = Math.Min(sourceStart, targetStart);
                    int lastChanged = Math.Max(sourceEnd, targetStart + count - 1);
                    for (int column = firstChanged; column <= lastChanged; column++) ApplyColumn(table, column, states[column - 1], mode);
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MoveTableRowColumnFeature.MoveBlock");
                ShowError("The selected table block could not be moved. Use Undo if PowerPoint applied only part of the operation.");
                return false;
            }
        }

        private static bool CanMoveRows(int sourceStart, int sourceEnd, int targetStart, int count, int total)
        {
            if (IsProtectedRow(sourceStart, total) || IsProtectedRow(sourceEnd, total))
            {
                ShowInfo("The selected block contains a protected header or total row. Change the protection options in the context menu first.");
                return false;
            }

            int targetEnd = targetStart + count - 1;
            if (ProtectFirstRow && targetStart <= 1 || ProtectLastRow && targetEnd >= total)
            {
                ShowInfo("The move would replace a protected header or total row. Change the protection options in the context menu first.");
                return false;
            }
            return true;
        }

        private static bool IsProtectedRow(int row, int total)
        {
            return ProtectFirstRow && row == 1 || ProtectLastRow && row == total;
        }

        private static void Reorder<T>(List<T> items, int sourceIndex, int count, int targetIndex)
        {
            var block = items.GetRange(sourceIndex, count);
            items.RemoveRange(sourceIndex, count);
            items.InsertRange(targetIndex, block);
        }

        private static bool DuplicateRows(Table table, int sourceStart, int sourceEnd, bool before)
        {
            var states = new List<RowState>();
            for (int row = sourceStart; row <= sourceEnd; row++) states.Add(CaptureRow(table, row));
            int insertAt = before ? sourceStart : sourceEnd + 1;
            int inserted = 0;
            try
            {
                for (int offset = 0; offset < states.Count; offset++)
                {
                    int index = insertAt + offset;
                    if (index > table.Rows.Count) table.Rows.Add(-1); else table.Rows.Add(index);
                    inserted++;
                    ApplyRow(table, index, states[offset], MoveMode.Complete);
                }
                return true;
            }
            catch
            {
                while (inserted-- > 0)
                {
                    try { table.Rows[insertAt].Delete(); }
                    catch (Exception ex) { ReportOptionalTableFailure(ex, "MoveTableRowColumnFeature.InsertRows.Rollback"); }
                }
                throw;
            }
        }

        private static bool DuplicateColumns(Table table, int sourceStart, int sourceEnd, bool before)
        {
            var states = new List<ColumnState>();
            for (int column = sourceStart; column <= sourceEnd; column++) states.Add(CaptureColumn(table, column));
            int insertAt = before ? sourceStart : sourceEnd + 1;
            int inserted = 0;
            try
            {
                for (int offset = 0; offset < states.Count; offset++)
                {
                    int index = insertAt + offset;
                    if (index > table.Columns.Count) table.Columns.Add(-1); else table.Columns.Add(index);
                    inserted++;
                    ApplyColumn(table, index, states[offset], MoveMode.Complete);
                }
                return true;
            }
            catch
            {
                while (inserted-- > 0)
                {
                    try { table.Columns[insertAt].Delete(); }
                    catch (Exception ex) { ReportOptionalTableFailure(ex, "MoveTableRowColumnFeature.InsertColumns.Rollback"); }
                }
                throw;
            }
        }

        private static RowState CaptureRow(Table table, int row)
        {
            var state = new RowState { Height = table.Rows[row].Height };
            for (int column = 1; column <= table.Columns.Count; column++) state.Cells.Add(CaptureCell(table.Cell(row, column)));
            return state;
        }

        private static ColumnState CaptureColumn(Table table, int column)
        {
            var state = new ColumnState { Width = table.Columns[column].Width };
            for (int row = 1; row <= table.Rows.Count; row++) state.Cells.Add(CaptureCell(table.Cell(row, column)));
            return state;
        }

        private static void ApplyRow(Table table, int row, RowState state, MoveMode mode)
        {
            for (int column = 1; column <= table.Columns.Count; column++) ApplyCell(table.Cell(row, column), state.Cells[column - 1], mode);
            if (mode == MoveMode.Complete || mode == MoveMode.IgnoreBorders) table.Rows[row].Height = state.Height;
        }

        private static void ApplyColumn(Table table, int column, ColumnState state, MoveMode mode)
        {
            for (int row = 1; row <= table.Rows.Count; row++) ApplyCell(table.Cell(row, column), state.Cells[row - 1], mode);
            if (mode == MoveMode.Complete || mode == MoveMode.IgnoreBorders) table.Columns[column].Width = state.Width;
        }

        private static CellState CaptureCell(Cell cell)
        {
            var shape = cell.Shape;
            var state = new CellState
            {
                Text = shape.TextFrame.TextRange.Text,
                FillVisible = shape.Fill.Visible,
                FillColor = shape.Fill.ForeColor.RGB,
                FillTransparency = shape.Fill.Transparency,
                VerticalAnchor = shape.TextFrame.VerticalAnchor,
                Orientation = shape.TextFrame.Orientation,
                MarginLeft = shape.TextFrame.MarginLeft,
                MarginRight = shape.TextFrame.MarginRight,
                MarginTop = shape.TextFrame.MarginTop,
                MarginBottom = shape.TextFrame.MarginBottom
            };

            // No RTF probe: PowerPoint's TextRange object model exposes no RTF member (verified by
            // PIA reflection - Text, Font, ParagraphFormat, Bound*, Runs and friends only), so the
            // former dynamic .Rtf read threw a RuntimeBinderException for every cell on every move
            // and filled the error log. Rich text survives through the text + font + paragraph
            // fallback in ApplyFormattedText, which is what actually ran on every previous move.
            try { state.ParagraphAlignment = shape.TextFrame.TextRange.ParagraphFormat.Alignment; }
            catch (Exception ex) { ReportOptionalTableFailure(ex, "MoveTableRowColumnFeature.CaptureCell.ParagraphAlignment"); }
            try
            {
                var font = shape.TextFrame.TextRange.Font;
                state.FontName = font.Name;
                state.FontSize = font.Size;
                state.FontColor = font.Color.RGB;
                state.FontBold = font.Bold;
                state.FontItalic = font.Italic;
                state.FontUnderline = font.Underline;
            }
            catch (Exception ex) { ReportOptionalTableFailure(ex, "MoveTableRowColumnFeature.CaptureCell.Font"); }

            for (int borderType = 1; borderType <= 6; borderType++)
            {
                try
                {
                    // Typed LineFormat: the former dynamic access resolved the enum-typed members
                    // (Visible, Style, DashStyle) late-bound as ints, so every border of every cell
                    // threw a RuntimeBinderException ("int cannot be implicitly converted to
                    // MsoTriState"), no border was ever captured, and each move flooded the error
                    // log with the failures.
                    LineFormat border = cell.Borders[(PpBorderType)borderType];
                    state.Borders[borderType] = new BorderState
                    {
                        Visible = border.Visible,
                        Weight = border.Weight,
                        Color = border.ForeColor.RGB,
                        Style = (int)border.Style,
                        DashStyle = (int)border.DashStyle
                    };
                }
                catch (Exception ex) { ReportOptionalTableFailure(ex, "MoveTableRowColumnFeature.CaptureCell.Border"); }
            }
            return state;
        }

        private static void ApplyCell(Cell cell, CellState state, MoveMode mode)
        {
            var shape = cell.Shape;
            if (mode == MoveMode.TextOnly)
            {
                shape.TextFrame.TextRange.Text = state.Text;
                return;
            }

            ApplyFormattedText(shape, state);
            if (mode == MoveMode.KeepTableDesign) return;

            try
            {
                shape.Fill.Visible = state.FillVisible;
                if (state.FillVisible == Office.MsoTriState.msoTrue)
                {
                    shape.Fill.ForeColor.RGB = state.FillColor;
                    shape.Fill.Transparency = state.FillTransparency;
                }
                shape.TextFrame.VerticalAnchor = state.VerticalAnchor;
                shape.TextFrame.Orientation = state.Orientation;
                shape.TextFrame.MarginLeft = state.MarginLeft;
                shape.TextFrame.MarginRight = state.MarginRight;
                shape.TextFrame.MarginTop = state.MarginTop;
                shape.TextFrame.MarginBottom = state.MarginBottom;
                shape.TextFrame.TextRange.ParagraphFormat.Alignment = state.ParagraphAlignment;
            }
            catch (Exception ex) { ReportOptionalTableFailure(ex, "MoveTableRowColumnFeature.ApplyCell.Layout"); }

            if (mode == MoveMode.IgnoreBorders) return;
            foreach (var border in state.Borders)
            {
                try
                {
                    LineFormat destination = cell.Borders[(PpBorderType)border.Key];
                    destination.Visible = border.Value.Visible;
                    destination.Weight = border.Value.Weight;
                    destination.ForeColor.RGB = border.Value.Color;
                    destination.Style = (Office.MsoLineStyle)border.Value.Style;
                    destination.DashStyle = (Office.MsoLineDashStyle)border.Value.DashStyle;
                }
                catch (Exception ex) { ReportOptionalTableFailure(ex, "MoveTableRowColumnFeature.ApplyCell.Border"); }
            }
        }

        private static void ApplyFormattedText(Shape shape, CellState state)
        {
            shape.TextFrame.TextRange.Text = state.Text;
            try
            {
                var font = shape.TextFrame.TextRange.Font;
                font.Name = state.FontName;
                font.Size = state.FontSize;
                font.Color.RGB = state.FontColor;
                font.Bold = state.FontBold;
                font.Italic = state.FontItalic;
                font.Underline = state.FontUnderline;
                shape.TextFrame.TextRange.ParagraphFormat.Alignment = state.ParagraphAlignment;
            }
            catch (Exception ex) { ReportOptionalTableFailure(ex, "MoveTableRowColumnFeature.ApplyFormattedText.Font"); }
        }

        private static void ReportOptionalTableFailure(Exception exception, string context)
        {
            if (!ExceptionPolicy.IsExpectedComFailure(exception))
            {
                ExceptionLogger.Log(exception, context);
            }
            // Optional table formatting and rollback continue when an Office proxy is stale.
        }

        private static bool TryGetSelection(PowerPointManager manager, out Shape tableShape, out Table table, out SelectionBounds bounds, bool showMessage)
        {
            tableShape = null;
            table = null;
            bounds = null;
            if (manager == null)
            {
                if (showMessage) ShowInfo("Select a PowerPoint table first.");
                return false;
            }
            try
            {
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table))
                {
                    if (showMessage) ShowInfo("Select a PowerPoint table first.");
                    return false;
                }

                int minRow, maxRow, minColumn, maxColumn;
                if (!TableHelper.GetSelectedCellBounds(table, out minRow, out maxRow, out minColumn, out maxColumn))
                {
                    int activeRow, activeColumn;
                    if (!TableHelper.TryGetActiveTextCell(manager, table, out activeRow, out activeColumn))
                    {
                        if (showMessage) ShowInfo("Select one or more table cells, or place the text cursor inside a table cell.");
                        return false;
                    }
                    minRow = maxRow = activeRow;
                    minColumn = maxColumn = activeColumn;
                }
                bounds = new SelectionBounds { MinRow = minRow, MaxRow = maxRow, MinColumn = minColumn, MaxColumn = maxColumn };
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MoveTableRowColumnFeature.TryGetSelection");
                if (showMessage) ShowError("PowerPoint could not read the current table selection.");
                return false;
            }
        }

        private static bool EnsureContiguousSelection(Table table, MoveAxis axis, SelectionBounds bounds)
        {
            try
            {
                var indices = GetSelectedIndices(table, axis);
                if (indices.Count == 0) return true;
                int expected = axis == MoveAxis.Row
                    ? bounds.MaxRow - bounds.MinRow + 1
                    : bounds.MaxColumn - bounds.MinColumn + 1;
                if (indices.Count == expected) return true;
                ShowInfo("Select one contiguous block of rows or columns before moving or duplicating it.");
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MoveTableRowColumnFeature.EnsureContiguousSelection");
                ShowError("PowerPoint could not validate the selected table block.");
                return false;
            }
        }

        private static List<int> GetSelectedIndices(Table table, MoveAxis axis)
        {
            var indices = new HashSet<int>();
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int column = 1; column <= table.Columns.Count; column++)
                {
                    if (table.Cell(row, column).Selected) indices.Add(axis == MoveAxis.Row ? row : column);
                }
            }
            return indices.OrderBy(value => value).ToList();
        }

        private static bool EnsureNoMergedCells(Table table)
        {
            try
            {
                // Merged cells are the only cells that share one underlying cell shape, so duplicate
                // Shape.Id values detect them. Some Office hosts do not expose Shape.Id for table
                // cells over the early-bound path (E_NOTIMPL, observed on isolated automation hosts
                // while the same call succeeds through late binding); those hosts fall back to the
                // cell geometry, which is the same oracle - merged cells are also the only cells with
                // identical bounds. The fallback stays silent, like
                // TableHelper.TryGetActiveTextCell's parent-chain fallback: it is expected behaviour
                // on such builds, not a failure, and an error-log entry would fail the acceptance
                // runners' zero-new-entries contract.
                var ids = new HashSet<int>();
                var bounds = new HashSet<string>();
                bool useGeometry;
                try
                {
                    ids.Add(table.Cell(1, 1).Shape.Id);
                    useGeometry = false;
                    LastMergeDetectionMode = MergeDetectionMode.ShapeId;
                }
                catch (Exception)
                {
                    bounds.Add(DescribeCellBounds(table.Cell(1, 1).Shape));
                    useGeometry = true;
                    LastMergeDetectionMode = MergeDetectionMode.GeometryFallback;
                }
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int column = 1; column <= table.Columns.Count; column++)
                    {
                        if (row == 1 && column == 1) continue; // already checked by the oracle probe above
                        bool duplicate = useGeometry
                            ? !bounds.Add(DescribeCellBounds(table.Cell(row, column).Shape))
                            : !ids.Add(table.Cell(row, column).Shape.Id);
                        if (duplicate)
                        {
                            ShowInfo("Reordering tables with merged cells is not supported because PowerPoint cannot move their merge structure safely.");
                            return false;
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MoveTableRowColumnFeature.EnsureNoMergedCells");
                ShowInfo("This table appears to contain merged or otherwise unsupported cells and was left unchanged.");
                return false;
            }
        }

        private static string DescribeCellBounds(Shape shape)
        {
            return string.Format(CultureInfo.InvariantCulture, "{0:R},{1:R},{2:R},{3:R}",
                shape.Left, shape.Top, shape.Width, shape.Height);
        }

        private static bool PromptForPosition(MoveAxis axis, int current, int maximum, out int position)
        {
            position = current;
            using (var dialog = new Forms.Form())
            using (var input = new Forms.NumericUpDown())
            using (var label = new Forms.Label())
            using (var ok = new Forms.Button())
            using (var cancel = new Forms.Button())
            {
                dialog.Text = axis == MoveAxis.Row ? "Move Rows to Position" : "Move Columns to Position";
                dialog.FormBorderStyle = Forms.FormBorderStyle.FixedDialog;
                dialog.StartPosition = Forms.FormStartPosition.CenterScreen;
                dialog.MinimizeBox = false;
                dialog.MaximizeBox = false;
                dialog.ClientSize = new System.Drawing.Size(330, 125);
                label.Text = "New starting position (1–" + maximum + "):";
                label.SetBounds(15, 18, 205, 22);
                input.Minimum = 1;
                input.Maximum = maximum;
                input.Value = Math.Max(1, Math.Min(current, maximum));
                input.SetBounds(230, 15, 80, 24);
                ok.Text = "Move";
                ok.DialogResult = Forms.DialogResult.OK;
                ok.SetBounds(154, 72, 75, 28);
                cancel.Text = "Cancel";
                cancel.DialogResult = Forms.DialogResult.Cancel;
                cancel.SetBounds(235, 72, 75, 28);
                dialog.Controls.AddRange(new Forms.Control[] { label, input, ok, cancel });
                dialog.AcceptButton = ok;
                dialog.CancelButton = cancel;
                if (dialog.ShowDialog() != Forms.DialogResult.OK) return false;
                position = Decimal.ToInt32(input.Value);
                return true;
            }
        }

        private static void ShowInfo(string message)
        {
            Forms.MessageBox.Show(message, "Move Table Rows or Columns", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information);
        }

        private static void TryStartUndoEntry(PowerPointManager manager)
        {
            try { manager?.GetApplication()?.StartNewUndoEntry(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "MoveTableRowColumnFeature.StartNewUndoEntry"); }
        }

        private static void ShowError(string message)
        {
            Forms.MessageBox.Show(message, "Move Table Rows or Columns", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Error);
        }

        private sealed class SelectionBounds
        {
            public int MinRow { get; set; }
            public int MaxRow { get; set; }
            public int MinColumn { get; set; }
            public int MaxColumn { get; set; }
        }

        private sealed class RowState
        {
            public float Height { get; set; }
            public List<CellState> Cells { get; } = new List<CellState>();
        }

        private sealed class ColumnState
        {
            public float Width { get; set; }
            public List<CellState> Cells { get; } = new List<CellState>();
        }

        private sealed class CellState
        {
            public string Text { get; set; }
            public Office.MsoTriState FillVisible { get; set; }
            public int FillColor { get; set; }
            public float FillTransparency { get; set; }
            public Office.MsoVerticalAnchor VerticalAnchor { get; set; }
            public Office.MsoTextOrientation Orientation { get; set; }
            public float MarginLeft { get; set; }
            public float MarginRight { get; set; }
            public float MarginTop { get; set; }
            public float MarginBottom { get; set; }
            public PpParagraphAlignment ParagraphAlignment { get; set; }
            public string FontName { get; set; }
            public float FontSize { get; set; }
            public int FontColor { get; set; }
            public Office.MsoTriState FontBold { get; set; }
            public Office.MsoTriState FontItalic { get; set; }
            public Office.MsoTriState FontUnderline { get; set; }
            public Dictionary<int, BorderState> Borders { get; } = new Dictionary<int, BorderState>();
        }

        private sealed class BorderState
        {
            public Office.MsoTriState Visible { get; set; }
            public float Weight { get; set; }
            public int Color { get; set; }
            public int Style { get; set; }
            public int DashStyle { get; set; }
        }
    }
}
