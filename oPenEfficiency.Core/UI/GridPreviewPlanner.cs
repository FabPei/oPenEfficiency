using System;
using System.Globalization;

namespace oPenEfficiency.UI
{
    /// <summary>
    /// Immutable schematic preview state for the Arrange Grid dialog. Pure planning data:
    /// the window renders it, and no PowerPoint state or COM access is involved.
    /// </summary>
    public sealed class GridPreviewState
    {
        public const int MaxVisibleColumns = 6;
        public const int MaxVisibleRows = 4;

        private GridPreviewState(bool isValid, int columns, int rows, int visibleColumns, int visibleRows, int? arrangedShapeCount)
        {
            IsValid = isValid;
            Columns = columns;
            Rows = rows;
            VisibleColumns = visibleColumns;
            VisibleRows = visibleRows;
            IsCondensed = isValid && (visibleColumns < columns || visibleRows < rows);
            IsSelectionAware = isValid && arrangedShapeCount.HasValue;
            ArrangedCellCount = !isValid
                ? 0
                : IsSelectionAware
                    ? Math.Min(arrangedShapeCount.Value, columns * rows)
                    : columns * rows;
            DimensionLabel = isValid
                ? string.Format(CultureInfo.InvariantCulture, "{0} \u00D7 {1}", columns, rows)
                : "\u2013 \u00D7 \u2013";
            AutomationName = !isValid
                ? "Grid preview, invalid dimensions"
                : IsSelectionAware
                    ? string.Format(CultureInfo.InvariantCulture,
                        "Grid preview, {0} columns by {1} rows, {2} of {3} cells filled",
                        columns, rows, ArrangedCellCount, columns * rows)
                    : string.Format(CultureInfo.InvariantCulture,
                        "Grid preview, {0} columns by {1} rows", columns, rows);
        }

        public bool IsValid { get; }
        public int Columns { get; }
        public int Rows { get; }
        public int VisibleColumns { get; }
        public int VisibleRows { get; }
        public bool IsCondensed { get; }
        /// <summary>True when a selected-shape count was supplied, so unfilled cells are meaningful.</summary>
        public bool IsSelectionAware { get; }
        /// <summary>Cells that will hold a shape. Without a count (unknown) all cells count as filled.</summary>
        public int ArrangedCellCount { get; }
        public int CellCount { get { return VisibleColumns * VisibleRows; } }
        public string DimensionLabel { get; }
        public string AutomationName { get; }

        internal static GridPreviewState Valid(int columns, int rows, int? arrangedShapeCount)
        {
            return new GridPreviewState(
                true, columns, rows,
                Math.Min(MaxVisibleColumns, columns),
                Math.Min(MaxVisibleRows, rows),
                arrangedShapeCount);
        }

        internal static GridPreviewState CreateInvalid()
        {
            return new GridPreviewState(false, 0, 0, 0, 0, null);
        }
    }

    /// <summary>
    /// Plans the Arrange Grid live preview from the current column/row inputs,
    /// including the condensed cell cap for large values, the neutral
    /// invalid-input state, and selection-aware unfilled cells when the
    /// selected-shape count is known.
    /// </summary>
    public static class GridPreviewPlanner
    {
        public static GridPreviewState Create(int columns, int rows)
        {
            return Create(columns, rows, null);
        }

        public static GridPreviewState Create(int columns, int rows, int? arrangedShapeCount)
        {
            if (columns < 1 || rows < 1) return GridPreviewState.CreateInvalid();
            return GridPreviewState.Valid(columns, rows, arrangedShapeCount);
        }

        public static GridPreviewState Create(string columnsText, string rowsText)
        {
            return Create(columnsText, rowsText, null);
        }

        public static GridPreviewState Create(string columnsText, string rowsText, int? arrangedShapeCount)
        {
            int columns, rows;
            if (string.IsNullOrWhiteSpace(columnsText) || string.IsNullOrWhiteSpace(rowsText) ||
                !int.TryParse(columnsText.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out columns) ||
                !int.TryParse(rowsText.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out rows))
            {
                return GridPreviewState.CreateInvalid();
            }
            return Create(columns, rows, arrangedShapeCount);
        }
    }
}
