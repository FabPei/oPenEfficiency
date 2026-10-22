using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Copy column widths of a table.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTableColumnWidth",
        Name = "Copy Column Widths",
        SearchAliases = new[] { "Column Width" },
        Tooltip = "Copy table column widths",
        IconData = "M2,2H22V14H2Z M4,4V12H8V4Z M10,4V12H14V4Z M16,4V12H20V4Z M2,19L6,15L7.4,16.4L5.8,18H18.2L16.6,16.4L18,15L22,19L18,23L16.6,21.6L18.2,20H5.8L7.4,21.6L6,23Z",
        Color = "#F43F5E",
        Description = "Copies all column widths from the selected table for later application to a compatible table.",
        DetailedHelpText = "### Copy Column Widths\n\nClick to copy every column width from the selected table for the current add-in session. Right-click to copy, paste, make all selected-table columns equally wide while preserving total width, or clear the saved widths. Paste skips targets with a different column count instead of resizing them partially.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableColumnWidthFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, paste: false);
        }

        public static bool Execute(PowerPointManager manager, bool paste = false)
        {
            if (manager == null) return false;
            return paste ? TableDimensionsFeature.PasteColumnWidths(manager) : TableDimensionsFeature.CopyColumnWidths(manager);
        }

        public static bool Equalize(PowerPointManager manager) => TableDimensionsFeature.EqualizeColumnWidths(manager);
        public static bool ClearCopiedWidths() => TableDimensionsFeature.ClearCopiedColumnWidths();
    }
}
