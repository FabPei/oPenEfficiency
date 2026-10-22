using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Copy row heights of a table.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTableRowHeight",
        Name = "Copy Row Heights",
        SearchAliases = new[] { "Row Height" },
        Tooltip = "Copy table row heights",
        IconData = "M2,2V22H14V2Z M4,4H12V8H4Z M4,10H12V14H4Z M4,16H12V20H4Z M19,2L15,6L16.4,7.4L18,5.8V18.2L16.4,16.6L15,18L19,22L23,18L21.6,16.6L20,18.2V5.8L21.6,7.4L23,6Z",
        Color = "#F43F5E",
        Description = "Copies all row heights from the selected table for later application to a compatible table.",
        DetailedHelpText = "### Copy Row Heights\n\nClick to copy every row height from the selected table for the current add-in session. Right-click to copy, paste, make all selected-table rows equally high while preserving total height, or clear the saved heights. Paste skips targets with a different row count instead of resizing them partially.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableRowHeightFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, paste: false);
        }

        public static bool Execute(PowerPointManager manager, bool paste = false)
        {
            if (manager == null) return false;
            return paste ? TableDimensionsFeature.PasteRowHeights(manager) : TableDimensionsFeature.CopyRowHeights(manager);
        }

        public static bool Equalize(PowerPointManager manager) => TableDimensionsFeature.EqualizeRowHeights(manager);
        public static bool ClearCopiedHeights() => TableDimensionsFeature.ClearCopiedRowHeights();
    }
}
