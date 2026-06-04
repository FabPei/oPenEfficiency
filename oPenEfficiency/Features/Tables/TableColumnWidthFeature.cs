using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Copy column widths of a table.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTableColumnWidth",
        Name = "Column Width",
        Tooltip = "Copy column width",
        IconData = "M3,3H21V21H3V3M7,5V19H9V13H15V19H17V5H7M9,11V7H15V11H9Z",
        Color = "#F43F5E",
        Description = "Copies column widths from the selected table.",
        DetailedHelpText = "### Column Width\nSets all columns in the selected table to a uniform width, evenly dividing the total table width across all columns.",
        Keywords = "width, size, copy, dimension, column, match",
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
    }
}
