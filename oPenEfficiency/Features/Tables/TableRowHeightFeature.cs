using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Copy row heights of a table.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTableRowHeight",
        Name = "Row Height",
        Tooltip = "Copy row height",
        IconData = "M3,3H21V21H3V3M7,5V19H17V11H11V5H7M13,13V19H7V13H13Z",
        Color = "#F43F5E",
        Description = "Copies row heights from the selected table.",
        DetailedHelpText = "### Row Height\nSets all rows in the selected table to a uniform height, evenly distributing the total table height across all rows.",
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
    }
}
