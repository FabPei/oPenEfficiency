using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert a new column to the right of the current selection in a PowerPoint table.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnInsertColumnRight",
        Name = "Insert Column Right",
        Tooltip = "Insert column right",
        IconData = "M3,3H21V21H3V3M5,5V19H9V5H5M11,5V19H13V5H11M15,5V19H17V5H15Z",
        Color = "#F43F5E",
        Description = "Inserts a new column to the right of the selected column.",
        DetailedHelpText = "### Insert Column Right\nInserts a new empty column immediately to the right of the currently selected table column.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class InsertTableColumnRightFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - inserts column right without preserving width.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, preserveWidth: false);
        }

        public static bool Execute(PowerPointManager manager, bool preserveWidth = false)
        {
            return TableColumnInsertionFeature.Execute(manager, false, preserveWidth);
        }
    }
}
