using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert a new row above the current selection in a PowerPoint table.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnInsertRowTop",
        Name = "Insert Row Above",
        Tooltip = "Insert row above",
        IconData = "M3,3H21V21H3V3M5,5V19H19V5H5M7,7H17V9H7V7M7,11H17V13H7V11M7,15H17V17H7V15Z",
        Color = "#F43F5E",
        Description = "Inserts a new row above the selected row.",
        DetailedHelpText = "### Insert Row Above\nInserts a new empty row directly above the currently selected table row.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class InsertTableRowTopFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - inserts row above without preserving height.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, preserveHeight: false);
        }

        public static bool Execute(PowerPointManager manager, bool preserveHeight = false)
        {
            return TableRowInsertionFeature.Execute(manager, true, preserveHeight);
        }
    }
}
