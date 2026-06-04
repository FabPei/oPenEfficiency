using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert a new column to the left of the current selection in a PowerPoint table.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnInsertColumnLeft",
        Name = "Insert Column Left",
        Tooltip = "Insert column left",
        IconData = "M3,3H21V21H3V3M7,5V19H9V5H7M11,5V19H13V5H11M15,5V19H19V5H15Z",
        Color = "#F43F5E",
        Description = "Inserts a new column to the left of the selected column.",
        DetailedHelpText = "### Insert Column Left\nInserts a new empty column immediately to the left of the currently selected table column, preserving all existing formatting.",
        Keywords = "add, column, table, left, insert, grid",
        MinSelection = 0,
        RequiresTable = true)]
    public static class InsertTableColumnLeftFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - inserts column left without preserving width.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, preserveWidth: false);
        }

        /// <summary>
        /// Executes the feature to insert a new column to the left of the current selection.
        /// </summary>
        /// <param name="manager">The PowerPointManager instance.</param>
        /// <param name="preserveWidth">If set to <c>true</c>, attempts to preserve the width of the table.</param>
        /// <returns><c>true</c> if the column was inserted successfully; otherwise, <c>false</c>.</returns>
        public static bool Execute(PowerPointManager manager, bool preserveWidth = false)
        {
            return TableColumnInsertionFeature.Execute(manager, true, preserveWidth);
        }
    }
}
