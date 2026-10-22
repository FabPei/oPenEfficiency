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
        Description = "Inserts one or more columns to the right of the selected column, with optional width preservation or format duplication.",
        DetailedHelpText = "### Insert Column Right\n\nInserts one column after the rightmost selected table column. The standard action grows the table. Right-click for **Insert Column Right & Preserve Width**, **Duplicate Selected Column Right – Format Only**, or **Insert Multiple Columns…**. Format-only duplication copies the selected column's width and basic cell formatting while leaving all new cells empty. Multiple insertion supports 2–50 columns and can optionally preserve the total table width.",
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

        public static bool ExecuteMultiple(PowerPointManager manager, int count, bool preserveWidth)
        {
            return TableColumnInsertionFeature.ExecuteMultiple(manager, false, preserveWidth, count);
        }

        public static bool DuplicateFormatOnly(PowerPointManager manager)
        {
            return TableColumnInsertionFeature.DuplicateColumnFormat(manager, false);
        }
    }
}
