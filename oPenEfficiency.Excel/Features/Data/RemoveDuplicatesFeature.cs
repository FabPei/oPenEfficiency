using System;
using Xl = global::Microsoft.Office.Interop.Excel;
using oPenEfficiency.Excel.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Excel.Features.Data
{
    /// <summary>
    /// Removes duplicate rows within the currently selected cell range.
    /// First Excel-specific feature and the vertical-slice example for the
    /// shared [FeatureMetadata] auto-discovery mechanism.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnRemoveDuplicates",
        Name = "Remove Duplicates",
        Tooltip = "Remove duplicates in the selected range",
        IconData = "",
        Color = "#10B981",
        Description = "Removes duplicate rows from the currently selected cell range.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Cells)]
    public static class RemoveDuplicatesFeature
    {
        public static bool Execute(ExcelManager manager)
        {
            try
            {
                var range = manager?.SelectedRange;
                if (range == null || range.Columns.Count < 1)
                {
                    return false;
                }

                int colCount = range.Columns.Count;
                var columns = new object[colCount];
                for (int i = 0; i < colCount; i++)
                {
                    columns[i] = i + 1; // 1-based column indices within the range
                }

                // Assume the first row of the selection is a header.
                range.RemoveDuplicates(columns, Xl.XlYesNoGuess.xlYes);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "RemoveDuplicatesFeature.Execute");
                return false;
            }
        }
    }
}
