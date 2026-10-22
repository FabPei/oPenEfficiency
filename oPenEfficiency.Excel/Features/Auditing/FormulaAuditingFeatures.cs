using System;
using Xl = global::Microsoft.Office.Interop.Excel;
using oPenEfficiency.Excel.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Excel.Features.Auditing
{
    /// <summary>
    /// Shows the native "trace dependents" arrows for the active cell
    /// (the same blue arrows as Formulas &gt; Trace Dependents in the ribbon).
    /// Fully reversible via <see cref="ClearArrowsFeature"/>; no cell content changes.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTraceDependents",
        Name = "Trace Dependents",
        Tooltip = "Show dependent arrows",
        IconData = "",
        Color = "#3B82F6",
        Description = "Displays tracer arrows showing which cells depend on the active cell (their formulas reference this cell).",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Cells)]
    public static class TraceDependentsFeature
    {
        public static bool Execute(ExcelManager manager)
        {
            try
            {
                var cell = AuditingHelper.GetActiveCell(manager);
                if (cell == null) return false;
                cell.ShowDependents();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TraceDependentsFeature.Execute");
                return false;
            }
        }
    }

    /// <summary>
    /// Shows the native "trace precedents" arrows for the active cell:
    /// which cells feed into the active cell's formula.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTracePrecedents",
        Name = "Trace Precedents",
        Tooltip = "Show precedent arrows",
        IconData = "",
        Color = "#3B82F6",
        Description = "Displays tracer arrows showing which cells feed into the active cell's formula.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Cells)]
    public static class TracePrecedentsFeature
    {
        public static bool Execute(ExcelManager manager)
        {
            try
            {
                var cell = AuditingHelper.GetActiveCell(manager);
                if (cell == null) return false;
                cell.ShowPrecedents();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TracePrecedentsFeature.Execute");
                return false;
            }
        }
    }

    /// <summary>
    /// Removes all tracer arrows from the active worksheet (Formulas &gt; Remove Arrows).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnClearArrows",
        Name = "Remove Arrows",
        Tooltip = "Remove all tracer arrows",
        IconData = "",
        Color = "#EF4444",
        Description = "Removes all visible tracer arrows (precedents and dependents) from the active worksheet.",
        RequiredType = OfficeSelectionType.None)]
    public static class ClearArrowsFeature
    {
        public static bool Execute(ExcelManager manager)
        {
            try
            {
                var sheet = manager?.ActiveSheet;
                if (sheet == null) return false;
                sheet.ClearArrows();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ClearArrowsFeature.Execute");
                return false;
            }
        }
    }

    internal static class AuditingHelper
    {
        /// <summary>
        /// Reduces the current selection to a single cell (top-left of the first area).
        /// ShowDependents/ShowPrecedents throw on multi-cell or multi-area selections,
        /// so we always operate on one cell.
        /// </summary>
        public static Xl.Range GetActiveCell(ExcelManager manager)
        {
            var range = manager?.SelectedRange;
            if (range == null) return null;

            try
            {
                var firstArea = range.Areas[1];
                return firstArea.Cells[1, 1] as Xl.Range;
            }
            catch
            {
                return range.Cells[1, 1] as Xl.Range;
            }
        }
    }
}
