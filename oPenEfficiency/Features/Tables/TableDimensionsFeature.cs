using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>Copies and pastes table column widths and row heights.</summary>
    [FeatureMetadata(
        Id = "BtnTableDimensions",
        Name = "Copy/Paste Table Dimensions",
        SearchAliases = new[] { "Table Dimensions" },
        Tooltip = "Copy or paste table dimensions",
        IconData = "M3,3H21V21H3V3M7,5V19H19V5H7M9,7H11V17H9V7M13,7H15V17H13V7M17,7H19V17H17V7Z",
        Color = "#F43F5E",
        Description = "Copies and pastes a table's column widths, row heights, or both through the feature context menu.",
        DetailedHelpText = "### Copy/Paste Table Dimensions\n\nSelect a source table, then choose **Copy Column Widths**, **Copy Row Heights**, or **Copy Both** from the context menu. Select one or more target tables and choose the matching paste command. A target must have the same number of relevant columns or rows; mismatched tables are skipped instead of being changed partially. Copied values remain available for the current PowerPoint add-in session only.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableDimensionsFeature
    {
        private static float[] _copiedColumnWidths;
        private static float[] _copiedRowHeights;

        public static bool Execute(PowerPointManager manager) => CopyAll(manager);

        public static bool Execute(PowerPointManager manager, bool paste)
        {
            if (manager == null) return false;
            return paste ? PasteAll(manager) : CopyAll(manager);
        }

        public static bool CopyColumnWidths(PowerPointManager manager)
        {
            Shape shape;
            Table table;
            if (!TableHelper.GetSelectedTable(manager, out shape, out table)) return false;
            try
            {
                var widths = ReadColumnWidths(table);
                _copiedColumnWidths = widths;
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableDimensionsFeature.CopyColumnWidths");
                return false;
            }
        }

        public static bool CopyRowHeights(PowerPointManager manager)
        {
            Shape shape;
            Table table;
            if (!TableHelper.GetSelectedTable(manager, out shape, out table)) return false;
            try
            {
                var heights = ReadRowHeights(table);
                _copiedRowHeights = heights;
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableDimensionsFeature.CopyRowHeights");
                return false;
            }
        }

        public static bool CopyAll(PowerPointManager manager)
        {
            Shape shape;
            Table table;
            if (!TableHelper.GetSelectedTable(manager, out shape, out table)) return false;
            try
            {
                // Publish both buffers only after the complete read succeeds.
                var widths = ReadColumnWidths(table);
                var heights = ReadRowHeights(table);
                _copiedColumnWidths = widths;
                _copiedRowHeights = heights;
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableDimensionsFeature.CopyAll");
                return false;
            }
        }

        public static bool PasteColumnWidths(PowerPointManager manager) => Paste(manager, pasteWidths: true, pasteHeights: false);

        public static bool PasteRowHeights(PowerPointManager manager) => Paste(manager, pasteWidths: false, pasteHeights: true);

        public static bool PasteAll(PowerPointManager manager) => Paste(manager, pasteWidths: true, pasteHeights: true);

        public static bool EqualizeColumnWidths(PowerPointManager manager) => Equalize(manager, equalizeWidths: true);

        public static bool EqualizeRowHeights(PowerPointManager manager) => Equalize(manager, equalizeWidths: false);

        public static bool ClearCopiedColumnWidths()
        {
            bool hadValues = _copiedColumnWidths != null && _copiedColumnWidths.Length > 0;
            _copiedColumnWidths = null;
            return hadValues;
        }

        public static bool ClearCopiedRowHeights()
        {
            bool hadValues = _copiedRowHeights != null && _copiedRowHeights.Length > 0;
            _copiedRowHeights = null;
            return hadValues;
        }

        private static bool Equalize(PowerPointManager manager, bool equalizeWidths)
        {
            if (manager == null) return false;
            var targets = GetSelectedTables(manager);
            if (targets.Count == 0) return false;

            int applied = 0;
            int failed = 0;
            try { manager.GetApplication().StartNewUndoEntry(); } catch { /* Dimension changes remain valid when undo grouping is unavailable. */ }
            foreach (var table in targets)
            {
                float[] original = null;
                try
                {
                    original = equalizeWidths ? ReadColumnWidths(table) : ReadRowHeights(table);
                    float total = 0f;
                    for (int index = 0; index < original.Length; index++) total += original[index];
                    var equalValues = new float[original.Length];
                    float equalValue = total / original.Length;
                    for (int index = 0; index < equalValues.Length; index++) equalValues[index] = equalValue;

                    if (equalizeWidths) ApplyColumnWidths(table, equalValues);
                    else ApplyRowHeights(table, equalValues);
                    applied++;
                }
                catch (Exception ex)
                {
                    failed++;
                    ExceptionLogger.Log(ex, "TableDimensionsFeature.EqualizeTarget");
                    try
                    {
                        if (original != null)
                        {
                            if (equalizeWidths) ApplyColumnWidths(table, original);
                            else ApplyRowHeights(table, original);
                        }
                    }
                    catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "TableDimensionsFeature.EqualizeRollback"); }
                }
            }

            if (failed > 0)
                MessageBox.Show($"Dimensions equalized for {applied} table(s); {failed} table(s) failed and were restored.", "Copy/Paste Table Dimensions", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return applied > 0;
        }

        private static bool Paste(PowerPointManager manager, bool pasteWidths, bool pasteHeights)
        {
            if (manager == null) return false;
            if ((pasteWidths && (_copiedColumnWidths == null || _copiedColumnWidths.Length == 0)) ||
                (pasteHeights && (_copiedRowHeights == null || _copiedRowHeights.Length == 0)))
            {
                MessageBox.Show("Copy the requested table dimensions before pasting them.", "Copy/Paste Table Dimensions", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            var targets = GetSelectedTables(manager);
            if (targets.Count == 0) return false;

            int applied = 0;
            int skipped = 0;
            int failed = 0;
            try { manager.GetApplication().StartNewUndoEntry(); } catch { /* Dimension changes remain valid when undo grouping is unavailable. */ }

            foreach (var table in targets)
            {
                if ((pasteWidths && table.Columns.Count != _copiedColumnWidths.Length) ||
                    (pasteHeights && table.Rows.Count != _copiedRowHeights.Length))
                {
                    skipped++;
                    continue;
                }

                float[] oldWidths = null;
                float[] oldHeights = null;
                try
                {
                    oldWidths = pasteWidths ? ReadColumnWidths(table) : null;
                    oldHeights = pasteHeights ? ReadRowHeights(table) : null;
                    if (pasteWidths) ApplyColumnWidths(table, _copiedColumnWidths);
                    if (pasteHeights) ApplyRowHeights(table, _copiedRowHeights);
                    applied++;
                }
                catch (Exception ex)
                {
                    failed++;
                    ExceptionLogger.Log(ex, "TableDimensionsFeature.PasteTarget");
                    try
                    {
                        if (oldWidths != null) ApplyColumnWidths(table, oldWidths);
                        if (oldHeights != null) ApplyRowHeights(table, oldHeights);
                    }
                    catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "TableDimensionsFeature.RollbackTarget"); }
                }
            }

            if (skipped > 0 || failed > 0)
            {
                MessageBox.Show($"Dimensions applied to {applied} table(s). Skipped {skipped} table(s) with incompatible row or column counts; {failed} table(s) failed and were restored.", "Copy/Paste Table Dimensions", MessageBoxButtons.OK, failed > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
            return applied > 0;
        }

        private static List<Table> GetSelectedTables(PowerPointManager manager)
        {
            var result = new List<Table>();
            try
            {
                var selection = manager.GetApplication()?.ActiveWindow?.Selection;
                if (selection == null || (selection.Type != PpSelectionType.ppSelectionShapes && selection.Type != PpSelectionType.ppSelectionText)) return result;
                var range = selection.ShapeRange;
                for (int index = 1; index <= range.Count; index++)
                {
                    if (range[index].HasTable == Office.MsoTriState.msoTrue) result.Add(range[index].Table);
                }
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "TableDimensionsFeature.GetSelectedTables"); }
            return result;
        }

        private static float[] ReadColumnWidths(Table table)
        {
            var values = new float[table.Columns.Count];
            for (int index = 1; index <= table.Columns.Count; index++) values[index - 1] = table.Columns[index].Width;
            return values;
        }

        private static float[] ReadRowHeights(Table table)
        {
            var values = new float[table.Rows.Count];
            for (int index = 1; index <= table.Rows.Count; index++) values[index - 1] = table.Rows[index].Height;
            return values;
        }

        private static void ApplyColumnWidths(Table table, float[] values)
        {
            for (int index = 1; index <= values.Length; index++) table.Columns[index].Width = values[index - 1];
        }

        private static void ApplyRowHeights(Table table, float[] values)
        {
            for (int index = 1; index <= values.Length; index++) table.Rows[index].Height = values[index - 1];
        }
    }
}
