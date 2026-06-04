using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Copy and paste table dimensions (column widths and row heights).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTableDimensions",
        Name = "Table Dimensions",
        Tooltip = "Table dimensions",
        IconData = "M3,3H21V21H3V3M7,5V19H19V5H7M9,7H11V17H9V7M13,7H15V17H13V7M17,7H19V17H17V7Z",
        Color = "#F43F5E",
        Description = "Copies all table dimensions (column widths and row heights).",
        DetailedHelpText = "### Table Dimensions\nDisplays and allows direct numeric editing of the selected table's total width, height, column count, and row count.",
        Keywords = "size, dimensions, width, height, copy, paste",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableDimensionsFeature
    {
        private static float[] _copiedColumnWidths;
        private static float[] _copiedRowHeights;

        // Required zero-arg overload for FeatureDiscovery auto-registration.
        public static bool Execute(PowerPointManager manager)
            => CopyAll(manager);

        public static bool Execute(PowerPointManager manager, bool paste)
        {
            if (manager == null) return false;
            return paste ? PasteAll(manager) : CopyAll(manager);
        }

        public static bool CopyColumnWidths(PowerPointManager manager)
        {
            try
            {
                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                int colCount = table.Columns.Count;
                _copiedColumnWidths = new float[colCount];

                for (int i = 1; i <= colCount; i++)
                {
                    _copiedColumnWidths[i - 1] = table.Columns[i].Width;
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableDimensionsFeature.CopyColumnWidths");
                return false;
            }
        }

        public static bool PasteColumnWidths(PowerPointManager manager)
        {
            try
            {
                if (manager == null || _copiedColumnWidths == null || _copiedColumnWidths.Length == 0) return false;
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;
                var selection = app.ActiveWindow.Selection;
                if (selection == null || selection.Type != PpSelectionType.ppSelectionShapes) return false;
                var shapes = selection.ShapeRange;

                bool anyApplied = false;
                foreach (Shape s in shapes)
                {
                    try
                    {
                        if (s.HasTable == Office.MsoTriState.msoTrue)
                        {
                            var table = s.Table;
                            int colCount = table.Columns.Count;
                            int applyCount = Math.Min(colCount, _copiedColumnWidths.Length);

                            for (int i = 1; i <= applyCount; i++)
                            {
                                table.Columns[i].Width = _copiedColumnWidths[i - 1];
                            }
                            anyApplied = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "TableDimensionsFeature.PasteColumnWidths");
                    }
                }

                return anyApplied;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableDimensionsFeature.PasteColumnWidths");
                return false;
            }
        }

        public static bool CopyRowHeights(PowerPointManager manager)
        {
            try
            {
                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                int rowCount = table.Rows.Count;
                _copiedRowHeights = new float[rowCount];

                for (int i = 1; i <= rowCount; i++)
                {
                    _copiedRowHeights[i - 1] = table.Rows[i].Height;
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableDimensionsFeature.CopyRowHeights");
                return false;
            }
        }

        public static bool PasteRowHeights(PowerPointManager manager)
        {
            try
            {
                if (manager == null || _copiedRowHeights == null || _copiedRowHeights.Length == 0) return false;
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;
                var selection = app.ActiveWindow.Selection;
                if (selection == null || selection.Type != PpSelectionType.ppSelectionShapes) return false;
                var shapes = selection.ShapeRange;

                bool anyApplied = false;
                foreach (Shape s in shapes)
                {
                    try
                    {
                        if (s.HasTable == Office.MsoTriState.msoTrue)
                        {
                            var table = s.Table;
                            int rowCount = table.Rows.Count;
                            int applyCount = Math.Min(rowCount, _copiedRowHeights.Length);

                            for (int i = 1; i <= applyCount; i++)
                            {
                                table.Rows[i].Height = _copiedRowHeights[i - 1];
                            }
                            anyApplied = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "TableDimensionsFeature.PasteRowHeights");
                    }
                }

                return anyApplied;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableDimensionsFeature.PasteRowHeights");
                return false;
            }
        }

        public static bool CopyAll(PowerPointManager manager)
        {
            bool res1 = CopyColumnWidths(manager);
            bool res2 = CopyRowHeights(manager);
            return res1 && res2;
        }

        public static bool PasteAll(PowerPointManager manager)
        {
            bool res1 = PasteColumnWidths(manager);
            bool res2 = PasteRowHeights(manager);
            return res1 || res2;
        }
    }
}
