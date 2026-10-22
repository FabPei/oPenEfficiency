using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;
using oPenEfficiency.Models;
using oPenEfficiency.Utils;
using Excel = Microsoft.Office.Interop.Excel;

namespace oPenEfficiency.Services
{
    public sealed class GanttExcelImportResult
    {
        public List<GanttTask> Tasks { get; } = new List<GanttTask>();
        public List<string> Warnings { get; } = new List<string>();
    }

    /// <summary>
    /// Manages bi-directional data flow between Excel and the Gantt Chart.
    /// Supports template generation, export, import, and live-sync.
    /// </summary>
    public static class GanttExcelManager
    {
        private const string HEADER_WBS = "No. (WBS)";
        private const string HEADER_TITLE = "Task Name";
        private const string HEADER_START = "Start Date";
        private const string HEADER_END = "End Date";
        private const string HEADER_DURATION = "Duration";
        private const string HEADER_PREDECESSORS = "Predecessors";
        private const string HEADER_PERCENT = "% Complete";
        private const string HEADER_MILESTONE_SHAPE = "Milestone Shape";
        private const string HEADER_TASK_SHAPE = "Task Shape";
        private const string HEADER_COLOR = "Color (Hex)";

        /// <summary>
        /// Creates a new Excel workbook with the standard Gantt data headers.
        /// </summary>
        public static void CreateTemplate(string filePath)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                worksheet.Name = "Gantt Data";

                // Set Headers
                string[] headers = { HEADER_WBS, HEADER_TITLE, HEADER_START, HEADER_END, HEADER_DURATION, HEADER_PREDECESSORS, HEADER_PERCENT, HEADER_MILESTONE_SHAPE, HEADER_TASK_SHAPE, HEADER_COLOR };
                for (int i = 0; i < headers.Length; i++)
                {
                    Excel.Range cell = null;
                    Excel.Font font = null;
                    Excel.Interior interior = null;
                    try
                    {
                        cell = (Excel.Range)worksheet.Cells[1, i + 1];
                        cell.Value2 = headers[i];
                        font = cell.Font;
                        font.Bold = true;
                        interior = cell.Interior;
                        interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    }
                    finally
                    {
                        if (interior != null) Marshal.ReleaseComObject(interior);
                        if (font != null) Marshal.ReleaseComObject(font);
                        if (cell != null) Marshal.ReleaseComObject(cell);
                    }
                }

                // Add some dummy data or formatting
                Excel.Range dateRange = null;
                Excel.Range percentRange = null;
                Excel.Range columns = null;
                try
                {
                    dateRange = worksheet.Range["C2", "D100"];
                    dateRange.NumberFormat = "yyyy-mm-dd";
                    percentRange = worksheet.Range["G2", "G100"];
                    percentRange.NumberFormat = "0%";
                    columns = worksheet.Columns;
                    columns.AutoFit();
                }
                finally
                {
                    if (columns != null) Marshal.ReleaseComObject(columns);
                    if (percentRange != null) Marshal.ReleaseComObject(percentRange);
                    if (dateRange != null) Marshal.ReleaseComObject(dateRange);
                }

                workbook.SaveAs(filePath);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttExcelManager.CreateTemplate");
                throw;
            }
            finally
            {
                Cleanup(excelApp, workbook, worksheet);
            }
        }

        /// <summary>
        /// Exports the current project tasks to an Excel file.
        /// </summary>
        public static void ExportToExcel(GanttProject project, string filePath)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                worksheet.Name = "Gantt Data";

                // Set Headers
                string[] headers = { HEADER_WBS, HEADER_TITLE, HEADER_START, HEADER_END, HEADER_DURATION, HEADER_PREDECESSORS, HEADER_PERCENT, HEADER_MILESTONE_SHAPE, HEADER_TASK_SHAPE, HEADER_COLOR };
                for (int i = 0; i < headers.Length; i++)
                {
                    Excel.Range cell = null;
                    Excel.Font font = null;
                    try
                    {
                        cell = (Excel.Range)worksheet.Cells[1, i + 1];
                        cell.Value2 = headers[i];
                        font = cell.Font;
                        font.Bold = true;
                    }
                    finally
                    {
                        if (font != null) Marshal.ReleaseComObject(font);
                        if (cell != null) Marshal.ReleaseComObject(cell);
                    }
                }

                // Populate Data
                object[,] data = new object[project.Tasks.Count, 10];
                for (int index = 0; index < project.Tasks.Count; index++)
                {
                    GanttTask task = project.Tasks[index];
                    data[index, 0] = task.WBS;
                    data[index, 1] = task.Title;
                    data[index, 2] = task.Start;
                    data[index, 3] = task.End;
                    data[index, 4] = task.DurationDays;
                    data[index, 5] = task.PredecessorsString;
                    data[index, 6] = task.PercentComplete;
                    data[index, 7] = task.MilestoneShape.ToString();
                    data[index, 8] = task.Shape.ToString();
                    data[index, 9] = task.CustomColor ?? "";
                }

                Excel.Range dataRange = null;
                Excel.Range startCell = null;
                Excel.Range endCell = null;
                Excel.Range columns = null;
                try
                {
                    if (project.Tasks.Count > 0)
                    {
                        startCell = (Excel.Range)worksheet.Cells[2, 1];
                        endCell = (Excel.Range)worksheet.Cells[project.Tasks.Count + 1, 10];
                        dataRange = worksheet.Range[startCell, endCell];
                        dataRange.Value2 = data;
                    }

                    for (int index = 0; index < project.Tasks.Count; index++)
                    {
                        if (!project.Tasks[index].IsSummary) continue;
                        Excel.Range rowRange = null;
                        Excel.Font rowFont = null;
                        try
                        {
                            rowRange = (Excel.Range)worksheet.Rows[index + 2];
                            rowFont = rowRange.Font;
                            rowFont.Bold = true;
                        }
                        finally
                        {
                            if (rowFont != null) Marshal.ReleaseComObject(rowFont);
                            if (rowRange != null) Marshal.ReleaseComObject(rowRange);
                        }
                    }

                    columns = worksheet.Columns;
                    columns.AutoFit();
                }
                finally
                {
                    if (columns != null) Marshal.ReleaseComObject(columns);
                    if (dataRange != null) Marshal.ReleaseComObject(dataRange);
                    if (endCell != null) Marshal.ReleaseComObject(endCell);
                    if (startCell != null) Marshal.ReleaseComObject(startCell);
                }
                workbook.SaveAs(filePath);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttExcelManager.ExportToExcel");
                throw;
            }
            finally
            {
                Cleanup(excelApp, workbook, worksheet);
            }
        }

        /// <summary>
        /// Imports or Syncs tasks from an Excel file.
        /// </summary>
        public static GanttExcelImportResult ImportFromExcel(string filePath, IEnumerable<GanttTask> existingTasks = null)
        {
            var result = new GanttExcelImportResult();
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range range = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Open(filePath, ReadOnly: true);
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                range = worksheet.UsedRange;

                object[,] values = range.Value2 as object[,];
                if (values == null)
                {
                    result.Warnings.Add("The worksheet does not contain a tabular data range.");
                    return result;
                }

                int rowCount = values.GetLength(0);
                int colCount = values.GetLength(1);

                // Map headers to column indices
                var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int c = 1; c <= colCount; c++)
                {
                    string header = values[1, c]?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(header)) headerMap[header] = c;
                }

                // Default mappings if headers aren't perfect
                int colWbs = headerMap.ContainsKey(HEADER_WBS) ? headerMap[HEADER_WBS] : 1;
                int colTitle = headerMap.ContainsKey(HEADER_TITLE) ? headerMap[HEADER_TITLE] : 2;
                int colStart = headerMap.ContainsKey(HEADER_START) ? headerMap[HEADER_START] : 3;
                int colEnd = headerMap.ContainsKey(HEADER_END) ? headerMap[HEADER_END] : 4;
                int colPred = headerMap.ContainsKey(HEADER_PREDECESSORS) ? headerMap[HEADER_PREDECESSORS] : 6;
                int colPercent = headerMap.ContainsKey(HEADER_PERCENT) ? headerMap[HEADER_PERCENT] : 7;
                int colMShape = headerMap.ContainsKey(HEADER_MILESTONE_SHAPE) ? headerMap[HEADER_MILESTONE_SHAPE] : 8;
                int colTShape = headerMap.ContainsKey(HEADER_TASK_SHAPE) ? headerMap[HEADER_TASK_SHAPE] : 9;
                int colColor = headerMap.ContainsKey(HEADER_COLOR) ? headerMap[HEADER_COLOR] : 10;

                var existingByWbs = (existingTasks ?? Enumerable.Empty<GanttTask>())
                    .Where(task => !string.IsNullOrWhiteSpace(task.WBS))
                    .GroupBy(task => task.WBS, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
                var existingByTitle = (existingTasks ?? Enumerable.Empty<GanttTask>())
                    .Where(task => !string.IsNullOrWhiteSpace(task.Title))
                    .GroupBy(task => task.Title, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

                for (int i = 2; i <= rowCount; i++)
                {
                    try
                    {
                        object titleValue = values[i, colTitle];
                        if (titleValue == null) continue;

                        string wbs = values[i, colWbs]?.ToString()?.Trim() ?? "";
                        string title = titleValue.ToString().Trim();
                        if (!TryParseDate(values[i, colStart], out DateTime start))
                        {
                            result.Warnings.Add($"Row {i}: invalid start date; row skipped.");
                            continue;
                        }
                        if (!TryParseDate(values[i, colEnd], out DateTime end) || end < start)
                        {
                            result.Warnings.Add($"Row {i}: invalid end date; row skipped.");
                            continue;
                        }
                        string pred = values[i, colPred]?.ToString()?.Trim() ?? "";
                        double percent = ParsePercent(values[i, colPercent]);
                        string mShapeStr = values[i, colMShape]?.ToString() ?? "Diamond";
                        string tShapeStr = values[i, colTShape]?.ToString() ?? "Rectangle";
                        string colorStr = values[i, colColor]?.ToString()?.Trim() ?? "";

                        MilestoneShape mShape = MilestoneShape.Diamond;
                        Enum.TryParse(mShapeStr, true, out mShape);

                        TaskShape tShape = TaskShape.Rectangle;
                        Enum.TryParse(tShapeStr, true, out tShape);

                        int outlineLevel = 0;
                        if (!string.IsNullOrEmpty(wbs))
                        {
                            outlineLevel = wbs.Count(c => c == '.') ;
                        }

                        GanttTask task = null;
                        if (!string.IsNullOrWhiteSpace(wbs)) existingByWbs.TryGetValue(wbs, out task);
                        if (task == null && !string.IsNullOrWhiteSpace(title)) existingByTitle.TryGetValue(title, out task);
                        task = CreateImportedTask(task);

                        task.WBS = wbs;
                        task.Title = title;
                        task.Start = start.Date;
                        task.End = end.Date;
                        task.PredecessorsString = pred;
                        task.PercentComplete = percent;
                        task.MilestoneShape = mShape;
                        task.Shape = tShape;
                        if (!string.IsNullOrWhiteSpace(colorStr)) task.CustomColor = colorStr;
                        task.OutlineLevel = outlineLevel;
                        result.Tasks.Add(task);
                    }
                    catch (Exception rowEx)
                    {
                        ExceptionLogger.Log(rowEx, $"GanttExcelManager.ImportRow {i}");
                        result.Warnings.Add($"Row {i}: {rowEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttExcelManager.ImportFromExcel");
                result.Warnings.Add("Excel import failed: " + ex.Message);
            }
            finally
            {
                if (range != null) Marshal.ReleaseComObject(range);
                Cleanup(excelApp, workbook, worksheet);
            }

            return result;
        }

        private static GanttTask CreateImportedTask(GanttTask existing)
        {
            if (existing == null) return new GanttTask { Id = Guid.NewGuid() };

            return new GanttTask
            {
                Id = existing.Id,
                Responsible = existing.Responsible,
                CustomColor = existing.CustomColor,
                CustomTextColor = existing.CustomTextColor,
                CustomBarText = existing.CustomBarText,
                CustomBarTextSize = existing.CustomBarTextSize,
                LabelBold = existing.LabelBold,
                LabelItalic = existing.LabelItalic,
                LabelUnderline = existing.LabelUnderline,
                LabelFontSize = existing.LabelFontSize,
                LabelFontName = existing.LabelFontName,
                IsHighlighted = existing.IsHighlighted,
                IsCollapsed = existing.IsCollapsed,
                IsEventLine = existing.IsEventLine,
                SubEvents = existing.SubEvents,
                SubBars = existing.SubBars
            };
        }

        private static bool TryParseDate(object value, out DateTime result)
        {
            result = default(DateTime);
            if (value == null) return false;
            try
            {
                if (value is double d)
                {
                    result = DateTime.FromOADate(d).Date;
                    return true;
                }
                if (value is DateTime dt)
                {
                    result = dt.Date;
                    return true;
                }
                if (DateTime.TryParse(value.ToString(), out DateTime parsed))
                {
                    result = parsed.Date;
                    return true;
                }
            }
            catch (ArgumentException)
            {
                // Invalid external date values are reported as an unsuccessful parse.
            }
            catch (OverflowException)
            {
                // Excel serial dates outside DateTime's range are invalid input.
            }
            return false;
        }

        private static double ParsePercent(object value)
        {
            if (value == null) return 0;
            double parsed;
            if (value is double d) parsed = d;
            else if (!double.TryParse(value.ToString(), out parsed)) return 0;
            if (parsed > 1 && parsed <= 100) parsed /= 100d;
            return Math.Max(0, Math.Min(1, parsed));
        }

        private static void Cleanup(Excel.Application app, Excel.Workbook wb, Excel.Worksheet ws)
        {
            if (ws != null) Marshal.ReleaseComObject(ws);
            if (wb != null)
            {
                wb.Close(false);
                Marshal.ReleaseComObject(wb);
            }
            if (app != null)
            {
                app.Quit();
                Marshal.ReleaseComObject(app);
            }
        }
    }
}
