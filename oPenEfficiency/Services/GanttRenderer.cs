using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Orchestrates the visual rendering of the Gantt chart onto a PowerPoint slide.
    /// Includes a split-layout (Data Table + Gantt Timeline) and performance optimizations.
    /// </summary>
    public static class GanttRenderer
    {
        public const double COL_NO_WIDTH = 30;
        public const double COL_NAME_WIDTH = 170;
        public const double COL_DURATION_WIDTH = 50;

        public const double ROW_HEIGHT = 25;
        public const double BAR_HEIGHT = 15;
        public const string GANTT_GROUP_NAME = "GanttChart_Project";
        // PowerPoint Interop uses points; 0.1 cm equals 2.8346457 pt.
        private const float TASK_NAME_TEXT_MARGIN = 2.83465f;

        public static double GetRowHeight(GanttTask task, GanttProject project)
        {
            if (task == null || project == null || !project.ShowTaskNameColumn || string.IsNullOrEmpty(task.Title))
                return ROW_HEIGHT;

            double fontSize = task.LabelFontSize ?? 9;
            double indent = task.OutlineLevel * 15;
            double usableWidth = Math.Max(10, COL_NAME_WIDTH - indent - (2 * TASK_NAME_TEXT_MARGIN));
            double averageCharacterWidth = fontSize * ((task.LabelBold ?? task.IsSummary) ? 0.58 : 0.55);
            int charactersPerLine = Math.Max(1, (int)Math.Floor(usableWidth / averageCharacterWidth));
            int lineCount = EstimateWrappedLineCount(task.Title, charactersPerLine);

            double wrappedTextHeight = (2 * TASK_NAME_TEXT_MARGIN) + (lineCount * fontSize * 1.3);
            return Math.Max(ROW_HEIGHT, Math.Ceiling(wrappedTextHeight));
        }

        private static int EstimateWrappedLineCount(string text, int charactersPerLine)
        {
            int totalLines = 0;
            string normalizedText = text.Replace("\r\n", "\n").Replace('\r', '\n');

            foreach (string paragraph in normalizedText.Split('\n'))
            {
                string[] words = paragraph.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (words.Length == 0)
                {
                    totalLines++;
                    continue;
                }

                int currentLineLength = 0;
                foreach (string word in words)
                {
                    int remainingWordLength = word.Length;
                    while (remainingWordLength > charactersPerLine)
                    {
                        if (currentLineLength > 0)
                        {
                            totalLines++;
                            currentLineLength = 0;
                        }

                        totalLines++;
                        remainingWordLength -= charactersPerLine;
                    }

                    if (remainingWordLength == 0) continue;

                    int requiredLength = currentLineLength == 0
                        ? remainingWordLength
                        : currentLineLength + 1 + remainingWordLength;

                    if (requiredLength > charactersPerLine)
                    {
                        totalLines++;
                        currentLineLength = remainingWordLength;
                    }
                    else
                    {
                        currentLineLength = requiredLength;
                    }
                }

                if (currentLineLength > 0) totalLines++;
            }

            return Math.Max(1, totalLines);
        }

        public static int GetTaskIndexAtVerticalOffset(GanttProject project, double verticalOffset)
        {
            if (project == null || verticalOffset < 0) return -1;

            double currentOffset = 0;
            for (int i = 0; i < project.Tasks.Count; i++)
            {
                GanttTask task = project.Tasks[i];
                if (!task.IsVisible) continue;

                currentOffset += GetRowHeight(task, project);
                if (verticalOffset < currentOffset) return i;
            }

            return -1;
        }

        public static double GetHeaderHeight(GanttProject project)
        {
            int levelsCount = 0;
            if (project.ShowYears) levelsCount++;
            if (project.ShowMonths) levelsCount++;
            if (project.ShowWeeks) levelsCount++;
            if (project.ShowDays) levelsCount++;
            return Math.Max(20, levelsCount * 20);
        }

        public static double GetTableWidth(GanttProject project)
        {
            double tableWidth = 0;
            if (project.ShowNrColumn) tableWidth += COL_NO_WIDTH;
            if (project.ShowTaskNameColumn) tableWidth += COL_NAME_WIDTH;
            if (project.ShowDurationColumn) tableWidth += COL_DURATION_WIDTH;
            return tableWidth;
        }

        public static double GetPointsPerDay(GanttProject project, double totalWidth)
        {
            double tableWidth = GetTableWidth(project);
            double exactDays = (project.ProjectEnd - project.ProjectStart).TotalDays;
            double totalDays = Math.Max(1, Math.Ceiling(exactDays));

            if (project.AutoFitToSlide)
            {
                double availableGanttWidth = totalWidth - tableWidth;
                if (availableGanttWidth < 50) availableGanttWidth = 50;
                return availableGanttWidth / totalDays;
            }

            if (project.ShowDays) return 20;
            if (project.ShowWeeks) return 5;
            if (project.ShowMonths) return 1.5;
            if (project.ShowYears) return 0.4;

            return 20;
        }

        public static Shape Render(Slide slide, GanttProject project, double topOffset, double leftMargin, double totalWidth)
        {
            if (slide == null || project == null) return null;

            try
            {
                // Anchor Logic: if this chart instance already exists, use its current position.
                // A slide may contain multiple independent Gantt/timeline charts.
                double effectiveTop = topOffset;
                double effectiveLeft = leftMargin;
                Shape existingGroup = PowerPointShapeManager.FindChartGroup(slide, project.ChartId);
                if (existingGroup != null)
                {
                    effectiveTop = existingGroup.Top;
                    effectiveLeft = existingGroup.Left;
                }

                double tableWidth = GetTableWidth(project);

                double exactDays = (project.ProjectEnd - project.ProjectStart).TotalDays;
                double totalDays = Math.Max(1, Math.Ceiling(exactDays));
                DateTime visualProjectEnd = project.ProjectStart.AddDays(totalDays);

                // Dynamic scale factor based on which finest level is enabled
                double pointsPerDay = GetPointsPerDay(project, totalWidth);

                double ganttTotalWidth = totalDays * pointsPerDay;
                double ganttLeftMargin = effectiveLeft + tableWidth;

                // 1. Read user-modified colors from existing shapes before clearing
                ReadColorsFromSlide(existingGroup, project);

                // 2. Clean existing shapes
                if (existingGroup != null) Marshal.ReleaseComObject(existingGroup);
                PowerPointShapeManager.ClearGanttShapes(slide, project.ChartId);
                var preExistingShapeNames = CaptureTopLevelShapeNames(slide);

                // 2. Draw Headers (Table Header + Timeline Header)
                double dynamicHeaderHeight = GetHeaderHeight(project);
                DrawHeaders(slide, project, visualProjectEnd, effectiveLeft, effectiveTop, tableWidth, pointsPerDay, dynamicHeaderHeight);

                double timelineTop = effectiveTop + dynamicHeaderHeight;
                double totalTimelineHeight = project.Tasks
                    .Where(t => t.IsVisible)
                    .Sum(t => GetRowHeight(t, project));

                // 3. Draw Table Area & Background Grid
                DrawTableArea(slide, project, effectiveLeft, timelineTop, totalTimelineHeight, tableWidth, ganttTotalWidth);
                DrawGrid(slide, project, visualProjectEnd, ganttLeftMargin, timelineTop, totalHeight: totalTimelineHeight, pointsPerDay);

                // Draw "today" vertical marker if today falls within the chart range
                if (project.ProjectStart.Date <= DateTime.Today && DateTime.Today < visualProjectEnd.Date)
                {
                    double todayX = ganttLeftMargin + (DateTime.Today - project.ProjectStart.Date).TotalDays * pointsPerDay;
                    Shape todayLine = slide.Shapes.AddLine((float)todayX, (float)timelineTop, (float)todayX, (float)(timelineTop + totalTimelineHeight));
                    todayLine.Line.ForeColor.RGB = ColorMath.HexToBgr("#E74C3C");
                    todayLine.Line.Weight = 1.5f;
                    todayLine.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;
                    PowerPointShapeManager.TagShape(todayLine, Guid.Empty, PowerPointShapeManager.GanttShapeType.GridLine);
                    Marshal.ReleaseComObject(todayLine);
                }

                // 4. Draw/Update Tasks
                double currentTop = timelineTop;
                var renderedTaskShapes = new Dictionary<Guid, Shape>();

                // Top border above the first task row
                Shape topBorder = slide.Shapes.AddLine((float)effectiveLeft, (float)timelineTop, (float)(effectiveLeft + tableWidth + ganttTotalWidth), (float)timelineTop);
                topBorder.Line.ForeColor.RGB = 0xE0E0E0;
                topBorder.Line.Weight = 1.0f;
                PowerPointShapeManager.TagShape(topBorder, Guid.Empty, PowerPointShapeManager.GanttShapeType.Label);
                Marshal.ReleaseComObject(topBorder);

                // Right outer boundary: a vertical line exactly at the timeline extent - the same x
                // the horizontal borders and row hit areas already reach, so the group's bounds are
                // unchanged. Rerender, resize, date-range changes and copies move it with the extent.
                double rightBoundaryX = effectiveLeft + tableWidth + ganttTotalWidth;
                Shape rightBoundary = slide.Shapes.AddLine((float)rightBoundaryX, (float)timelineTop,
                    (float)rightBoundaryX, (float)(timelineTop + totalTimelineHeight));
                rightBoundary.Line.ForeColor.RGB = 0xE0E0E0;
                rightBoundary.Line.Weight = 1.0f;
                PowerPointShapeManager.TagShape(rightBoundary, Guid.Empty, PowerPointShapeManager.GanttShapeType.Label);
                Marshal.ReleaseComObject(rightBoundary);

                for (int i = 0; i < project.Tasks.Count; i++)
                {
                    var task = project.Tasks[i];
                    if (!task.IsVisible) continue;
                    double rowHeight = GetRowHeight(task, project);

                    // Invisible full-height rectangle so right-click hit-testing works anywhere in the row.
                    // AddLine shapes have Height=0, giving only a ~5pt effective hit zone; this covers the full lane.
                    Shape rowHitArea = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                        (float)effectiveLeft, (float)currentTop,
                        (float)(tableWidth + ganttTotalWidth), (float)rowHeight);
                    rowHitArea.Fill.ForeColor.RGB = task.IsHighlighted ? ColorMath.HexToBgr("#FFF9C4") : 0xFFFFFF;
                    rowHitArea.Fill.Transparency = 0f;
                    rowHitArea.Line.Visible = Office.MsoTriState.msoFalse;
                    PowerPointShapeManager.TagShape(rowHitArea, Guid.Empty, PowerPointShapeManager.GanttShapeType.GridLine);
                    Marshal.ReleaseComObject(rowHitArea);

                    // Draw the table row text
                    DrawTableRow(slide, task, i, currentTop, rowHeight, effectiveLeft, project);

                    // Draw the Gantt shape
                    Shape taskShape = null;
                    if (task.IsEventLine)
                    {
                        taskShape = DrawEventLine(slide, task, project, timelineTop, totalTimelineHeight, ganttLeftMargin, pointsPerDay);
                    }
                    else if (task.IsMilestone)
                    {
                        taskShape = DrawMilestone(slide, task, project, currentTop, rowHeight, ganttLeftMargin, pointsPerDay);
                    }
                    else if (task.IsSummary)
                    {
                        taskShape = DrawSummaryBar(slide, task, project, currentTop, rowHeight, ganttLeftMargin, pointsPerDay);
                    }
                    else
                    {
                        taskShape = DrawTaskBar(slide, task, project, currentTop, rowHeight, ganttLeftMargin, pointsPerDay);
                    }

                    if (taskShape != null)
                    {
                        renderedTaskShapes[task.Id] = taskShape;
                    }

                    // Draw extra roadmap events for this lane if they exist
                    DrawSubEvents(slide, task, project, currentTop, rowHeight, ganttLeftMargin, pointsPerDay);
                    DrawSubBars(slide, task, project, currentTop, rowHeight, ganttLeftMargin, pointsPerDay);

                    // Optional: Draw subtle horizontal row separator
                    Shape separator = slide.Shapes.AddLine((float)effectiveLeft, (float)(currentTop + rowHeight), (float)(effectiveLeft + tableWidth + ganttTotalWidth), (float)(currentTop + rowHeight));
                    separator.Line.ForeColor.RGB = 0xE0E0E0; // Very light grey
                    PowerPointShapeManager.TagShape(separator, Guid.Empty, PowerPointShapeManager.GanttShapeType.Label);
                    Marshal.ReleaseComObject(separator);

                    currentTop += rowHeight;
                }

                // 5. Draw Dependencies (Arrows)
                try
                {
                    DrawDependencies(slide, project, renderedTaskShapes);
                }
                finally
                {
                    foreach (Shape renderedShape in renderedTaskShapes.Values)
                    {
                        if (renderedShape != null) Marshal.ReleaseComObject(renderedShape);
                    }
                }

                // 6. Draw Legend (below everything)
                if (project.ShowLegend)
                {
                    DrawLegend(slide, project, currentTop + 20, effectiveLeft, project.FontFamily);
                }

                // 7. Final Layering (do this before grouping)
                PowerPointShapeManager.ReorderZLayers(slide);

                // 8. Group all Gantt shapes
                Shape group = GroupGanttShapes(slide, project.ChartId, preExistingShapeNames, pointsPerDay, project.ProjectStart);

                // 9. Apply Locks if requested
                if (group != null && project.LockElements)
                {
                    try
                    {
                        // In modern PowerPoint (Office 365), shapes have a "Locked" property.
                        // It's not in the standard Interop interface, so we use 'dynamic' to set it.
                        // We lock the individual items so they can't be accidentally dragged,
                        // but leave the main group unlocked so the whole chart can still be moved.
                        int groupItemCount = group.GroupItems.Count;
                        for (int i = 1; i <= groupItemCount; i++)
                        {
                            Shape subShape = null;
                            try
                            {
                                subShape = group.GroupItems[i];
                                dynamic dynShape = subShape;
                                dynShape.Locked = true;
                            }
                            catch (COMException)
                            {
                                // The optional Locked property is unavailable in older Office versions.
                            }
                            catch (MissingMemberException)
                            {
                                // The dynamic Office shape does not expose Locked.
                            }
                            finally
                            {
                                if (subShape != null) Marshal.ReleaseComObject(subShape);
                            }
                        }
                    }
                    catch (COMException)
                    {
                        // Locking is optional; rendering remains valid without it.
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "GanttRenderer.Render.ApplyLocks");
                    }
                }

                return group;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttRenderer.Render");
                return null;
            }
        }

        private static void ReadColorsFromSlide(Shape group, GanttProject project)
        {
            if (group == null || project == null || !project.Tasks.Any()) return;

            var taskById = project.Tasks.ToDictionary(t => t.Id.ToString(), t => t);

            try
            {
                int count = group.GroupItems.Count;
                for (int i = 1; i <= count; i++)
                {
                    Shape shape = null;
                    try
                    {
                        shape = group.GroupItems[i];

                        string idStr = PowerPointShapeManager.GetGanttId(shape);
                        if (string.IsNullOrEmpty(idStr) || idStr == Guid.Empty.ToString()) continue;
                        if (!taskById.TryGetValue(idStr, out Models.GanttTask task)) continue;

                        string typeStr = PowerPointShapeManager.GetGanttType(shape);
                        if (!Enum.TryParse(typeStr, out PowerPointShapeManager.GanttShapeType shapeType)) continue;

                        if (shapeType == PowerPointShapeManager.GanttShapeType.TaskBar ||
                            shapeType == PowerPointShapeManager.GanttShapeType.Milestone ||
                            shapeType == PowerPointShapeManager.GanttShapeType.SummaryBar)
                        {
                            try
                            {
                                var color = System.Drawing.ColorTranslator.FromOle(shape.Fill.ForeColor.RGB);
                                task.CustomColor = System.Drawing.ColorTranslator.ToHtml(color);
                            }
                            catch (COMException)
                            {
                                // Optional persisted fill formatting is unavailable.
                            }

                            // Preserve any text the user typed directly into the shape
                            try
                            {
                                string txt = shape.TextFrame.TextRange.Text;
                                task.CustomBarText = string.IsNullOrEmpty(txt) ? null : txt;
                                if (!string.IsNullOrEmpty(txt))
                                {
                                    float sz = shape.TextFrame.TextRange.Font.Size;
                                    if (sz > 0) task.CustomBarTextSize = sz;
                                }
                            }
                            catch (COMException)
                            {
                                // The task bar has no readable text frame.
                            }
                        }
                        else if (shapeType == PowerPointShapeManager.GanttShapeType.RowAnchor)
                        {
                            var font = shape.TextFrame.TextRange.Font;

                            // Color — only when explicitly set (not automatic/theme)
                            try
                            {
                                var fontColor = font.Color;
                                if (fontColor.Type == Office.MsoColorType.msoColorTypeRGB)
                                {
                                    var color = System.Drawing.ColorTranslator.FromOle(fontColor.RGB);
                                    task.CustomTextColor = System.Drawing.ColorTranslator.ToHtml(color);
                                }
                            }
                            catch (COMException)
                            {
                                // Theme or automatic font colors are optional input.
                            }

                            // Bold
                            try
                            {
                                var v = font.Bold;
                                task.LabelBold = v == Office.MsoTriState.msoTrue ? true : v == Office.MsoTriState.msoFalse ? false : (bool?)null;
                            }
                            catch (COMException)
                            {
                                // Mixed or unavailable font state leaves the persisted value unchanged.
                            }

                            // Italic
                            try
                            {
                                var v = font.Italic;
                                task.LabelItalic = v == Office.MsoTriState.msoTrue ? true : v == Office.MsoTriState.msoFalse ? false : (bool?)null;
                            }
                            catch (COMException)
                            {
                                // Mixed or unavailable font state leaves the persisted value unchanged.
                            }

                            // Underline
                            try
                            {
                                var v = font.Underline;
                                task.LabelUnderline = v == Office.MsoTriState.msoTrue ? true : v == Office.MsoTriState.msoFalse ? false : (bool?)null;
                            }
                            catch (COMException)
                            {
                                // Mixed or unavailable font state leaves the persisted value unchanged.
                            }

                            // Font size
                            try
                            {
                                float sz = font.Size;
                                if (sz > 0) task.LabelFontSize = sz;
                            }
                            catch (COMException)
                            {
                                // Mixed or unavailable font state leaves the persisted value unchanged.
                            }

                            // Font name
                            try
                            {
                                string name = font.Name;
                                if (!string.IsNullOrEmpty(name)) task.LabelFontName = name;
                            }
                            catch (COMException)
                            {
                                // Mixed or unavailable font state leaves the persisted value unchanged.
                            }
                        }
                    }
                    catch (COMException)
                    {
                        // A deleted child shape is skipped while synchronizing optional formatting.
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, $"GanttRenderer.ReadColorsFromSlide.Shape[{i}]");
                    }
                    finally
                    {
                        if (shape != null) Marshal.ReleaseComObject(shape);
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttRenderer.ReadColorsFromSlide");
            }
        }

        private static void DrawHeaders(Slide slide, GanttProject project, DateTime visualProjectEnd, double leftMargin, double top, double tableWidth, double pointsPerDay, double totalHeaderHeight)
        {
            double ganttLeft = leftMargin + tableWidth;
            double currentX = leftMargin;

            // Table Headers (align with the bottom header row, just above the first task)
            double tableHeaderTop = top + totalHeaderHeight - 20;
            if (project.ShowNrColumn)
            {
                DrawTableHeaderCell(slide, currentX, tableHeaderTop, COL_NO_WIDTH, "No.", project.FontFamily);
                currentX += COL_NO_WIDTH;
            }

            if (project.ShowTaskNameColumn)
            {
                DrawTableHeaderCell(slide, currentX, tableHeaderTop, COL_NAME_WIDTH, "Task Name", project.FontFamily, PowerPoint.PpParagraphAlignment.ppAlignLeft);
                currentX += COL_NAME_WIDTH;
            }

            if (project.ShowDurationColumn)
            {
                DrawTableHeaderCell(slide, currentX, tableHeaderTop, COL_DURATION_WIDTH, "Duration", project.FontFamily);
            }

            if (project.Tasks.Count == 0) return;

            // Draw active timeline levels
            double currentLevelTop = top;
            if (project.ShowYears)
            {
                DrawYearLevel(slide, project, visualProjectEnd, currentLevelTop, ganttLeft, pointsPerDay);
                currentLevelTop += 20;
            }
            if (project.ShowMonths)
            {
                DrawMonthLevel(slide, project, visualProjectEnd, currentLevelTop, ganttLeft, pointsPerDay);
                currentLevelTop += 20;
            }
            if (project.ShowWeeks)
            {
                DrawWeekLevel(slide, project, visualProjectEnd, currentLevelTop, ganttLeft, pointsPerDay);
                currentLevelTop += 20;
            }
            if (project.ShowDays)
            {
                DrawDayLevel(slide, project, visualProjectEnd, currentLevelTop, ganttLeft, pointsPerDay);
                currentLevelTop += 20;
            }
        }

        private static void DrawYearLevel(Slide slide, GanttProject project, DateTime visualProjectEnd, double top, double ganttLeft, double pointsPerDay)
        {
            for (int y = project.ProjectStart.Year; y <= visualProjectEnd.Year; y++)
            {
                DateTime segmentStart = new DateTime(y, 1, 1);
                DateTime segmentEnd = new DateTime(y + 1, 1, 1);
                if (segmentStart < project.ProjectStart) segmentStart = project.ProjectStart;
                if (segmentEnd > visualProjectEnd) segmentEnd = visualProjectEnd;

                var bounds = GanttMathEngine.CalculateBounds(segmentStart, segmentEnd, project.ProjectStart, ganttLeft, pointsPerDay);
                if (bounds.Width > 10)
                {
                    DrawHeaderLabel(slide, bounds.X, top, bounds.Width, 20, y.ToString(), project.FontFamily, true, 0xDDDDDD);
                }
            }
        }

        private static void DrawMonthLevel(Slide slide, GanttProject project, DateTime visualProjectEnd, double top, double ganttLeft, double pointsPerDay)
        {
            DateTime currentMonth = new DateTime(project.ProjectStart.Year, project.ProjectStart.Month, 1);
            while (currentMonth <= visualProjectEnd)
            {
                DateTime nextMonth = currentMonth.AddMonths(1);
                DateTime segmentStart = currentMonth < project.ProjectStart ? project.ProjectStart : currentMonth;
                DateTime segmentEnd = nextMonth > visualProjectEnd ? visualProjectEnd : nextMonth;

                var bounds = GanttMathEngine.CalculateBounds(segmentStart, segmentEnd, project.ProjectStart, ganttLeft, pointsPerDay);
                if (bounds.Width > 5)
                {
                    string label = bounds.Width > 35 ? currentMonth.ToString("MMMM", CultureInfo.InvariantCulture) : currentMonth.ToString("MMM", CultureInfo.InvariantCulture);
                    DrawHeaderLabel(slide, bounds.X, top, bounds.Width, 20, label, project.FontFamily, false, 0xEEEEEE, 8);
                }
                currentMonth = nextMonth;
            }
        }

        private static void DrawWeekLevel(Slide slide, GanttProject project, DateTime visualProjectEnd, double top, double ganttLeft, double pointsPerDay)
        {
            DateTime currentDay = project.ProjectStart.Date;
            while (currentDay.DayOfWeek != DayOfWeek.Monday) currentDay = currentDay.AddDays(-1);

            while (currentDay <= visualProjectEnd)
            {
                DateTime nextWeek = currentDay.AddDays(7);
                DateTime segmentStart = currentDay < project.ProjectStart ? project.ProjectStart : currentDay;
                DateTime segmentEnd = nextWeek > visualProjectEnd ? visualProjectEnd : nextWeek;

                var bounds = GanttMathEngine.CalculateBounds(segmentStart, segmentEnd, project.ProjectStart, ganttLeft, pointsPerDay);
                if (bounds.Width > 5)
                {
                    int cw = GanttMathEngine.GetIso8601WeekOfYear(currentDay);
                    DrawHeaderLabel(slide, bounds.X, top, bounds.Width, 20, bounds.Width > 20 ? "CW " + cw : cw.ToString(), project.FontFamily, false, 0xF5F5F5, 8);
                }
                currentDay = nextWeek;
            }
        }

        private static void DrawDayLevel(Slide slide, GanttProject project, DateTime visualProjectEnd, double top, double ganttLeft, double pointsPerDay)
        {
            DateTime currentDay = project.ProjectStart.Date;
            while (currentDay < visualProjectEnd.Date)
            {
                double x = ganttLeft + (currentDay - project.ProjectStart.Date).TotalDays * pointsPerDay;
                string label = pointsPerDay > 7 ? currentDay.ToString("dddd", CultureInfo.InvariantCulture).Substring(0, 1).ToUpper() : "";
                DrawHeaderLabel(slide, x, top, pointsPerDay, 20, label, project.FontFamily, false, 0xFFFFFF, 7);
                currentDay = currentDay.AddDays(1);
            }
        }

        private static void DrawHeaderLabel(Slide slide, double x, double y, double width, double height, string text, string fontFamily, bool bold, int bgColor, int fontSize = 9)
        {
            Shape rect = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, (float)x, (float)y, (float)width, (float)height);
            rect.Fill.ForeColor.RGB = bgColor;
            rect.Line.ForeColor.RGB = 0xCCCCCC;
            PowerPointShapeManager.TagShape(rect, Guid.Empty, PowerPointShapeManager.GanttShapeType.Label);

            if (!string.IsNullOrEmpty(text))
            {
                Shape label = slide.Shapes.AddLabel(Office.MsoTextOrientation.msoTextOrientationHorizontal, (float)x, (float)y, (float)width, (float)height);
                label.TextFrame.TextRange.Text = text;
                label.TextFrame.TextRange.Font.Size = fontSize;
                label.TextFrame.TextRange.Font.Name = fontFamily;
                if (bold) label.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                label.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                label.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                PowerPointShapeManager.TagShape(label, Guid.Empty, PowerPointShapeManager.GanttShapeType.Label);
                Marshal.ReleaseComObject(label);
            }
            Marshal.ReleaseComObject(rect);
        }

        private static void DrawTableHeaderCell(Slide slide, double x, double y, double width, string text, string fontFamily, PowerPoint.PpParagraphAlignment align = PowerPoint.PpParagraphAlignment.ppAlignCenter)
        {
            Shape label = slide.Shapes.AddLabel(Office.MsoTextOrientation.msoTextOrientationHorizontal, (float)x, (float)y, (float)width, 20);
            label.TextFrame.TextRange.Text = text;
            label.TextFrame.TextRange.Font.Size = 9;
            label.TextFrame.TextRange.Font.Name = fontFamily;
            label.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            label.TextFrame.TextRange.ParagraphFormat.Alignment = align;
            PowerPointShapeManager.TagShape(label, Guid.Empty, PowerPointShapeManager.GanttShapeType.Label);
            Marshal.ReleaseComObject(label);
        }

        private static void DrawTableArea(Slide slide, GanttProject project, double leftMargin, double topStart, double totalHeight, double tableWidth, double ganttTotalWidth)
        {
            List<float> xs = new List<float>();
            xs.Add((float)leftMargin);

            float currentX = (float)leftMargin;
            if (project.ShowNrColumn)
            {
                currentX += (float)COL_NO_WIDTH;
                xs.Add(currentX);
            }
            if (project.ShowTaskNameColumn)
            {
                currentX += (float)COL_NAME_WIDTH;
                xs.Add(currentX);
            }
            if (project.ShowDurationColumn)
            {
                currentX += (float)COL_DURATION_WIDTH;
                xs.Add(currentX);
            }

            foreach (float x in xs)
            {
                Shape line = slide.Shapes.AddLine(x, (float)topStart, x, (float)(topStart + totalHeight));
                line.Line.ForeColor.RGB = 0xCCCCCC;
                line.Line.Weight = 1.0f;
                PowerPointShapeManager.TagShape(line, Guid.Empty, PowerPointShapeManager.GanttShapeType.Label);
                Marshal.ReleaseComObject(line);
            }
        }

        private static void DrawTableRow(Slide slide, GanttTask task, int index, double top, double rowHeight, double leftMargin, GanttProject project)
        {
            double currentX = leftMargin;

            // No.
            if (project.ShowNrColumn)
            {
                Shape noLabel = slide.Shapes.AddLabel(Office.MsoTextOrientation.msoTextOrientationHorizontal, (float)currentX, (float)top, (float)COL_NO_WIDTH, (float)rowHeight);
                // WBS labels follow the hierarchy (indenting task 5 under 4 renders 4 / 4.1 / 5 with
                // the renumbered successors); a project without WBS keeps the positional number.
                noLabel.TextFrame.TextRange.Text = string.IsNullOrEmpty(task.WBS) ? (index + 1).ToString() : task.WBS;
                noLabel.TextFrame.TextRange.Font.Size = 9;
                noLabel.TextFrame.TextRange.Font.Name = project.FontFamily;
                noLabel.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                noLabel.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                PowerPointShapeManager.TagShape(noLabel, task.Id, PowerPointShapeManager.GanttShapeType.Label);
                Marshal.ReleaseComObject(noLabel);
                currentX += COL_NO_WIDTH;
            }

            // Name
            if (project.ShowTaskNameColumn)
            {
                float indent = task.OutlineLevel * 15;
                float nameWidth = (float)(COL_NAME_WIDTH - indent);
                if (nameWidth < 10) nameWidth = 10;

                Shape nameLabel = slide.Shapes.AddLabel(Office.MsoTextOrientation.msoTextOrientationHorizontal, (float)(currentX + indent), (float)top, nameWidth, (float)rowHeight);
                nameLabel.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;
                nameLabel.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
                nameLabel.TextFrame.MarginLeft = TASK_NAME_TEXT_MARGIN;
                nameLabel.TextFrame.MarginRight = TASK_NAME_TEXT_MARGIN;
                nameLabel.TextFrame.MarginTop = TASK_NAME_TEXT_MARGIN;
                nameLabel.TextFrame.MarginBottom = TASK_NAME_TEXT_MARGIN;
                nameLabel.Width = nameWidth;
                nameLabel.Height = (float)rowHeight;
                nameLabel.TextFrame.TextRange.Text = task.Title;
                var labelFont = nameLabel.TextFrame.TextRange.Font;
                labelFont.Size = task.LabelFontSize ?? 9;
                labelFont.Name = task.LabelFontName ?? project.FontFamily;
                var boldVal = task.LabelBold ?? (task.IsSummary ? true : false);
                labelFont.Bold = boldVal ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                if (task.LabelItalic.HasValue)
                    labelFont.Italic = task.LabelItalic.Value ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                if (task.LabelUnderline.HasValue)
                    labelFont.Underline = task.LabelUnderline.Value ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                if (!string.IsNullOrEmpty(task.CustomTextColor))
                    labelFont.Color.RGB = ParseHtmlColorOrDefault(task.CustomTextColor, 0x000000);
                nameLabel.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                nameLabel.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;
                PowerPointShapeManager.TagShape(nameLabel, task.Id, PowerPointShapeManager.GanttShapeType.RowAnchor);
                Marshal.ReleaseComObject(nameLabel);
                currentX += COL_NAME_WIDTH;
            }

            // Duration
            if (project.ShowDurationColumn)
            {
                Shape durLabel = slide.Shapes.AddLabel(Office.MsoTextOrientation.msoTextOrientationHorizontal, (float)currentX, (float)top, (float)COL_DURATION_WIDTH, (float)rowHeight);
                durLabel.TextFrame.TextRange.Text = task.DurationDays + "d";
                durLabel.TextFrame.TextRange.Font.Size = 9;
                durLabel.TextFrame.TextRange.Font.Name = project.FontFamily;
                durLabel.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                durLabel.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                PowerPointShapeManager.TagShape(durLabel, task.Id, PowerPointShapeManager.GanttShapeType.Label);
                Marshal.ReleaseComObject(durLabel);
            }
        }

        private static void DrawGrid(Slide slide, GanttProject project, DateTime visualProjectEnd, double leftMargin, double topStart, double totalHeight, double pointsPerDay)
        {
            if (project.Tasks.Count == 0) return;

            // 0. Draw Custom Shades first (behind grid lines)
            if (project.CustomShades != null && project.CustomShades.Any())
            {
                foreach (var shade in project.CustomShades)
                {
                    var bounds = GanttMathEngine.CalculateBounds(shade.Start, shade.End, project.ProjectStart, leftMargin, pointsPerDay);
                    if (bounds.Width > 0.1)
                    {
                        Shape rect = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, (float)bounds.X, (float)topStart, (float)bounds.Width, (float)totalHeight);
                        int color = ParseHtmlColorOrDefault(shade.Color, 0xE0E0E0);

                        rect.Fill.ForeColor.RGB = color;
                        rect.Fill.Transparency = shade.Transparency;
                        rect.Line.Visible = Office.MsoTriState.msoFalse;
                        PowerPointShapeManager.TagShape(rect, Guid.Empty, PowerPointShapeManager.GanttShapeType.GridLine);
                        Marshal.ReleaseComObject(rect);
                    }
                }
            }

            DateTime current = project.ProjectStart.Date;

            // Limit grid drawing to prevent COM overhead on massive projects
            int maxShapes = 500;
            int shapeCount = 0;
            int maxIterations = 2000; // Safety against runaway loops
            int iterationCount = 0;

            while (current < visualProjectEnd.Date && iterationCount < maxIterations && shapeCount < maxShapes)
            {
                iterationCount++;
                double x = leftMargin + (current - project.ProjectStart.Date).TotalDays * pointsPerDay;

                // Weekend Highlight - only if dense enough to see and option is enabled
                if (project.ShadeWeekends && pointsPerDay > 5 && !GanttMathEngine.IsWorkDay(current))
                {
                    Shape weekendRect = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, (float)x, (float)topStart, (float)pointsPerDay, (float)totalHeight);
                    weekendRect.Fill.ForeColor.RGB = 0xE0E0E0;
                    weekendRect.Fill.Transparency = 0.8f;
                    weekendRect.Line.Visible = Office.MsoTriState.msoFalse;
                    PowerPointShapeManager.TagShape(weekendRect, Guid.Empty, PowerPointShapeManager.GanttShapeType.GridLine);
                    Marshal.ReleaseComObject(weekendRect);
                    shapeCount++;
                }

                // Vertical Line Logic
                bool drawLine = false;
                if (project.ShowDays && pointsPerDay > 15) drawLine = true; // Every day
                else if (project.ShowWeeks && current.DayOfWeek == DayOfWeek.Monday) drawLine = true; // Every Monday
                else if (project.ShowMonths && current.Day == 1) drawLine = true; // Every 1st of month
                else if (project.ShowYears && current.Day == 1 && current.Month == 1) drawLine = true; // Every 1st of year

                if (drawLine)
                {
                    Shape line = slide.Shapes.AddLine((float)x, (float)topStart, (float)x, (float)(topStart + totalHeight));
                    line.Line.ForeColor.RGB = 0x999999;
                    line.Line.Weight = 0.5f;
                    PowerPointShapeManager.TagShape(line, Guid.Empty, PowerPointShapeManager.GanttShapeType.GridLine);
                    Marshal.ReleaseComObject(line);
                    shapeCount++;
                }

                current = current.AddDays(1);

                // Optimization: skip days in loop for non-daily views
                if (!project.ShowDays)
                {
                    if (project.ShowWeeks && current.DayOfWeek != DayOfWeek.Monday)
                    {
                        // Skip to next monday
                        while (current.DayOfWeek != DayOfWeek.Monday && current < visualProjectEnd) current = current.AddDays(1);
                        current = current.AddDays(-1);
                    }
                    else if (project.ShowMonths && !project.ShowWeeks && current.Day != 1)
                    {
                        current = new DateTime(current.Year, current.Month, 1).AddMonths(1).AddDays(-1);
                    }
                    else if (project.ShowYears && !project.ShowMonths && !project.ShowWeeks && (current.Day != 1 || current.Month != 1))
                    {
                        current = new DateTime(current.Year, 1, 1).AddYears(1).AddDays(-1);
                    }
                }
            }
        }

        private static Shape DrawEventLine(Slide slide, GanttTask task, GanttProject project, double timelineTop, double totalTimelineHeight, double leftMargin, double pointsPerDay)
        {
            double x = leftMargin + (task.Start - project.ProjectStart).TotalDays * pointsPerDay;
            int color = ColorMath.HexToBgr(task.CustomColor ?? "#FF5722");

            // 1. The Line
            Shape line = slide.Shapes.AddLine((float)x, (float)timelineTop, (float)x, (float)(timelineTop + totalTimelineHeight));
            line.Line.ForeColor.RGB = color;
            line.Line.Weight = 1.5f;
            line.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;
            PowerPointShapeManager.TagShape(line, task.Id, PowerPointShapeManager.GanttShapeType.EventLine);

            // 2. The Triangle at the bottom
            float triSize = 10;
            Shape tri = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeIsoscelesTriangle,
                (float)(x - triSize / 2), (float)(timelineTop + totalTimelineHeight), triSize, triSize);

            tri.Fill.ForeColor.RGB = color;
            tri.Line.Visible = Office.MsoTriState.msoFalse;
            tri.Rotation = 0; // Pointing Up
            PowerPointShapeManager.TagShape(tri, task.Id, PowerPointShapeManager.GanttShapeType.EventLine);
            Marshal.ReleaseComObject(tri);
            return line;
        }

        private static Shape DrawTaskBar(Slide slide, GanttTask task, GanttProject project, double top, double rowHeight, double leftMargin, double pointsPerDay)
        {
            var bounds = GanttMathEngine.CalculateBounds(task.Start, task.End, project.ProjectStart, leftMargin, pointsPerDay);
            float y = (float)(top + ((rowHeight - BAR_HEIGHT) / 2));

            var (themeBg, themeFg) = GetThemeColors(slide, project.Theme);

            // Determine Shape Type
            Office.MsoAutoShapeType shapeType = Office.MsoAutoShapeType.msoShapeRectangle;
            switch (task.Shape)
            {
                case TaskShape.RoundedRectangle: shapeType = Office.MsoAutoShapeType.msoShapeRoundedRectangle; break;
                case TaskShape.RightArrow: shapeType = Office.MsoAutoShapeType.msoShapeRightArrow; break;
                case TaskShape.Chevron: shapeType = Office.MsoAutoShapeType.msoShapeChevron; break;
                case TaskShape.Pentagon: shapeType = Office.MsoAutoShapeType.msoShapePentagon; break;
            }

            Shape bar = slide.Shapes.AddShape(shapeType, (float)bounds.X, y, (float)bounds.Width, (float)BAR_HEIGHT);

            // Apply Custom Color if present
            int barColor = themeBg;
            int lineColor = themeFg;

            if (!string.IsNullOrEmpty(task.CustomColor))
            {
                barColor = ParseHtmlColorOrDefault(task.CustomColor, themeBg);
            }

            if (task.IsCritical)
            {
                bar.Fill.ForeColor.RGB = 0xFFCCCC;
                bar.Line.ForeColor.RGB = 0xFF0000;
                bar.Line.Weight = 1.5f;
            }
            else
            {
                bar.Fill.ForeColor.RGB = barColor;
                bar.Line.ForeColor.RGB = lineColor;
                bar.Line.Weight = 0.75f;
            }

            PowerPointShapeManager.TagShape(bar, task.Id, PowerPointShapeManager.GanttShapeType.TaskBar);

            // Re-apply any text the user typed directly into the bar
            if (!string.IsNullOrEmpty(task.CustomBarText))
            {
                try
                {
                    bar.TextFrame.TextRange.Text = task.CustomBarText;
                    bar.TextFrame.TextRange.Font.Size = task.CustomBarTextSize ?? 7;
                    bar.TextFrame.TextRange.Font.Name = project.FontFamily;
                    bar.TextFrame.TextRange.Font.Color.RGB = IsColorDark(barColor) ? 0xFFFFFF : 0x000000;
                    bar.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                }
                catch (COMException)
                {
                    // Custom text is optional; the task bar remains usable without it.
                }
            }

            if (task.PercentComplete > 0)
            {
                double progressWidth = bounds.Width * Math.Min(task.PercentComplete, 1.0);
                if (progressWidth > 0)
                {
                    // For progress, we usually stick to rectangle or matching shape but solid
                    Shape progress = slide.Shapes.AddShape(shapeType, (float)bounds.X, y, (float)progressWidth, (float)BAR_HEIGHT);
                    int progressColor = task.IsCritical ? 0xFF0000 : themeFg;

                    if (!string.IsNullOrEmpty(task.CustomColor))
                    {
                        // Darken the custom color for progress? Or just use it solid.
                        progressColor = barColor;
                    }

                    progress.Fill.ForeColor.RGB = progressColor;
                    progress.Line.Visible = Office.MsoTriState.msoFalse;

                    if (progressWidth > 15 && task.Shape == TaskShape.Rectangle) // Text only looks good in rectangles
                    {
                        progress.TextFrame.TextRange.Text = task.PercentComplete.ToString("P0");
                        progress.TextFrame.TextRange.Font.Size = 7;
                        progress.TextFrame.TextRange.Font.Name = project.FontFamily;
                        progress.TextFrame.TextRange.Font.Color.RGB = IsColorDark(progressColor) ? 0xFFFFFF : 0x000000;
                        progress.TextFrame.MarginLeft = 0;
                        progress.TextFrame.MarginRight = 0;
                        progress.TextFrame.MarginTop = 0;
                        progress.TextFrame.MarginBottom = 0;
                    }

                    PowerPointShapeManager.TagShape(progress, task.Id, PowerPointShapeManager.GanttShapeType.Label);
                    Marshal.ReleaseComObject(progress);
                }
            }

            return bar;
        }

        private static Shape DrawMilestone(Slide slide, GanttTask task, GanttProject project, double top, double rowHeight, double leftMargin, double pointsPerDay)
        {
            var bounds = GanttMathEngine.CalculateBounds(task.Start, task.End, project.ProjectStart, leftMargin, pointsPerDay);
            float size = 10;
            float x = (float)(bounds.X - (size / 2));
            float y = (float)(top + (rowHeight / 2) - (size / 2));

            Office.MsoAutoShapeType shapeType = Office.MsoAutoShapeType.msoShapeDiamond;
            switch (task.MilestoneShape)
            {
                case MilestoneShape.Square: shapeType = Office.MsoAutoShapeType.msoShapeRectangle; break;
                case MilestoneShape.Star: shapeType = Office.MsoAutoShapeType.msoShape5pointStar; break;
                case MilestoneShape.Triangle: shapeType = Office.MsoAutoShapeType.msoShapeIsoscelesTriangle; break;
                case MilestoneShape.Circle: shapeType = Office.MsoAutoShapeType.msoShapeOval; break;
                default: shapeType = Office.MsoAutoShapeType.msoShapeDiamond; break;
            }

            Shape milestoneShape = slide.Shapes.AddShape(shapeType, x, y, size, size);
            int milestoneColor = 0x000000;
            if (!task.IsCritical && !string.IsNullOrEmpty(task.CustomColor))
                milestoneColor = ParseHtmlColorOrDefault(task.CustomColor, 0x000000);
            milestoneShape.Fill.ForeColor.RGB = task.IsCritical ? 0xFF0000 : milestoneColor;
            milestoneShape.Line.Visible = Office.MsoTriState.msoFalse;
            PowerPointShapeManager.TagShape(milestoneShape, task.Id, PowerPointShapeManager.GanttShapeType.Milestone);

            return milestoneShape;
        }

        private static Shape DrawSummaryBar(Slide slide, GanttTask task, GanttProject project, double top, double rowHeight, double leftMargin, double pointsPerDay)
        {
            var bounds = GanttMathEngine.CalculateBounds(task.Start, task.End, project.ProjectStart, leftMargin, pointsPerDay);
            float barThickness = 4;
            float y = (float)(top + ((rowHeight - barThickness) / 2));

            int summaryColor = 0x000000;
            if (!string.IsNullOrEmpty(task.CustomColor))
                summaryColor = ParseHtmlColorOrDefault(task.CustomColor, 0x000000);

            Shape summary = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, (float)bounds.X, y, (float)bounds.Width, barThickness);
            summary.Fill.ForeColor.RGB = summaryColor;
            summary.Line.Visible = Office.MsoTriState.msoFalse;
            PowerPointShapeManager.TagShape(summary, task.Id, PowerPointShapeManager.GanttShapeType.SummaryBar);

            Shape leftTip = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeDownArrow, (float)bounds.X, y, 6, 8);
            leftTip.Fill.ForeColor.RGB = summaryColor;
            leftTip.Line.Visible = Office.MsoTriState.msoFalse;
            PowerPointShapeManager.TagShape(leftTip, task.Id, PowerPointShapeManager.GanttShapeType.Label);
            Marshal.ReleaseComObject(leftTip);

            Shape rightTip = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeDownArrow, (float)(bounds.X + bounds.Width - 6), y, 6, 8);
            rightTip.Fill.ForeColor.RGB = summaryColor;
            rightTip.Line.Visible = Office.MsoTriState.msoFalse;
            PowerPointShapeManager.TagShape(rightTip, task.Id, PowerPointShapeManager.GanttShapeType.Label);
            Marshal.ReleaseComObject(rightTip);

            return summary;
        }

        private static void DrawSubEvents(Slide slide, GanttTask task, GanttProject project, double top, double rowHeight, double leftMargin, double pointsPerDay)
        {
            if (task.SubEvents == null || !task.SubEvents.Any()) return;

            foreach (var subEvent in task.SubEvents)
            {
                float size = 10;
                double daysFromStart = (subEvent.Date - project.ProjectStart.Date).TotalDays;
                float x = (float)(leftMargin + (daysFromStart * pointsPerDay) - (size / 2));
                float y = (float)(top + (rowHeight / 2) - (size / 2));

                Office.MsoAutoShapeType shapeType = Office.MsoAutoShapeType.msoShapeDiamond;
                switch (subEvent.Shape)
                {
                    case MilestoneShape.Square: shapeType = Office.MsoAutoShapeType.msoShapeRectangle; break;
                    case MilestoneShape.Star: shapeType = Office.MsoAutoShapeType.msoShape5pointStar; break;
                    case MilestoneShape.Triangle: shapeType = Office.MsoAutoShapeType.msoShapeIsoscelesTriangle; break;
                    case MilestoneShape.Circle: shapeType = Office.MsoAutoShapeType.msoShapeOval; break;
                    default: shapeType = Office.MsoAutoShapeType.msoShapeDiamond; break;
                }

                Shape subShape = slide.Shapes.AddShape(shapeType, x, y, size, size);

                int color = 0x000000; // Default Black
                if (!string.IsNullOrEmpty(subEvent.Color))
                {
                    color = ParseHtmlColorOrDefault(subEvent.Color, 0x000000);
                }

                subShape.Fill.ForeColor.RGB = color;
                subShape.Line.Visible = Office.MsoTriState.msoFalse;

                // Add a small label if title exists
                if (!string.IsNullOrEmpty(subEvent.Title))
                {
                    Shape label = slide.Shapes.AddLabel(Office.MsoTextOrientation.msoTextOrientationHorizontal, x, y + size, 0, 0);
                    label.TextFrame.TextRange.Text = subEvent.Title;
                    label.TextFrame.TextRange.Font.Size = 7;
                    label.TextFrame.TextRange.Font.Name = project.FontFamily;
                    label.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    PowerPointShapeManager.TagShape(label, task.Id, PowerPointShapeManager.GanttShapeType.Label);
                    Marshal.ReleaseComObject(label);
                }

                PowerPointShapeManager.TagShape(subShape, task.Id, PowerPointShapeManager.GanttShapeType.SubMilestone);
                Marshal.ReleaseComObject(subShape);
            }
        }

        private static void DrawSubBars(Slide slide, GanttTask task, GanttProject project, double top, double rowHeight, double leftMargin, double pointsPerDay)
        {
            if (task.SubBars == null || !task.SubBars.Any()) return;

            foreach (var subBar in task.SubBars)
            {
                var bounds = GanttMathEngine.CalculateBounds(subBar.Start, subBar.End, project.ProjectStart, leftMargin, pointsPerDay);
                float y = (float)(top + ((rowHeight - BAR_HEIGHT) / 2));

                // Determine Shape Type
                Office.MsoAutoShapeType shapeType = Office.MsoAutoShapeType.msoShapeRectangle;
                switch (subBar.Shape)
                {
                    case TaskShape.RoundedRectangle: shapeType = Office.MsoAutoShapeType.msoShapeRoundedRectangle; break;
                    case TaskShape.RightArrow: shapeType = Office.MsoAutoShapeType.msoShapeRightArrow; break;
                    case TaskShape.Chevron: shapeType = Office.MsoAutoShapeType.msoShapeChevron; break;
                    case TaskShape.Pentagon: shapeType = Office.MsoAutoShapeType.msoShapePentagon; break;
                }

                Shape bar = slide.Shapes.AddShape(shapeType, (float)bounds.X, y, (float)bounds.Width, (float)BAR_HEIGHT);

                int color = 0xADD8E6; // Default light blue
                if (!string.IsNullOrEmpty(subBar.Color))
                {
                    color = ParseHtmlColorOrDefault(subBar.Color, 0xADD8E6);
                }

                bar.Fill.ForeColor.RGB = color;
                bar.Line.ForeColor.RGB = 0x000000;
                bar.Line.Weight = 0.5f;

                // Add small label if title exists
                if (!string.IsNullOrEmpty(subBar.Title))
                {
                    bar.TextFrame.TextRange.Text = subBar.Title;
                    bar.TextFrame.TextRange.Font.Size = 7;
                    bar.TextFrame.TextRange.Font.Name = project.FontFamily;
                    bar.TextFrame.TextRange.Font.Color.RGB = IsColorDark(color) ? 0xFFFFFF : 0x000000;
                }

                PowerPointShapeManager.TagShape(bar, task.Id, PowerPointShapeManager.GanttShapeType.SubMilestone); // Reusing tag for simplicity
                Marshal.ReleaseComObject(bar);
            }
        }

        private static void DrawDependencies(Slide slide, GanttProject project, Dictionary<Guid, Shape> taskShapes)
        {
            foreach (var task in project.Tasks)
            {
                if (!taskShapes.ContainsKey(task.Id)) continue;
                var successorShape = taskShapes[task.Id];

                foreach (var predId in task.PredecessorIds)
                {
                    if (!taskShapes.ContainsKey(predId)) continue;
                    var predShape = taskShapes[predId];

                    var link = task.PredecessorLinks.FirstOrDefault(l => l.PredecessorId == predId);
                    var type = link?.Type ?? Models.DependencyType.FinishToStart;

                    // Connection site indices: 1=top, 2=right, 3=bottom, 4=left
                    int beginSite, endSite;
                    int lineColor;
                    switch (type)
                    {
                        case Models.DependencyType.StartToStart:
                            beginSite = 4; endSite = 4;  // left → left
                            lineColor = ColorMath.HexToBgr("#2980B9"); // blue
                            break;
                        case Models.DependencyType.FinishToFinish:
                            beginSite = 2; endSite = 2;  // right → right
                            lineColor = ColorMath.HexToBgr("#27AE60"); // green
                            break;
                        default: // FinishToStart
                            beginSite = 2; endSite = 4;  // right → left
                            lineColor = 0x808080; // grey
                            break;
                    }

                    try
                    {
                        Shape connector = slide.Shapes.AddConnector(Office.MsoConnectorType.msoConnectorElbow, 0, 0, 100, 100);
                        connector.Line.ForeColor.RGB = lineColor;
                        connector.Line.Weight = 1.0f;
                        connector.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadTriangle;
                        connector.ConnectorFormat.BeginConnect(predShape, beginSite);
                        connector.ConnectorFormat.EndConnect(successorShape, endSite);
                        PowerPointShapeManager.TagShape(connector, task.Id, PowerPointShapeManager.GanttShapeType.DependencyArrow);
                        Marshal.ReleaseComObject(connector);
                    }
                    catch (COMException)
                    {
                        // PowerPoint can reject a connector site for a shape type; skip that arrow.
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, $"GanttRenderer.DrawDependencies.{predId}->{task.Id}");
                    }
                }
            }
        }

        private static void DrawLegend(Slide slide, GanttProject project, double top, double left, string fontFamily)
        {
            float currentX = (float)left;
            float y = (float)top;
            float boxSize = 10;
            float padding = 5;
            float itemSpacing = 100;

            var (themeBg, themeFg) = GetThemeColors(slide, project.Theme);

            Shape tBox = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, currentX, y, boxSize, boxSize);
            tBox.Fill.ForeColor.RGB = themeBg;
            tBox.Line.ForeColor.RGB = themeFg;
            PowerPointShapeManager.TagShape(tBox, Guid.Empty, PowerPointShapeManager.GanttShapeType.Label);
            DrawLegendLabel(slide, currentX + boxSize + padding, y, "Task", fontFamily);
            currentX += itemSpacing;

            Shape cBox = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, currentX, y, boxSize, boxSize);
            cBox.Fill.ForeColor.RGB = 0xFF0000;
            PowerPointShapeManager.TagShape(cBox, Guid.Empty, PowerPointShapeManager.GanttShapeType.Label);
            DrawLegendLabel(slide, currentX + boxSize + padding, y, "Critical Path", fontFamily);
            currentX += itemSpacing;

            Shape mDia = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeDiamond, currentX, y, boxSize, boxSize);
            mDia.Fill.ForeColor.RGB = 0x000000;
            PowerPointShapeManager.TagShape(mDia, Guid.Empty, PowerPointShapeManager.GanttShapeType.Label);
            DrawLegendLabel(slide, currentX + boxSize + padding, y, "Milestone", fontFamily);
            currentX += itemSpacing;

            Shape sLine = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, currentX, y + 3, boxSize, 3);
            sLine.Fill.ForeColor.RGB = 0x000000;
            PowerPointShapeManager.TagShape(sLine, Guid.Empty, PowerPointShapeManager.GanttShapeType.Label);
            DrawLegendLabel(slide, currentX + boxSize + padding, y, "Summary Task", fontFamily);

            Marshal.ReleaseComObject(tBox);
            Marshal.ReleaseComObject(cBox);
            Marshal.ReleaseComObject(mDia);
            Marshal.ReleaseComObject(sLine);
        }

        private static void DrawLegendLabel(Slide slide, float x, float y, string text, string fontFamily)
        {
            Shape label = slide.Shapes.AddLabel(Office.MsoTextOrientation.msoTextOrientationHorizontal, x, y - 2, 0, 0);
            label.TextFrame.TextRange.Text = text;
            label.TextFrame.TextRange.Font.Size = 8;
            label.TextFrame.TextRange.Font.Name = fontFamily;
            label.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            PowerPointShapeManager.TagShape(label, Guid.Empty, PowerPointShapeManager.GanttShapeType.Label);
            Marshal.ReleaseComObject(label);
        }

        private static HashSet<string> CaptureTopLevelShapeNames(Slide slide)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Shapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    Shape shape = null;
                    try
                    {
                        shape = shapes[i];
                        names.Add(shape.Name);
                    }
                    finally
                    {
                        if (shape != null) Marshal.ReleaseComObject(shape);
                    }
                }
            }
            finally
            {
                if (shapes != null) Marshal.ReleaseComObject(shapes);
            }

            return names;
        }

        private static Shape GroupGanttShapes(Slide slide, Guid chartId, HashSet<string> preExistingShapeNames, double pointsPerDay, DateTime projectStart)
        {
            var ganttShapeNames = new List<string>();
            Shapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    Shape shape = null;
                    try
                    {
                        shape = shapes[i];
                        if (!preExistingShapeNames.Contains(shape.Name) && PowerPointShapeManager.IsGanttShape(shape))
                        {
                            PowerPointShapeManager.SetChartId(shape, chartId);
                            ganttShapeNames.Add(shape.Name);
                        }
                    }
                    finally
                    {
                        if (shape != null) Marshal.ReleaseComObject(shape);
                    }
                }
            }
            finally
            {
                if (shapes != null) Marshal.ReleaseComObject(shapes);
            }

            if (ganttShapeNames.Count > 1)
            {
                try
                {
                    ShapeRange range = slide.Shapes.Range(ganttShapeNames.ToArray());
                    Shape group = range.Group();
                    group.Name = GANTT_GROUP_NAME + "_" + chartId.ToString("N");
                    group.Tags.Add(PowerPointShapeManager.GANTT_PROJECT_TAG, "TRUE");
                    PowerPointShapeManager.SetChartId(group, chartId, true);
                    PowerPointShapeManager.SetGanttScale(group, pointsPerDay);
                    PowerPointShapeManager.SetGanttProjectStart(group, projectStart);
                    Marshal.ReleaseComObject(range);
                    return group;
                }
                catch (COMException)
                {
                    // Grouping may fail if PowerPoint removed a rendered shape.
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "GanttRenderer.GroupGanttShapes");
                }
            }
            return null;
        }

        private static (int bg, int fg) GetThemeColors(Slide slide, GanttTheme theme)
        {
            switch (theme)
            {
                case GanttTheme.HighContrast: return (0x000000, 0x00FF00);
                case GanttTheme.Pastel: return (0xFFDAB9, 0xFF69B4);
                default:
                    try
                    {
                        if (slide != null)
                        {
                            Office.ThemeColorScheme themeColors = null;
                            object accent1Color = null;
                            object accent2Color = null;
                            try
                            {
                                themeColors = slide.ThemeColorScheme;
                                accent1Color = themeColors.Colors(Office.MsoThemeColorSchemeIndex.msoThemeAccent1);
                                accent2Color = themeColors.Colors(Office.MsoThemeColorSchemeIndex.msoThemeAccent2);
                                return (((dynamic)accent1Color).RGB, ((dynamic)accent2Color).RGB);
                            }
                            finally
                            {
                                if (accent2Color != null && Marshal.IsComObject(accent2Color)) Marshal.ReleaseComObject(accent2Color);
                                if (accent1Color != null && Marshal.IsComObject(accent1Color)) Marshal.ReleaseComObject(accent1Color);
                                if (themeColors != null) Marshal.ReleaseComObject(themeColors);
                            }
                        }
                    }
                    catch (COMException)
                    {
                        // Theme colors are optional; use the documented fallback.
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "GanttRenderer.GetThemeColors");
                    }
                    return (0xADD8E6, 0x4682B4); // Fallback
            }
        }

        private static int ParseHtmlColorOrDefault(string htmlColor, int fallback)
        {
            if (string.IsNullOrWhiteSpace(htmlColor)) return fallback;

            try
            {
                return System.Drawing.ColorTranslator.ToOle(
                    System.Drawing.ColorTranslator.FromHtml(htmlColor));
            }
            catch (ArgumentException)
            {
                // Persisted user input can contain a stale or invalid color; retain the caller's fallback.
                return fallback;
            }
        }

        private static bool IsColorDark(int rgb)
        {
            int r = rgb & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = (rgb >> 16) & 0xFF;
            double luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
            return luminance < 0.5;
        }
    }
}
