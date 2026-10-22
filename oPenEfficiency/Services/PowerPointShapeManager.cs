using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Manages the creation, identification, and cleanup of PowerPoint shapes for the Gantt chart.
    /// Handles COM exceptions and prevents memory leaks through explicit reference management.
    /// </summary>
    public static class PowerPointShapeManager
    {
        public const string GANTT_ID_TAG = "oPE_GanttId";
        public const string GANTT_TYPE_TAG = "oPE_GanttType";
        public const string GANTT_PROJECT_TAG = "oPE_GanttProject";
        public const string GANTT_CHART_ID_TAG = "oPE_GanttChartId";
        public const string GANTT_SCALE_TAG = "oPE_GanttPointsPerDay";
        public const string GANTT_PROJECT_START_TAG = "oPE_GanttProjectStart";

        private const string LegacyGanttIdTag = "GANTT_ID";
        private const string LegacyGanttTypeTag = "GANTT_TYPE";
        private const string LegacyGanttProjectTag = "GANTT_PROJECT";

        public enum GanttShapeType
        {
            TaskBar,
            SummaryBar,
            Milestone,
            DependencyArrow,
            GridLine,
            Label,
            RowAnchor,
            EventLine,
            SubMilestone
        }

        /// <summary>
        /// Finds the top-level group belonging to a specific chart instance.
        /// </summary>
        public static Shape FindChartGroup(Slide slide, Guid chartId)
        {
            if (slide == null || chartId == Guid.Empty) return null;

            Shapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                string chartIdText = chartId.ToString();
                for (int i = 1; i <= shapes.Count; i++)
                {
                    Shape shape = null;
                    try
                    {
                        shape = shapes[i];
                        if (GetChartId(shape) == chartIdText && HasGanttProjectTag(shape))
                        {
                            return shape;
                        }
                    }
                    catch (COMException)
                    {
                        // A stale shape is skipped while probing the chart group.
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, $"PowerPointShapeManager.FindChartGroup.Shape[{i}]");
                    }

                    if (shape != null) Marshal.ReleaseComObject(shape);
                }
            }
            finally
            {
                if (shapes != null) Marshal.ReleaseComObject(shapes);
            }

            return null;
        }

        /// <summary>
        /// Clears only the shapes belonging to one chart instance. Other Gantt/timeline charts
        /// on the same slide are deliberately left untouched.
        /// </summary>
        public static void ClearGanttShapes(Slide slide, Guid chartId)
        {
            if (slide == null || chartId == Guid.Empty) return;

            Shapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                string chartIdText = chartId.ToString();

                // Delete one matching root group first. Copy/paste can temporarily create duplicate
                // identities; deleting only the selected/first root prevents collateral chart loss.
                for (int i = 1; i <= shapes.Count; i++)
                {
                    Shape candidate = null;
                    try
                    {
                        candidate = shapes[i];
                        if (GetChartId(candidate) == chartIdText && HasGanttProjectTag(candidate))
                        {
                            candidate.Delete();
                            break;
                        }
                    }
                    catch (COMException)
                    {
                        // PowerPoint may remove a candidate while the collection is traversed.
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, $"PowerPointShapeManager.ClearGanttShapes.Shape[{i}]");
                    }
                    finally
                    {
                        if (candidate != null) Marshal.ReleaseComObject(candidate);
                    }
                }

                // Clean up ungrouped remnants of this chart only.
                int shapeCount = shapes.Count;

                // Iterate backwards when deleting
                for (int i = shapeCount; i >= 1; i--)
                {
                    Shape shape = null;
                    try
                    {
                        shape = shapes[i];

                        if (!HasGanttProjectTag(shape) && GetChartId(shape) == chartIdText)
                        {
                            shape.Delete();
                        }
                    }
                    catch (COMException)
                    {
                        // Shape might have been deleted or is inaccessible
                    }
                    finally
                    {
                        if (shape != null) Marshal.ReleaseComObject(shape);
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PowerPointShapeManager.ClearGanttShapes");
            }
            finally
            {
                if (shapes != null) Marshal.ReleaseComObject(shapes);
            }
        }

        /// <summary>
        /// Finds a specific shape by its Gantt type and Task ID.
        /// </summary>
        public static Shape GetShapeByTypeAndId(Slide slide, GanttShapeType type, Guid id)
        {
            return GetShapeByTypeAndId(slide, type, id, Guid.Empty);
        }

        public static Shape GetShapeByTypeAndId(Slide slide, GanttShapeType type, Guid id, Guid chartId)
        {
            if (slide == null) return null;
            Shapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                string idString = id.ToString();
                string typeString = type.ToString();

                for (int shapeIndex = 1; shapeIndex <= shapes.Count; shapeIndex++)
                {
                    Shape shape = null;
                    bool returnShape = false;
                    try
                    {
                        shape = shapes[shapeIndex];
                        if (MatchesChart(shape, chartId) && GetGanttId(shape) == idString && GetGanttType(shape) == typeString)
                        {
                            returnShape = true;
                            return shape;
                        }

                        if (shape.Type == Office.MsoShapeType.msoGroup)
                        {
                            for (int index = 1; index <= shape.GroupItems.Count; index++)
                            {
                                Shape child = null;
                                bool returnChild = false;
                                try
                                {
                                    child = shape.GroupItems[index];
                                    if (MatchesChart(shape, chartId) && GetGanttId(child) == idString && GetGanttType(child) == typeString)
                                    {
                                        returnChild = true;
                                        return child;
                                    }
                                }
                                finally
                                {
                                    if (child != null && !returnChild) Marshal.ReleaseComObject(child);
                                }
                            }
                        }
                    }
                    catch (COMException)
                    {
                        // A stale shape is skipped while resolving a task element.
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, $"PowerPointShapeManager.GetShapeByTypeAndId.Shape[{shapeIndex}]");
                    }
                    finally
                    {
                        if (shape != null && !returnShape) Marshal.ReleaseComObject(shape);
                    }
                }
            }
            finally
            {
                if (shapes != null) Marshal.ReleaseComObject(shapes);
            }
            return null;
        }

        /// <summary>
        /// Tags a shape with a Gantt ID and Type for future identification.
        /// </summary>
        public static void TagShape(Shape shape, Guid id, GanttShapeType type)
        {
            if (shape == null) return;

            try
            {
                shape.Tags.Add(GANTT_ID_TAG, id.ToString());
                shape.Tags.Add(GANTT_TYPE_TAG, type.ToString());
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PowerPointShapeManager.TagShape");
            }
        }

        /// <summary>
        /// Checks if a shape belongs to any Gantt chart.
        /// </summary>
        public static bool IsGanttShape(Shape shape)
        {
            if (shape == null) return false;
            try
            {
                return HasGanttProjectTag(shape) ||
                       !string.IsNullOrEmpty(GetChartId(shape)) ||
                       !string.IsNullOrEmpty(GetGanttId(shape));
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Finds all shapes on a slide associated with a specific Task ID.
        /// </summary>
        public static List<Shape> GetShapesByTaskId(Slide slide, Guid id)
        {
            var result = new List<Shape>();
            if (slide == null) return result;

            Shapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                string idString = id.ToString();

                for (int i = 1; i <= shapes.Count; i++)
                {
                    Shape shape = null;
                    bool keepShape = false;
                    try
                    {
                        shape = shapes[i];
                        if (GetGanttId(shape) == idString)
                        {
                            result.Add(shape);
                            keepShape = true;
                        }
                    }
                    catch (COMException)
                    {
                        // Skip an individual stale shape and continue collecting valid task shapes.
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, $"PowerPointShapeManager.GetShapesByTaskId.Shape[{i}]");
                    }
                    finally
                    {
                        if (shape != null && !keepShape) Marshal.ReleaseComObject(shape);
                    }
                }
            }
            finally
            {
                if (shapes != null) Marshal.ReleaseComObject(shapes);
            }

            return result;
        }

        /// <summary>
        /// Adjusts the Z-order of Gantt shapes to ensure correct layering.
        /// Gridlines -> Background
        /// Arrows -> Foreground
        /// </summary>
        public static void ReorderZLayers(Slide slide)
        {
            if (slide == null) return;

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
                        string typeStr = GetGanttType(shape);
                        if (string.IsNullOrEmpty(typeStr)) continue;

                        if (Enum.TryParse(typeStr, out GanttShapeType type))
                        {
                            switch (type)
                            {
                                case GanttShapeType.GridLine:
                                    shape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                                    break;
                                case GanttShapeType.DependencyArrow:
                                    shape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                                    break;
                                // Task bars and summaries naturally sit in the middle
                            }
                        }
                    }
                    catch (COMException)
                    {
                        // A deleted shape cannot be reordered and is skipped.
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, $"PowerPointShapeManager.ReorderZLayers.Shape[{i}]");
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
        }

        public static string GetGanttId(Shape shape)
        {
            return GetTagValue(shape, GANTT_ID_TAG, LegacyGanttIdTag);
        }

        public static string GetGanttType(Shape shape)
        {
            return GetTagValue(shape, GANTT_TYPE_TAG, LegacyGanttTypeTag);
        }

        /// <summary>
        /// Records the points-per-day scale a chart was actually rendered with, so slide
        /// synchronization can interpret bar geometry in the chart's own scale (REG-GANTT-RESIZE-ADOPT-001).
        /// </summary>
        public static void SetGanttScale(Shape shape, double pointsPerDay)
        {
            if (shape == null) return;

            try
            {
                shape.Tags.Add(GANTT_SCALE_TAG,
                    pointsPerDay.ToString("R", System.Globalization.CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PowerPointShapeManager.SetGanttScale");
            }
        }

        /// <summary>
        /// Returns the recorded render scale, or 0 when the chart carries none (legacy charts).
        /// </summary>
        public static double GetGanttScale(Shape shape)
        {
            string value = GetTagValue(shape, GANTT_SCALE_TAG, null);
            return double.TryParse(value, System.Globalization.NumberStyles.Float,
                       System.Globalization.CultureInfo.InvariantCulture, out double scale) && scale > 0
                ? scale
                : 0;
        }

        /// <summary>
        /// Records the timeline day-zero the chart was rendered against. ProjectStart is a live
        /// boundary (the earliest task start), so slide synchronization must map bar geometry with
        /// the render-time origin, not the drifting one (REG-GANTT-RESIZE-ADOPT-001).
        /// </summary>
        public static void SetGanttProjectStart(Shape shape, DateTime projectStart)
        {
            if (shape == null) return;

            try
            {
                shape.Tags.Add(GANTT_PROJECT_START_TAG,
                    projectStart.Date.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PowerPointShapeManager.SetGanttProjectStart");
            }
        }

        /// <summary>
        /// Returns the recorded render-time project start, or default when the chart carries none.
        /// </summary>
        public static DateTime GetGanttProjectStart(Shape shape)
        {
            string value = GetTagValue(shape, GANTT_PROJECT_START_TAG, null);
            return DateTime.TryParse(value, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out DateTime parsed)
                ? parsed.Date
                : default;
        }

        public static string GetChartId(Shape shape)
        {
            return GetTagValue(shape, GANTT_CHART_ID_TAG, null);
        }

        public static void SetChartId(Shape shape, Guid chartId, bool includeGroupItems = false)
        {
            if (shape == null || chartId == Guid.Empty) return;

            try
            {
                shape.Tags.Add(GANTT_CHART_ID_TAG, chartId.ToString());

                if (includeGroupItems && shape.Type == Office.MsoShapeType.msoGroup)
                {
                    int count = shape.GroupItems.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        Shape child = null;
                        try
                        {
                            child = shape.GroupItems[i];
                            child.Tags.Add(GANTT_CHART_ID_TAG, chartId.ToString());
                        }
                        finally
                        {
                            if (child != null) Marshal.ReleaseComObject(child);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PowerPointShapeManager.SetChartId");
            }
        }

        public static bool HasGanttProjectTag(Shape shape)
        {
            return !string.IsNullOrEmpty(GetTagValue(shape, GANTT_PROJECT_TAG, LegacyGanttProjectTag));
        }

        private static string GetTagValue(Shape shape, string currentName, string legacyName)
        {
            if (shape == null) return string.Empty;

            try
            {
                string value = shape.Tags[currentName];
                if (!string.IsNullOrEmpty(value)) return value;

                return string.IsNullOrEmpty(legacyName) ? string.Empty : shape.Tags[legacyName] ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool MatchesChart(Shape shape, Guid chartId)
        {
            return chartId == Guid.Empty || GetChartId(shape) == chartId.ToString();
        }
    }
}
