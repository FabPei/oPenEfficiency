using System;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Host-free date-adoption math for Gantt slide synchronization (REG-GANTT-RESIZE-ADOPT-001).
    /// SyncFromSlide derives task dates from the rendered bar geometry; this policy is the single
    /// place where shape coordinates become dates so the behavior is deterministically testable.
    /// </summary>
    public static class GanttSyncAdoptionPolicy
    {
        /// <summary>
        /// Shape geometry is float32-quantized: a bar dragged to exactly 1.5 days reads as
        /// 1.49967 days. Values within this tolerance of a half-day boundary round away from
        /// zero, so a visual "half day more" reliably becomes one day more (REG-GANTT-RESIZE-ADOPT-001).
        /// </summary>
        private const double HalfDayRoundingTolerance = 0.005;

        /// <summary>
        /// Rounds a day count to whole days, tolerant against float32 shape-quantization ripples
        /// at the half-day boundary (rounds away from zero there).
        /// </summary>
        public static double RoundDays(double days)
        {
            double floor = Math.Floor(days);
            double fraction = days - floor;
            return fraction >= 0.5 - HalfDayRoundingTolerance ? floor + 1 : floor;
        }
        /// <summary>
        /// The scale recorded at render time wins over any recomputation: only the chart itself
        /// knows which points-per-day scale it was actually drawn with (the render caller may pass
        /// any total width), and a resized bar must be interpreted in that scale.
        /// Returns the recomputed fallback when no usable scale was recorded (legacy charts).
        /// </summary>
        public static double ResolvePointsPerDay(double recordedPointsPerDay, double recomputedPointsPerDay)
        {
            return recordedPointsPerDay > 0 ? recordedPointsPerDay : recomputedPointsPerDay;
        }

        /// <summary>
        /// Derives a task's start/end from its bar geometry exactly as the pre-existing inline
        /// SyncFromSlide logic did: start from the left edge (milestones from the center), duration
        /// from the width, both rounded to whole days, optionally snapped to work days in
        /// rule-based mode.
        /// </summary>
        public static void AdoptDates(
            double shapeLeft,
            double shapeWidth,
            double ganttLeftMargin,
            double pointsPerDay,
            DateTime projectStart,
            bool followGanttRules,
            bool isMilestone,
            out DateTime newStart,
            out DateTime newEnd)
        {
            double left = shapeLeft;
            double width = shapeWidth;
            if (isMilestone)
            {
                left += width / 2;
                width = 0;
            }

            double daysFromStart = (left - ganttLeftMargin) / pointsPerDay;
            newStart = projectStart.AddDays(RoundDays(daysFromStart)).Date;

            double durationDays = width / pointsPerDay;
            newEnd = newStart.AddDays(RoundDays(durationDays)).Date;

            if (followGanttRules)
            {
                newStart = GanttMathEngine.NextWorkDay(newStart);
                if (durationDays == 0)
                {
                    newEnd = newStart;
                }
                else
                {
                    newEnd = GanttMathEngine.NextWorkDay(newEnd);
                    if (newEnd <= newStart) newEnd = GanttMathEngine.AddWorkDays(newStart, 1);
                }
            }
        }
    }
}
