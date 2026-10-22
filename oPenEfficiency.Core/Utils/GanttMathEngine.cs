using System;

namespace oPenEfficiency.Utils
{
    /// <summary>
    /// The mathematical engine for the Gantt chart, responsible for time-to-coordinate conversions and calendar logic.
    /// Operates strictly in PowerPoint 'Points' (72 DPI).
    /// </summary>
    public static class GanttMathEngine
    {
        /// <summary>
        /// Calculates the horizontal bounds (X-position and Width) of a task on the slide.
        /// </summary>
        /// <param name="taskStart">The start date of the task.</param>
        /// <param name="taskEnd">The end date of the task.</param>
        /// <param name="projectStart">The start date of the entire project timeline.</param>
        /// <param name="projectEnd">The end date of the entire project timeline.</param>
        /// <param name="leftMargin">The X-coordinate where the Gantt area starts.</param>
        /// <param name="totalAvailableWidth">The total width available for the Gantt timeline in Points.</param>
        /// <returns>A tuple containing (X, Width) in Points.</returns>
        public static (double X, double Width) CalculateBounds(
            DateTime taskStart,
            DateTime taskEnd,
            DateTime projectStart,
            double leftMargin,
            double pointsPerDay)
        {
            double daysFromStart = (taskStart - projectStart).TotalDays;
            double xOffset = daysFromStart * pointsPerDay;
            double x = leftMargin + xOffset;

            double taskDuration = (taskEnd - taskStart).TotalDays;
            double width = taskDuration * pointsPerDay;

            if (width < 0) width = 0;

            return (x, width);
        }

        /// <summary>
        /// Determines if a specific date is a working day (Monday through Friday).
        /// </summary>
        public static bool IsWorkDay(DateTime date)
        {
            return date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday;
        }

        /// <summary>
        /// Returns the given date if it is a work day, otherwise advances to the next Monday.
        /// </summary>
        public static DateTime NextWorkDay(DateTime date)
        {
            while (!IsWorkDay(date))
                date = date.AddDays(1);
            return date;
        }

        public static DateTime PreviousWorkDay(DateTime date)
        {
            while (!IsWorkDay(date))
                date = date.AddDays(-1);
            return date;
        }

        /// <summary>
        /// Moves a date by a number of working days. Zero keeps the original boundary intact.
        /// </summary>
        public static DateTime AddWorkDays(DateTime date, int workDays)
        {
            if (workDays == 0) return date;

            int direction = workDays > 0 ? 1 : -1;
            int remaining = Math.Abs(workDays);
            DateTime result = date;
            while (remaining > 0)
            {
                result = result.AddDays(direction);
                if (IsWorkDay(result)) remaining--;
            }
            return result;
        }

        /// <summary>
        /// Counts working-day slots in the half-open range [start, end).
        /// A Friday-to-Monday task therefore has a duration of one working day.
        /// </summary>
        public static int CountWorkDays(DateTime start, DateTime end)
        {
            if (start == end) return 0;
            if (end < start) return -CountWorkDays(end, start);

            int count = 0;
            for (DateTime day = start.Date; day < end.Date; day = day.AddDays(1))
            {
                if (IsWorkDay(day)) count++;
            }
            return count;
        }

        /// <summary>
        /// Calculates the width of a single day in Points.
        /// Useful for drawing gridlines or snapping.
        /// </summary>
        public static double GetDayWidth(DateTime projectStart, DateTime projectEnd, double totalAvailableWidth)
        {
            double totalDays = (projectEnd - projectStart).TotalDays;
            if (totalDays <= 0) return 0;
            return totalAvailableWidth / totalDays;
        }

        /// <summary>
        /// Converts a X-coordinate back into a DateTime within the project context.
        /// Useful for drag-and-drop operations.
        /// </summary>
        public static DateTime CoordinateToDate(double x, double leftMargin, double totalAvailableWidth, DateTime projectStart, DateTime projectEnd)
        {
            if (totalAvailableWidth <= 0) return projectStart;

            double relativeX = x - leftMargin;
            double percentage = relativeX / totalAvailableWidth;

            double totalSeconds = (projectEnd - projectStart).TotalSeconds;
            double secondsOffset = totalSeconds * percentage;

            return projectStart.AddSeconds(secondsOffset);
        }

        /// <summary>
        /// Helper to convert WPF Pixels (96 DPI) to PowerPoint Points (72 DPI).
        /// PowerPoint Point = Pixel * (72 / 96) = Pixel * 0.75
        /// </summary>
        public static double PixelsToPoints(double pixels)
        {
            return pixels * 0.75;
        }

        /// <summary>
        /// Helper to convert PowerPoint Points (72 DPI) to WPF Pixels (96 DPI).
        /// WPF Pixel = Point * (96 / 72) = Point * 1.333...
        /// </summary>
        public static double PointsToPixels(double points)
        {
            return points * (96.0 / 72.0);
        }

        /// <summary>
        /// Returns the ISO 8601 week number for a given date.
        /// </summary>
        public static int GetIso8601WeekOfYear(DateTime time)
        {
            DayOfWeek day = System.Globalization.CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
            {
                time = time.AddDays(3);
            }

            return System.Globalization.CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }
    }
}
