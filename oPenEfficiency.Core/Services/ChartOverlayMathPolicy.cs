using System;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free mathematical contract for Add Analytical Chart Overlays (QA-S9 wave 5).
    /// This is the production implementation used by ChartOverlayDialog for stacked point totals,
    /// Delta, CAGR, Growth, Average Level and the historical 1.2-population-sigma outlier rule.
    /// </summary>
    public static class ChartOverlayMathPolicy
    {
        public const double OutlierStdDevMultiplier = 1.2d;

        /// <summary>
        /// Sums every series at each point index. This intentionally models the historical
        /// ChartOverlay behavior: multi-series charts are always analyzed as stacked totals,
        /// irrespective of the native chart grouping.
        /// </summary>
        public static double[] AggregateSeriesTotals(double[][] seriesValues, int pointCount)
        {
            var totals = new double[pointCount];
            if (seriesValues == null) return totals;

            for (int seriesIndex = 0; seriesIndex < seriesValues.Length; seriesIndex++)
            {
                double[] values = seriesValues[seriesIndex];
                if (values == null) continue;
                for (int pointIndex = 0; pointIndex < pointCount; pointIndex++)
                    totals[pointIndex] += values[pointIndex];
            }

            return totals;
        }

        public static double CalculateDelta(double startValue, double endValue)
        {
            return endValue - startValue;
        }

        /// <summary>
        /// Historical CAGR semantics: period count is point-index distance, not a date/category
        /// duration. A zero period count is coerced to one, matching the former dialog code.
        /// Non-positive starts intentionally retain the historical NaN/Infinity behavior.
        /// </summary>
        public static double CalculateCagr(double startValue, double endValue, int periods)
        {
            if (periods == 0) periods = 1;
            return (Math.Pow(endValue / startValue, 1.0 / periods) - 1.0) * 100.0;
        }

        public static double CalculateGrowth(double startValue, double endValue)
        {
            return ((endValue - startValue) / Math.Abs(startValue)) * 100.0;
        }

        public static double CalculateMean(double[] values)
        {
            double sum = 0;
            for (int i = 0; i < values.Length; i++) sum += values[i];
            return sum / values.Length;
        }

        public static double CalculatePopulationStdDev(double[] values, double mean)
        {
            double sumSq = 0;
            for (int i = 0; i < values.Length; i++)
                sumSq += Math.Pow(values[i] - mean, 2);
            return Math.Sqrt(sumSq / values.Length);
        }

        public static bool IsHistoricalOutlier(double value, double mean, double stdDev)
        {
            double thresholdUpper = mean + OutlierStdDevMultiplier * stdDev;
            double thresholdLower = mean - OutlierStdDevMultiplier * stdDev;
            return value > thresholdUpper || value < thresholdLower;
        }
    }
}
