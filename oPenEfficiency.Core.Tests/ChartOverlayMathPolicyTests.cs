using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave-5 coverage for Add Analytical Chart Overlays (QA9-BtnChartOverlay-002):
    /// the exact host-free math frozen by CH-S0 and now used by ChartOverlayDialog.
    /// </summary>
    [TestClass]
    public sealed class ChartOverlayMathPolicyTests
    {
        [TestMethod]
        public void AggregateSeriesTotalsAlwaysUsesCrossSeriesSums()
        {
            double[][] values =
            {
                new[] { 10d, 20d, 30d },
                new[] { 1d, 2d, 3d },
                new[] { -2d, 4d, 6d }
            };

            CollectionAssert.AreEqual(
                new[] { 9d, 26d, 39d },
                ChartOverlayMathPolicy.AggregateSeriesTotals(values, 3));
        }

        [TestMethod]
        public void DeltaAndGrowthMatchTheFrozenFormulas()
        {
            Assert.AreEqual(30d, ChartOverlayMathPolicy.CalculateDelta(120d, 150d), 0.000001);
            Assert.AreEqual(-30d, ChartOverlayMathPolicy.CalculateDelta(150d, 120d), 0.000001);
            Assert.AreEqual(25d, ChartOverlayMathPolicy.CalculateGrowth(120d, 150d), 0.000001);
            Assert.AreEqual(-20d, ChartOverlayMathPolicy.CalculateGrowth(-100d, -120d), 0.000001);
        }

        [TestMethod]
        public void CagrUsesPointIndexDistanceRatherThanCalendarTime()
        {
            Assert.AreEqual(10d, ChartOverlayMathPolicy.CalculateCagr(100d, 121d, 2), 0.000001);
            Assert.AreEqual(21d, ChartOverlayMathPolicy.CalculateCagr(100d, 121d, 0), 0.000001);
        }

        [TestMethod]
        public void NonPositiveStartKeepsTheHistoricalNonFiniteBehavior()
        {
            Assert.IsTrue(double.IsInfinity(ChartOverlayMathPolicy.CalculateGrowth(0d, 10d)));
            Assert.IsTrue(double.IsInfinity(ChartOverlayMathPolicy.CalculateCagr(0d, 10d, 1)));
            Assert.IsTrue(double.IsNaN(ChartOverlayMathPolicy.CalculateCagr(-100d, 121d, 2)));
        }

        [TestMethod]
        public void MeanAndPopulationStdDevUsePopulationDenominator()
        {
            double[] values = { 1d, 2d, 3d, 4d };
            double mean = ChartOverlayMathPolicy.CalculateMean(values);
            double stdDev = ChartOverlayMathPolicy.CalculatePopulationStdDev(values, mean);

            Assert.AreEqual(2.5d, mean, 0.000001);
            Assert.AreEqual(Math.Sqrt(1.25d), stdDev, 0.000001);
        }

        [TestMethod]
        public void OutlierRulePinsTheHistoricalOnePointTwoSigmaMultiplier()
        {
            Assert.AreEqual(1.2d, ChartOverlayMathPolicy.OutlierStdDevMultiplier, 0.000001);
            Assert.IsFalse(ChartOverlayMathPolicy.IsHistoricalOutlier(11.9d, 10d, 2d));
            Assert.IsTrue(ChartOverlayMathPolicy.IsHistoricalOutlier(12.5d, 10d, 2d));
            Assert.IsFalse(ChartOverlayMathPolicy.IsHistoricalOutlier(7.7d, 10d, 2d));
            Assert.IsTrue(ChartOverlayMathPolicy.IsHistoricalOutlier(7.5d, 10d, 2d));
        }
    }
}
