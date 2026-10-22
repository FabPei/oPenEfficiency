using System;
using System.Drawing;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Excel.Features.ChartScanner;
using oPenEfficiency.Excel.Features.Data;
using oPenEfficiency.Excel.Features.Rounding;
using oPenEfficiency.Excel.Services;

namespace oPenEfficiency.Excel.Tests
{
    [TestClass]
    public sealed class Program
    {
        [DataTestMethod]
        [DataRow(12.4, 1.0, OeRoundMode.Near, 12.0)]
        [DataRow(12.5, 1.0, OeRoundMode.Near, 13.0)]
        [DataRow(-12.5, 1.0, OeRoundMode.Near, -13.0)]
        [DataRow(12.1, 5.0, OeRoundMode.Up, 15.0)]
        [DataRow(-12.1, 5.0, OeRoundMode.Up, -15.0)]
        [DataRow(12.9, 5.0, OeRoundMode.Down, 10.0)]
        [DataRow(-12.9, 5.0, OeRoundMode.Down, -10.0)]
        [DataRow(1.234, 0.01, OeRoundMode.Optimize, 1.23)]
        public void RoundMultipleHandlesModesAndSigns(
            double value,
            double unit,
            OeRoundMode mode,
            double expected)
        {
            Assert.AreEqual(expected, OeRound.RoundMultiple(value, unit, mode), 1e-9);
        }

        [TestMethod]
        public void RoundMultiplePreservesValueForInvalidUnits()
        {
            Assert.AreEqual(4.2, OeRound.RoundMultiple(4.2, 0, OeRoundMode.Near));
            Assert.AreEqual(4.2, OeRound.RoundMultiple(4.2, double.NaN, OeRoundMode.Near));
            Assert.AreEqual(4.2, OeRound.RoundMultiple(4.2, double.PositiveInfinity, OeRoundMode.Near));
        }

        [TestMethod]
        public void ConsistentRoundingMatchesRoundedExactSum()
        {
            double[] input = { 1.4, 1.4, 1.4, -0.2 };
            double[] result = OeRound.SolveConsistentSum(input, 1);

            Assert.AreEqual(input.Length, result.Length);
            Assert.AreEqual(OeRound.RoundMultiple(input.Sum(), 1, OeRoundMode.Near), result.Sum(), 1e-9);
            Assert.IsTrue(result.All(value => Math.Abs(value % 1) < 1e-9));
        }

        [TestMethod]
        public void ConsistentRoundingIsDeterministicForTiesAndNegativeValues()
        {
            double[] input = { -1.5, 1.5, 2.5, -2.5 };
            double[] first = OeRound.SolveConsistentSum(input, 1);
            double[] second = OeRound.SolveConsistentSum(input, 1);

            CollectionAssert.AreEqual(first, second);
            Assert.AreEqual(0, first.Sum(), 1e-9);
        }

        [TestMethod]
        public void ConsistentRoundingHandlesEmptyAndInvalidUnits()
        {
            Assert.AreEqual(0, OeRound.SolveConsistentSum(Array.Empty<double>(), 1).Length);
            CollectionAssert.AreEqual(
                new[] { 1.2, 3.4 },
                OeRound.SolveConsistentSum(new[] { 1.2, 3.4 }, 0));
        }

        [DataTestMethod]
        [DataRow("2024-02-29")]
        [DataRow("29.02.2024")]
        [DataRow("02/29/2024")]
        [DataRow("29 Februar 2024")]
        [DataRow("29 février 2024")]
        [DataRow("29 febbraio 2024")]
        [DataRow("29 febrero 2024")]
        public void DateValidatorAcceptsValidInternationalDates(string text)
        {
            Assert.IsTrue(DateValidator.LooksLikeDate(text));
            Assert.IsNull(DateValidator.GetInvalidReason(text));
        }

        [DataTestMethod]
        [DataRow("2023-02-29", "February 2023 only has 28 days")]
        [DataRow("31 November 2026", "November 2026 only has 30 days")]
        [DataRow("2026-13-01", "Month 13 does not exist")]
        [DataRow("00.12.2026", "Day 0 is invalid")]
        public void DateValidatorExplainsImpossibleDates(string text, string expectedFragment)
        {
            Assert.IsTrue(DateValidator.LooksLikeDate(text));
            StringAssert.Contains(DateValidator.GetInvalidReason(text), expectedFragment);
        }

        [DataTestMethod]
        [DataRow(null)]
        [DataRow("")]
        [DataRow("top 30 in Feb 2026")]
        [DataRow("not a date")]
        [DataRow("2026-01")]
        public void DateValidatorIgnoresNonDates(string text)
        {
            Assert.IsFalse(DateValidator.LooksLikeDate(text));
            Assert.IsNull(DateValidator.GetInvalidReason(text));
        }

        [TestMethod]
        public void AxisCalibrationMapsHorizontalAndVerticalCoordinates()
        {
            var horizontal = new AxisCalibration(AxisOrientation.Horizontal, 10, 0, 110, 50);
            var vertical = new AxisCalibration(AxisOrientation.Vertical, 200, 0, 100, 100);

            Assert.IsTrue(horizontal.IsValid);
            Assert.AreEqual(25, horizontal.ValueAt(60, 999), 1e-9);
            Assert.AreEqual(50, vertical.ValueAt(999, 150), 1e-9);
            Assert.AreEqual(150, vertical.ToValue(50), 1e-9);
        }

        [TestMethod]
        public void InvalidAxisCalibrationReturnsNaN()
        {
            var calibration = new AxisCalibration(AxisOrientation.Horizontal, 10, 0, 10, 100);
            Assert.IsFalse(calibration.IsValid);
            Assert.IsTrue(double.IsNaN(calibration.ToValue(10)));
        }

        [TestMethod]
        public void BarDetectorFindsColumnTipsInStableOrder()
        {
            using (var image = NewWhiteBitmap(80, 60))
            using (Graphics graphics = Graphics.FromImage(image))
            {
                graphics.FillRectangle(Brushes.RoyalBlue, 10, 30, 10, 25);
                graphics.FillRectangle(Brushes.RoyalBlue, 40, 15, 12, 40);

                var points = BarDetector.Detect(image, Color.RoyalBlue, 0, ChartKind.Columns);

                Assert.AreEqual(2, points.Count);
                Assert.AreEqual(new Point(14, 30), points[0]);
                Assert.AreEqual(new Point(45, 15), points[1]);
            }
        }

        [TestMethod]
        public void BarDetectorFindsHorizontalBarEnds()
        {
            using (var image = NewWhiteBitmap(80, 60))
            using (Graphics graphics = Graphics.FromImage(image))
            {
                graphics.FillRectangle(Brushes.DarkOrange, 5, 8, 30, 8);
                graphics.FillRectangle(Brushes.DarkOrange, 5, 35, 55, 10);

                var points = BarDetector.Detect(image, Color.DarkOrange, 0, ChartKind.Bars);

                Assert.AreEqual(2, points.Count);
                Assert.AreEqual(new Point(34, 11), points[0]);
                Assert.AreEqual(new Point(59, 39), points[1]);
            }
        }

        [TestMethod]
        public void DetectorsRespectToleranceAndEmptyInputs()
        {
            Assert.AreEqual(0, BarDetector.Detect(null, Color.Black, 0, ChartKind.Columns).Count);
            Assert.AreEqual(0, LineTracer.Trace(null, Color.Black, 0, 5).Count);

            using (var image = NewWhiteBitmap(20, 20))
            {
                image.SetPixel(5, 5, Color.FromArgb(100, 100, 100));
                Assert.IsFalse(BarDetector.BuildMatchMask(image, Color.FromArgb(105, 100, 100), 4)[5, 5]);
                Assert.IsTrue(BarDetector.BuildMatchMask(image, Color.FromArgb(105, 100, 100), 5)[5, 5]);
            }
        }

        [TestMethod]
        public void LineTracerSamplesSyntheticLine()
        {
            using (var image = NewWhiteBitmap(50, 40))
            {
                for (int x = 5; x <= 45; x++) image.SetPixel(x, 30 - x / 2, Color.ForestGreen);

                var points = LineTracer.Trace(image, Color.ForestGreen, 0, 5);

                Assert.AreEqual(5, points.Count);
                Assert.AreEqual(5, points.First().X);
                Assert.AreEqual(45, points.Last().X);
                Assert.IsTrue(points.Zip(points.Skip(1), (left, right) => right.Y <= left.Y).All(value => value));
            }
        }

        [TestMethod]
        public void PowerPointInstancePolicyFailsClosedWhenProcessesAreAmbiguous()
        {
            Assert.AreEqual(
                PowerPointInstanceSelection.None,
                PowerPointInstancePolicy.Select(Array.Empty<PowerPointWindowCandidate>()));
            Assert.AreEqual(
                PowerPointInstanceSelection.Single,
                PowerPointInstancePolicy.Select(new[]
                {
                    new PowerPointWindowCandidate(new IntPtr(20), 7),
                    new PowerPointWindowCandidate(new IntPtr(10), 7)
                }));
            Assert.AreEqual(
                PowerPointInstanceSelection.Ambiguous,
                PowerPointInstancePolicy.Select(new[]
                {
                    new PowerPointWindowCandidate(new IntPtr(10), 7),
                    new PowerPointWindowCandidate(new IntPtr(20), 8)
                }));
        }

        private static Bitmap NewWhiteBitmap(int width, int height)
        {
            var bitmap = new Bitmap(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap)) graphics.Clear(Color.White);
            return bitmap;
        }
    }
}
