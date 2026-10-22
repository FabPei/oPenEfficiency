using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA9 wave 2 (tables): the host-neutral production code behind copy column widths, copy row heights and
    /// copy/paste table dimensions is <c>TableDimensionScaler.ScaleToTotal</c>. The COM leg of these commands is
    /// office-automated and blocked (QA-P0); everything asserted here runs in the canonical build.
    /// </summary>
    [TestClass]
    public class TableDimensionAcceptanceTests
    {
        [TestMethod]
        public void ScaleToTotalKeepsTheProportionsAndHitsTheTargetExactly()
        {
            var scaled = TableDimensionScaler.ScaleToTotal(new[] { 30f, 50f, 20f }, 200f);

            Assert.AreEqual(3, scaled.Length, "one scaled dimension per source dimension");
            Assert.AreEqual(60d, (double)scaled[0], 0.01d, "30 of 100 becomes 60 of 200");
            Assert.AreEqual(100d, (double)scaled[1], 0.01d, "50 of 100 becomes 100 of 200");
            Assert.AreEqual(40d, (double)scaled[2], 0.01d, "the last dimension carries the residual");
            Assert.AreEqual(200d, (double)scaled[0] + scaled[1] + scaled[2], 0.01d,
                "the sum is exactly the requested total, not merely close to it");
        }

        [TestMethod]
        public void ScaleToTotalIsAFixedPointWhenTheTotalDoesNotChange()
        {
            var once = TableDimensionScaler.ScaleToTotal(new[] { 17f, 4f, 29f, 3f }, 96f);
            var twice = TableDimensionScaler.ScaleToTotal(once, 96f);

            for (var index = 0; index < once.Length; index++)
            {
                Assert.AreEqual((double)once[index], twice[index], 0.001d,
                    $"copying the same dimensions onto the same total twice must not drift (index {index})");
            }
        }

        [TestMethod]
        public void ScaleToTotalRefusesDegenerateInput()
        {
            Assert.ThrowsException<ArgumentNullException>(() => TableDimensionScaler.ScaleToTotal(null, 100f));
            Assert.ThrowsException<ArgumentException>(() => TableDimensionScaler.ScaleToTotal(new float[0], 100f));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => TableDimensionScaler.ScaleToTotal(new[] { 10f }, 0f));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => TableDimensionScaler.ScaleToTotal(new[] { 10f, 0f }, 100f));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => TableDimensionScaler.ScaleToTotal(new[] { 10f, -4f }, 100f));
        }

        [TestMethod]
        public void ScaleToTotalAppliesTheTargetToASingleDimension()
        {
            var scaled = TableDimensionScaler.ScaleToTotal(new[] { 44f }, 88f);

            Assert.AreEqual(1, scaled.Length, "a single dimension stays a single dimension");
            Assert.AreEqual(88d, (double)scaled[0], 0.001d, "with one dimension the target total is the dimension");
        }
    }
}
