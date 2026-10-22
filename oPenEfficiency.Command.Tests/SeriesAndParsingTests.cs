using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Features;
using oPenEfficiency.Features.Utilities;
using oPenEfficiency.Models;

namespace oPenEfficiency.Command.Tests
{
    [TestClass]
    public sealed class SeriesAndParsingTests
    {
        [TestMethod]
        public void NumberSeriesSupportsCountEndAndReverseDirection()
        {
            CollectionAssert.AreEqual(
                new[] { "2", "5", "8", "11" },
                Generate("Numbers", "2", null, 4, 3).ToArray());
            CollectionAssert.AreEqual(
                new[] { "5", "3", "1" },
                Generate("Numbers", "5", "1", 0, -2).ToArray());
        }

        [TestMethod]
        public void LetterSeriesCrossesAlphabetBoundaryAndPreservesCase()
        {
            CollectionAssert.AreEqual(
                new[] { "Y", "Z", "AA", "AB" },
                Generate("Letters", "Y", null, 4, 1).ToArray());
            CollectionAssert.AreEqual(
                new[] { "c", "b", "a" },
                Generate("Letters", "c", "a", 0, -1).ToArray());
        }

        [TestMethod]
        public void DateSeriesHandlesLeapDaysAndMonthBoundaries()
        {
            var days = new SeriesConfig
            {
                SeriesType = "Days",
                StartValue = "2024-02-28",
                Count = 3,
                Step = 1,
                FormatString = "yyyy-MM-dd",
                Language = "en-US"
            };
            var months = new SeriesConfig
            {
                SeriesType = "Months",
                StartValue = "2024-01-31",
                Count = 3,
                Step = 1,
                FormatString = "yyyy-MM-dd",
                Language = "en-US"
            };

            CollectionAssert.AreEqual(
                new[] { "2024-02-28", "2024-02-29", "2024-03-01" },
                SeriesGeneratorFeature.GenerateSeriesData(days).ToArray());
            CollectionAssert.AreEqual(
                new[] { "2024-01-31", "2024-02-29", "2024-03-31" },
                SeriesGeneratorFeature.GenerateSeriesData(months).ToArray());
        }

        [TestMethod]
        public void QuarterSeriesCrossesYearBoundaryAndFormatsTokens()
        {
            var config = new SeriesConfig
            {
                SeriesType = "Quarters",
                StartValue = "2025-Q3",
                EndValue = "2026-Q2",
                Count = 0,
                Step = 1,
                FormatString = "{YYYY}-{Q}"
            };

            CollectionAssert.AreEqual(
                new[] { "2025-Q3", "2025-Q4", "2026-Q1", "2026-Q2" },
                SeriesGeneratorFeature.GenerateSeriesData(config).ToArray());
        }

        [TestMethod]
        public void SeriesGenerationRejectsInvalidConfigurationAndCapsOutput()
        {
            Assert.AreEqual(0, SeriesGeneratorFeature.GenerateSeriesData(null).Count);
            Assert.AreEqual(0, Generate("Numbers", "1", null, 5, 0).Count);
            Assert.AreEqual(0, Generate("Numbers", "not-a-number", null, 5, 1).Count);
            Assert.AreEqual(1000, Generate("Numbers", "1", null, 5000, 1).Count);
        }

        [DataTestMethod]
        [DataRow("1,234.50", 1234.5)]
        [DataRow("1.234,50", 1234.5)]
        [DataRow("1\u00a0234,5", 1234.5)]
        [DataRow("(1,234.5)", -1234.5)]
        [DataRow("12.5%", 12.5)]
        [DataRow("€1,234.50", 1234.5)]
        public void TableNumberParsingAcceptsSupportedFormats(string text, double expected)
        {
            Assert.IsTrue(TableHelper.TryParseDoubleFromCell(text, out double result));
            Assert.AreEqual(expected, result, 1e-9);
        }

        [DataTestMethod]
        [DataRow(null)]
        [DataRow("")]
        [DataRow("text")]
        [DataRow("NaN")]
        [DataRow("Infinity")]
        public void TableNumberParsingRejectsInvalidValues(string text)
        {
            Assert.IsFalse(TableHelper.TryParseDoubleFromCell(text, out _));
        }

        [TestMethod]
        public void JsonSerializationRoundTripsAndRejectsMalformedInput()
        {
            var source = new SeriesConfig
            {
                SeriesId = "series-ä",
                SeriesType = "Numbers",
                StartValue = "1",
                Count = 5,
                Step = 2
            };

            string json = JsonService.Serialize(source);
            SeriesConfig clone = JsonService.Deserialize<SeriesConfig>(json);

            Assert.AreEqual(source.SeriesId, clone.SeriesId);
            Assert.AreEqual(source.Count, clone.Count);
            Assert.AreEqual(source.Step, clone.Step);
            Assert.IsNull(JsonService.Deserialize<SeriesConfig>("{broken"));
            Assert.IsNull(JsonService.Deserialize<SeriesConfig>(null));
            Assert.AreEqual(string.Empty, JsonService.Serialize<SeriesConfig>(null));
        }

        private static System.Collections.Generic.List<string> Generate(
            string type,
            string start,
            string end,
            int count,
            int step)
        {
            return SeriesGeneratorFeature.GenerateSeriesData(new SeriesConfig
            {
                SeriesType = type,
                StartValue = start,
                EndValue = end,
                Count = count,
                Step = step,
                Language = "en-US"
            });
        }
    }
}
