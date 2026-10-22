using System;
using System.Globalization;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class TableTextPlannerTests
    {
        [TestMethod]
        public void SortRowsUsesNumericOrderAndMovesWholeRows()
        {
            string[][] rows =
            {
                new[] { "10", "ten" },
                new[] { "2", "two" },
                new[] { "1.5", "one" }
            };

            var result = TableTextPlanner.SortRows(rows, 0, true);

            CollectionAssert.AreEqual(new[] { "one", "two", "ten" }, result.Select(row => row[1]).ToArray());
            Assert.AreEqual("10", rows[0][0], "Sorting mutated the source rows.");
        }

        [TestMethod]
        public void SortRowsIsStableAndSupportsDescendingMixedValues()
        {
            string[][] rows =
            {
                new[] { "a", "first" },
                new[] { "A", "second" },
                new[] { "2", "numeric" }
            };

            var ascending = TableTextPlanner.SortRows(rows, 0, true);
            var descending = TableTextPlanner.SortRows(rows, 0, false);

            CollectionAssert.AreEqual(new[] { "numeric", "first", "second" }, ascending.Select(row => row[1]).ToArray());
            CollectionAssert.AreEqual(new[] { "first", "second", "numeric" }, descending.Select(row => row[1]).ToArray());
        }

        [TestMethod]
        public void TotalsPlannerCalculatesRowsColumnsAndGrandTotal()
        {
            var cells = new[,]
            {
                { "1", "2", "old" },
                { "3", "", "old" },
                { "old", "old", "old" }
            };

            var totals = TableTextPlanner.CalculateTotals(cells, true, true);

            Assert.AreEqual(5, totals.Count);
            Assert.AreEqual(4, Find(totals, 2, 0));
            Assert.AreEqual(2, Find(totals, 2, 1));
            Assert.AreEqual(3, Find(totals, 0, 2));
            Assert.AreEqual(3, Find(totals, 1, 2));
            Assert.AreEqual(6, Find(totals, 2, 2));
        }

        [TestMethod]
        public void TotalsPlannerSupportsSingleDirection()
        {
            var cells = new[,] { { "1", "2" }, { "3", "4" }, { "", "" } };
            var columnTotals = TableTextPlanner.CalculateTotals(cells, false, true);
            Assert.AreEqual(2, columnTotals.Count);
            Assert.AreEqual(4, Find(columnTotals, 2, 0));
            Assert.AreEqual(6, Find(columnTotals, 2, 1));

            var rowTotals = TableTextPlanner.CalculateTotals(new[,] { { "1", "2", "" }, { "3", "4", "" } }, true, false);
            Assert.AreEqual(3, Find(rowTotals, 0, 2));
            Assert.AreEqual(7, Find(rowTotals, 1, 2));
        }

        [TestMethod]
        public void TotalsPlannerReportsExactInvalidCellBeforeProducingWrites()
        {
            var cells = new[,] { { "1", "bad", "" }, { "", "", "" } };

            FormatException exception = Assert.ThrowsException<FormatException>(() =>
                TableTextPlanner.CalculateTotals(cells, true, true, 4, 27));

            StringAssert.Contains(exception.Message, "AB4");
            StringAssert.Contains(exception.Message, "bad");
        }

        [TestMethod]
        public void TablePlannerRejectsInvalidShapesAndSortColumns()
        {
            Assert.ThrowsException<ArgumentException>(() =>
                TableTextPlanner.SortRows(new[] { new[] { "a" }, new[] { "b", "c" } }, 0, true));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
                TableTextPlanner.SortRows(new[] { new[] { "a" } }, 1, true));
            Assert.AreEqual(0, TableTextPlanner.CalculateTotals(new[,] { { "1" } }, true, true).Count);
        }

        [TestMethod]
        public void FlexibleNumberParsingDoesNotDependOnProcessCulture()
        {
            CultureInfo originalCulture = CultureInfo.CurrentCulture;
            try
            {
                foreach (string cultureName in new[] { "en-US", "de-DE", "fr-FR" })
                {
                    CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(cultureName);
                    Assert.IsTrue(FlexibleNumberParser.TryParse("1,234.5", out double us));
                    Assert.IsTrue(FlexibleNumberParser.TryParse("1.234,5", out double german));
                    Assert.IsTrue(FlexibleNumberParser.TryParse("1\u00a0234,5", out double french));
                    Assert.AreEqual(1234.5, us, 1e-9, cultureName);
                    Assert.AreEqual(1234.5, german, 1e-9, cultureName);
                    Assert.AreEqual(1234.5, french, 1e-9, cultureName);
                }
            }
            finally
            {
                CultureInfo.CurrentCulture = originalCulture;
            }
        }

        [TestMethod]
        public void AmbiguousSingleSeparatorUsesCurrentCulture()
        {
            CultureInfo originalCulture = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
                Assert.IsTrue(FlexibleNumberParser.TryParse("1,234", out double english));
                Assert.AreEqual(1234, english, 1e-9);

                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
                Assert.IsTrue(FlexibleNumberParser.TryParse("1,234", out double german));
                Assert.AreEqual(1.234, german, 1e-9);
            }
            finally
            {
                CultureInfo.CurrentCulture = originalCulture;
            }
        }

        private static double Find(System.Collections.Generic.IEnumerable<TableCellTotal> totals, int row, int column)
            => totals.Single(total => total.Row == row && total.Column == column).Value;
    }
}
