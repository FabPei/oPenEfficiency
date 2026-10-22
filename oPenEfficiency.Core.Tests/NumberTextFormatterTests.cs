using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class NumberTextFormatterTests
    {
        [DataTestMethod]
        [DataRow("1,234.50", ".", ",", "Standard", "Remove", "", "1.234,50")]
        [DataRow("1.234,50", ",", ".", "Standard", "Remove", "", "1,234.50")]
        [DataRow("12345678.90", ",", ".", "Indian", "Remove", "", "1,23,45,678.90")]
        [DataRow("-1234", " ", ",", "Standard", "Remove", "", "-1 234")]
        [DataRow("€ 1.234,50", ",", ".", "Standard", "Prefix", "$", "$1,234.50")]
        [DataRow("$1,234.50", ".", ",", "Standard", "Suffix", "EUR", "1.234,50 EUR")]
        [DataRow("£12.50", ".", ",", "Standard", "Remove", "", "12,50")]
        public void NumberTokensAreNormalizedAndFormatted(
            string input,
            string thousand,
            string decimalSeparator,
            string grouping,
            string placement,
            string currency,
            string expected)
        {
            Assert.IsTrue(NumberTextFormatter.TryFormatToken(
                input,
                thousand,
                decimalSeparator,
                grouping,
                placement,
                currency,
                out string result));
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void InPlaceCurrencyAndInvalidTokensFollowExistingContract()
        {
            Assert.IsTrue(NumberTextFormatter.TryFormatToken(
                "12.50",
                ".",
                ",",
                "Standard",
                "InPlace",
                "€",
                out string result));
            Assert.AreEqual("12€50", result);

            Assert.IsFalse(NumberTextFormatter.TryFormatToken(
                "not-a-number",
                ",",
                ".",
                "Standard",
                "Remove",
                string.Empty,
                out result));
            Assert.AreEqual("not-a-number", result);
        }
    }
}
