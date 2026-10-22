using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class DelimitedSettingTokensTests
    {
        [TestMethod]
        public void ParseAndSerializePreserveLegacyCommaDelimitedContract()
        {
            var tokens = DelimitedSettingTokens.Parse(" #ff0000, Arial, #FF0000, Calibri ");

            CollectionAssert.AreEqual(new[] { "#ff0000", "Arial", "Calibri" }, (System.Collections.ICollection)tokens);
            Assert.AreEqual("#ff0000, Arial, Calibri", DelimitedSettingTokens.Serialize(tokens));
        }

        [DataTestMethod]
        [DataRow("#abcdef", "#ABCDEF")]
        [DataRow(" #80abcdef ", "#80ABCDEF")]
        public void NormalizeHexColorAcceptsSixAndEightDigitTokens(string input, string expected)
        {
            Assert.IsTrue(DelimitedSettingTokens.TryNormalizeHexColor(input, out string actual));
            Assert.AreEqual(expected, actual);
        }

        [DataTestMethod]
        [DataRow("")]
        [DataRow("red")]
        [DataRow("#12345")]
        [DataRow("#GG0000")]
        public void NormalizeHexColorRejectsInvalidTokens(string input)
        {
            Assert.IsFalse(DelimitedSettingTokens.TryNormalizeHexColor(input, out _));
        }

        [TestMethod]
        public void NormalizeFontNameRejectsCommaAndTrimsValidName()
        {
            Assert.IsFalse(DelimitedSettingTokens.TryNormalizeFontName("Arial, Calibri", out _));
            Assert.IsTrue(DelimitedSettingTokens.TryNormalizeFontName("  Aptos Display  ", out string normalized));
            Assert.AreEqual("Aptos Display", normalized);
        }
    }
}
