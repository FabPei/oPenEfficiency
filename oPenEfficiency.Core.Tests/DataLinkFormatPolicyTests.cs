using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// Pins the pure decisions behind the Excel-to-PowerPoint cell-format transfer
    /// (<see cref="DataLinkFormatPolicy"/> and <see cref="DataLinkFillPolicy"/>) that the link
    /// creation and the link refresh now share (FEAT-005 DL-FORMAT-CREATE-002).
    /// </summary>
    [TestClass]
    public sealed class DataLinkFormatPolicyTests
    {
        [TestMethod]
        public void ExcelNoFillIsDetectedAsAbsentFillTest()
        {
            Assert.IsFalse(DataLinkFormatPolicy.HasExplicitFill(DataLinkFormatPolicy.XlLineStyleNone),
                "xlNone (-4142) is Excel 'No Fill'.");
            Assert.IsFalse(DataLinkFormatPolicy.HasExplicitFill(-4142),
                "The xlNone value itself must stay the No-Fill marker.");
            Assert.IsTrue(DataLinkFormatPolicy.HasExplicitFill(3),
                "A palette-backed fill is an explicit fill.");
        }

        [TestMethod]
        public void NoFillSourceReachesTheDestinationAsExplicitWhiteTest()
        {
            // The owner-visible defect: Excel "No Fill" must not arrive as an invisible fill,
            // because an invisible fill exposes the PowerPoint table style's own gray.
            bool hasFill = DataLinkFormatPolicy.HasExplicitFill(DataLinkFormatPolicy.XlLineStyleNone);
            int sourceBgr = DataLinkFormatPolicy.ToBgr(0.0);

            Assert.AreEqual(DataLinkFillPolicy.NoFillBgr, DataLinkFillPolicy.ResolveFillBgr(hasFill, sourceBgr),
                "A No-Fill source cell resolves to the explicit white fill.");
            Assert.AreEqual(0xFFFFFF, DataLinkFillPolicy.NoFillBgr, "Explicit white is 0xFFFFFF in BGR.");

            // An explicit Excel fill keeps Excel's own color: red is 255 in BGR.
            bool redFill = DataLinkFormatPolicy.HasExplicitFill(3);
            int redBgr = DataLinkFormatPolicy.ToBgr(255.0);
            Assert.AreEqual(255, redBgr, "Excel's BGR red is 255.");
            Assert.AreEqual(255, DataLinkFillPolicy.ResolveFillBgr(redFill, redBgr),
                "An explicit Excel fill keeps its BGR color.");
        }

        [TestMethod]
        public void ExcelOleColorsAreNormalisedToBgrTest()
        {
            // Excel exposes Font.Color / Interior.Color as a double, so a cast would be wrong.
            Assert.AreEqual(255, DataLinkFormatPolicy.ToBgr(255.0));
            Assert.AreEqual(16711680, DataLinkFormatPolicy.ToBgr(16711680.0));
            Assert.AreEqual(65280, DataLinkFormatPolicy.ToBgr("65280"));
            Assert.AreEqual(0, DataLinkFormatPolicy.ToBgr(null));
            Assert.AreEqual(0, DataLinkFormatPolicy.ToBgr(System.DBNull.Value));
        }

        [TestMethod]
        public void MixedOrMissingBoldStaysNonBoldTest()
        {
            Assert.IsTrue(DataLinkFormatPolicy.ToBoolean(true));
            Assert.IsTrue(DataLinkFormatPolicy.ToBoolean("True"));
            Assert.IsFalse(DataLinkFormatPolicy.ToBoolean(null),
                "Excel returns null for a mixed range; a mixed cell must not turn bold.");
            Assert.IsFalse(DataLinkFormatPolicy.ToBoolean(System.DBNull.Value));
            Assert.IsFalse(DataLinkFormatPolicy.ToBoolean("not a bool"));
        }

        [DataTestMethod]
        [DataRow(DataLinkFormatPolicy.XlHAlignGeneral, false, DataLinkAlignment.Left)]
        [DataRow(DataLinkFormatPolicy.XlHAlignGeneral, true, DataLinkAlignment.Right)]
        [DataRow(DataLinkFormatPolicy.XlHAlignLeft, false, DataLinkAlignment.Left)]
        [DataRow(DataLinkFormatPolicy.XlHAlignCenter, false, DataLinkAlignment.Center)]
        [DataRow(DataLinkFormatPolicy.XlHAlignCenterAcrossSelection, false, DataLinkAlignment.Center)]
        [DataRow(DataLinkFormatPolicy.XlHAlignRight, false, DataLinkAlignment.Right)]
        [DataRow(DataLinkFormatPolicy.XlHAlignJustify, false, DataLinkAlignment.Justify)]
        [DataRow(DataLinkFormatPolicy.XlHAlignDistributed, false, DataLinkAlignment.Justify)]
        [DataRow(-9999, false, DataLinkAlignment.Left)]
        [DataRow(-9999, true, DataLinkAlignment.Right)]
        public void HorizontalAlignmentMappingTest(int xlHAlign, bool generalIsNumeric, DataLinkAlignment expected)
        {
            Assert.AreEqual(expected, DataLinkFormatPolicy.MapAlignment(xlHAlign, generalIsNumeric));
        }

        [DataTestMethod]
        [DataRow(DataLinkFormatPolicy.XlVAlignTop, DataLinkVerticalAnchor.Top)]
        [DataRow(DataLinkFormatPolicy.XlVAlignCenter, DataLinkVerticalAnchor.Middle)]
        [DataRow(DataLinkFormatPolicy.XlVAlignBottom, DataLinkVerticalAnchor.Bottom)]
        [DataRow(-9999, DataLinkVerticalAnchor.Middle)]
        public void VerticalAnchorMappingTest(int xlVAlign, DataLinkVerticalAnchor expected)
        {
            Assert.AreEqual(expected, DataLinkFormatPolicy.MapVerticalAnchor(xlVAlign));
        }

        [DataTestMethod]
        [DataRow(DataLinkFormatPolicy.XlBorderWeightHairline, 0.5f)]
        [DataRow(DataLinkFormatPolicy.XlBorderWeightThin, 1.0f)]
        [DataRow(DataLinkFormatPolicy.XlBorderWeightMedium, 2.0f)]
        [DataRow(DataLinkFormatPolicy.XlBorderWeightThick, 3.0f)]
        [DataRow(-9999, 1.0f)]
        public void BorderWeightMappingTest(int xlWeight, float expected)
        {
            Assert.AreEqual(expected, DataLinkFormatPolicy.MapBorderWeight(xlWeight));
        }

        [TestMethod]
        public void BorderVisibilityMatchesExcelLineStyleTest()
        {
            Assert.IsFalse(DataLinkFormatPolicy.HasVisibleBorder(DataLinkFormatPolicy.XlLineStyleNone),
                "xlLineStyleNone means the border draws no line.");
            Assert.IsTrue(DataLinkFormatPolicy.HasVisibleBorder(1),
                "Any other line style draws a visible border.");
        }
    }
}
