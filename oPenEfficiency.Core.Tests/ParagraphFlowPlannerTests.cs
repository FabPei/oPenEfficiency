using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA9 wave 3 (text and formatting): the host-free paragraph-flow rules behind Split Text Box by Paragraphs
    /// (including Merge Selected Text Boxes) and the paragraph chooser of Swap Text (Plain). The commands call
    /// <c>ParagraphFlowPlanner</c>; their COM legs are office-automated and QA-P0-blocked.
    /// </summary>
    [TestClass]
    public class ParagraphFlowPlannerTests
    {
        [TestMethod]
        public void ReadingOrderUsesTheDocumentedRowTolerance()
        {
            Assert.IsTrue(ParagraphFlowPlanner.CompareReadingOrder(10f, 100f, 15f, 50f) > 0,
                "within the 5 pt tolerance the shapes count as one row and the left edge decides");
            Assert.IsTrue(ParagraphFlowPlanner.CompareReadingOrder(10f, 50f, 15f, 100f) < 0,
                "the same row orders left to right");
            Assert.IsTrue(ParagraphFlowPlanner.CompareReadingOrder(10f, 100f, 15.1f, 50f) < 0,
                "beyond the tolerance the top edge decides, so a lower shape stays below");
            Assert.IsTrue(ParagraphFlowPlanner.CompareReadingOrder(30f, 100f, 15.1f, 50f) > 0,
                "a later row sorts after an earlier one regardless of the left edge");
            Assert.AreEqual(0, ParagraphFlowPlanner.CompareReadingOrder(10f, 50f, 10f, 50f),
                "identical positions compare equal");
        }

        [TestMethod]
        public void ParagraphCollectionKeepsOrDropsEmptyParagraphsAndStripsEndings()
        {
            Assert.AreEqual("Text", ParagraphFlowPlanner.WithoutParagraphEnding("Text\r\n"));
            Assert.AreEqual("Text", ParagraphFlowPlanner.WithoutParagraphEnding("Text"));
            Assert.AreEqual(string.Empty, ParagraphFlowPlanner.WithoutParagraphEnding(null));

            Assert.IsFalse(ParagraphFlowPlanner.KeepParagraphForSplit("   \r\n", false),
                "an empty paragraph is dropped by default");
            Assert.IsTrue(ParagraphFlowPlanner.KeepParagraphForSplit("   \r\n", true),
                "the keep-empty route retains it");
            Assert.IsTrue(ParagraphFlowPlanner.KeepParagraphForSplit("Text", false));

            Assert.AreEqual("A\rB\rC",
                ParagraphFlowPlanner.JoinParagraphs(new[] { "A\r\n", "B\r", "C" }),
                "merged paragraphs are joined with a single paragraph mark each");
            Assert.AreEqual(string.Empty, ParagraphFlowPlanner.JoinParagraphs(new string[0]));
        }

        [TestMethod]
        public void ParagraphChoicesDescribeAndTruncateLabels()
        {
            Assert.AreEqual("1: Hello", ParagraphFlowPlanner.DescribeParagraphChoice(1, "Hello\r\n"),
                "the paragraph mark never leaks into the label");
            Assert.AreEqual("2: a  b", ParagraphFlowPlanner.DescribeParagraphChoice(2, "a\r\nb"),
                "each interior break becomes its own space");
            Assert.AreEqual("3: (empty paragraph)", ParagraphFlowPlanner.DescribeParagraphChoice(3, " \r\n"),
                "a blank paragraph gets the documented placeholder");

            string longText = new string('x', 120);
            string label = ParagraphFlowPlanner.DescribeParagraphChoice(9, longText);
            Assert.AreEqual("9: " + new string('x', 97) + "...", label,
                "labels longer than 100 characters are truncated to 97 plus an ellipsis");
            Assert.AreEqual("9: ".Length + ParagraphFlowPlanner.ChoiceLabelMaximumLength, label.Length,
                "the truncated description keeps the documented budget");
        }
    }
}
