using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA9 wave 3 (text and formatting): the host-free matching rules behind Replace Formatting (BtnReplaceFont).
    /// The command builds its change list through <c>BulkFormatMatchPolicy</c>, so these assertions drive the
    /// production decision; the COM leg of the command is office-automated and QA-P0-blocked.
    /// </summary>
    [TestClass]
    public class BulkFormatMatchPolicyTests
    {
        [TestMethod]
        public void FontMatchIsCaseInsensitiveAndSkipsNoOps()
        {
            Assert.IsTrue(BulkFormatMatchPolicy.MatchesFont("Arial", "ARIAL", "Calibri"),
                "the font criterion matches case-insensitively");
            Assert.IsFalse(BulkFormatMatchPolicy.MatchesFont("Arial", "ARIAL", "Arial"),
                "a replacement whose name equals the run's own name exactly is a no-op");
            Assert.IsTrue(BulkFormatMatchPolicy.MatchesFont("Arial", "ARIAL", "ARIAL"),
                "the no-op guard compares exactly: a pure case change is still applied");
            Assert.IsFalse(BulkFormatMatchPolicy.MatchesFont("Arial", "Calibri", "Verdana"),
                "a run with another font stays untouched");
            Assert.IsFalse(BulkFormatMatchPolicy.MatchesFont("Arial", "  ", "Calibri"),
                "a blank find value never matches");
            Assert.IsFalse(BulkFormatMatchPolicy.MatchesFont("Arial", "Arial", ""),
                "a blank replace value never matches");
        }

        [TestMethod]
        public void SizeMatchUsesTheDocumentedTolerance()
        {
            Assert.IsTrue(BulkFormatMatchPolicy.MatchesSize(12.09f, 12f, 24f),
                "a run within the 0.1 pt band matches");
            Assert.IsFalse(BulkFormatMatchPolicy.MatchesSize(12.15f, 12f, 24f),
                "a run outside the 0.1 pt band stays untouched");
            Assert.IsFalse(BulkFormatMatchPolicy.MatchesSize(12f, 12f, 12.005f),
                "a replace size within 0.01 pt of the run's own size is a no-op");
            Assert.IsTrue(BulkFormatMatchPolicy.MatchesSize(12.005f, 12f, 24f),
                "the match band itself has no lower no-op guard - only the replacement does");
            Assert.IsFalse(BulkFormatMatchPolicy.MatchesSize(12f, null, 24f),
                "a missing find size never matches");
            Assert.IsFalse(BulkFormatMatchPolicy.MatchesSize(12f, 12f, null),
                "a missing replace size never matches");
        }

        [TestMethod]
        public void ColorMatchRequiresEqualityVisibilityAndSkipsNoOps()
        {
            Assert.IsTrue(BulkFormatMatchPolicy.MatchesColor(0x0000FF, 0x0000FF, 0xFF0000, true),
                "an equal visible fill or line color matches");
            Assert.IsFalse(BulkFormatMatchPolicy.MatchesColor(0x0000FF, 0x0000FF, 0xFF0000, false),
                "an invisible fill or line never matches");
            Assert.IsFalse(BulkFormatMatchPolicy.MatchesColor(0x00FF00, 0x0000FF, 0xFF0000, true),
                "a shape with another color stays untouched");
            Assert.IsFalse(BulkFormatMatchPolicy.MatchesColor(0x0000FF, 0x0000FF, 0x0000FF, true),
                "a same-color replacement is a no-op");
            Assert.IsFalse(BulkFormatMatchPolicy.MatchesColor(0x0000FF, null, 0xFF0000, true),
                "a missing find color never matches");
            Assert.IsTrue(BulkFormatMatchPolicy.MatchesColor(0x0000FF, 0x0000FF, 0xFF0000),
                "the run-color overload needs no visibility flag");
        }
    }
}
