using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// Pins the exact configured anchor-object-mode string mapping shared by every production and UI
    /// call site (<see cref="AnchorModes.IsLastSelected"/>). "Last Selected" (any casing) selects the
    /// last shape of the selection as the anchor object; every other value, including null and empty,
    /// keeps the historical first-shape default.
    /// </summary>
    [TestClass]
    public sealed class AnchorModeDecisionTests
    {
        [TestMethod]
        public void LastSelectedVariantsResolveToLastAnchor()
        {
            Assert.IsTrue(AnchorModes.IsLastSelected("Last Selected"));
            Assert.IsTrue(AnchorModes.IsLastSelected("last selected"));
            Assert.IsTrue(AnchorModes.IsLastSelected("LAST SELECTED"));
        }

        [TestMethod]
        public void EverythingElseResolvesToFirstAnchor()
        {
            Assert.IsFalse(AnchorModes.IsLastSelected("First selected"));
            Assert.IsFalse(AnchorModes.IsLastSelected("first selected"));
            Assert.IsFalse(AnchorModes.IsLastSelected("Last"));
            Assert.IsFalse(AnchorModes.IsLastSelected(string.Empty));
            Assert.IsFalse(AnchorModes.IsLastSelected(null));
        }
    }
}
