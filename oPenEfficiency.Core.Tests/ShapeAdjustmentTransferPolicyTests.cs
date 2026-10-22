using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;
using static oPenEfficiency.Core.Services.ShapeAdjustmentTransferPolicy;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 align-and-size-wave coverage for Copy Shape Adjustments
    /// (QA9-BtnAlignShapeAdjustments-002): the explicit-reference resolution and the
    /// compatibility rule behind ShapeAdjustmentTransfer.
    /// </summary>
    [TestClass]
    public sealed class ShapeAdjustmentTransferPolicyTests
    {
        [TestMethod]
        public void WithoutAnExplicitChoiceTheConfiguredAnchorDecides()
        {
            Assert.IsNull(ResolveExplicitReferenceIndex(3, null));
            Assert.IsNull(ResolveExplicitReferenceIndex(1, null));
        }

        [TestMethod]
        public void TheExplicitChoicePicksTheFirstOrTheLastSelectedShape()
        {
            Assert.AreEqual(1, ResolveExplicitReferenceIndex(3, false));
            Assert.AreEqual(3, ResolveExplicitReferenceIndex(3, true));
            Assert.AreEqual(1, ResolveExplicitReferenceIndex(1, true));
        }

        [TestMethod]
        public void OnlyTheSameGeometryReceivesTheAdjustments()
        {
            Assert.IsTrue(IsTransferCompatible(true, true, 2, 2));
            Assert.IsFalse(IsTransferCompatible(false, true, 2, 2));  // different shape type
            Assert.IsFalse(IsTransferCompatible(true, false, 2, 2));  // different AutoShapeType
            Assert.IsFalse(IsTransferCompatible(true, true, 1, 2));   // different adjustment count
            Assert.IsTrue(IsTransferCompatible(true, true, 0, 0));    // the caller's own
                                                                      // count>0 precondition
                                                                      // guards the reference
        }
    }
}
