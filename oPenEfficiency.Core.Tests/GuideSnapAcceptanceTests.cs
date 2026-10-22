using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave 1 acceptance coverage for the directional align features (BtnAlignLeft, BtnAlignRight,
    /// BtnAlignTop, BtnAlignBottom, BtnAlignCenter, BtnAlignCenterVertical). Each of those features decides its target coordinate in
    /// <c>AlignmentHelpers.SnapToNearestGuide</c>, which delegates the decision to
    /// <see cref="DirectionalGuideSelector"/> - the host-neutral production code that is driven here. The
    /// fixture is a 16:9 slide (960 x 540 pt) with presentation guides. The COM leg of the features (writing
    /// shape.Left/shape.Top and the PowerPoint Undo entry) is the office-automated leg of the same feature and
    /// is not covered by this test.
    /// </summary>
    [TestClass]
    public sealed class GuideSnapAcceptanceTests
    {
        private const float SlideWidth = 960f;
        private const float SlideHeight = 540f;

        [TestMethod]
        public void AlignLeftSnapsToTheNearestVerticalGuideAtOrBeforeTheLeftEdge()
        {
            // 400 lies to the right of the shape, so it may never be chosen: the shape must not move right.
            float snapped = DirectionalGuideSelector.NearestAtOrBefore(new[] { 0f, 120f, 400f }, 305f, 0f);
            Assert.AreEqual(120f, snapped, Precision, "The nearest guide at or before the left edge must win.");
            Assert.IsTrue(snapped <= 305f, "Align Left moved the shape to the right.");

            // No guide at or before the edge: the slide's left edge is the fallback.
            Assert.AreEqual(0f, DirectionalGuideSelector.NearestAtOrBefore(new[] { 400f }, 305f, 0f), Precision);

            // A guide exactly on the edge keeps the shape in place, and re-running the same alignment is a
            // no-op - the documented behaviour is idempotent.
            float onEdge = DirectionalGuideSelector.NearestAtOrBefore(new[] { 305f }, 305f, 0f);
            Assert.AreEqual(305f, onEdge, Precision);
            Assert.AreEqual(onEdge, DirectionalGuideSelector.NearestAtOrBefore(new[] { 305f }, onEdge, 0f), Precision);
            Assert.AreEqual(120f, DirectionalGuideSelector.NearestAtOrBefore(new[] { 0f, 120f, 400f }, 120f, 0f), Precision);
        }

        [TestMethod]
        public void AlignRightSnapsToTheNearestVerticalGuideAtOrAfterTheRightEdge()
        {
            // Right edge at 405: the nearest guide at or after it is 500, closer than the slide's right edge.
            const float width = 100f;
            float rightEdge = 305f + width;
            float target = DirectionalGuideSelector.NearestAtOrAfter(new[] { 0f, 120f, 500f, 900f }, rightEdge, SlideWidth);
            Assert.AreEqual(500f, target, Precision);
            float left = target - width;
            Assert.AreEqual(400f, left, Precision);
            Assert.IsTrue(target >= rightEdge, "Align Right moved the right edge to the left.");

            // Guides to the left of the edge are ignored; the slide edge is the fallback.
            float fallback = DirectionalGuideSelector.NearestAtOrAfter(new[] { 0f, 120f, 300f }, rightEdge, SlideWidth);
            Assert.AreEqual(SlideWidth, fallback, Precision);
            Assert.IsTrue(fallback - width >= rightEdge, "The fallback moved the right edge to the left.");

            // A guide exactly on the right edge keeps the shape in place, and the alignment stays idempotent.
            float onEdge = DirectionalGuideSelector.NearestAtOrAfter(new[] { rightEdge }, rightEdge, SlideWidth);
            Assert.AreEqual(rightEdge, onEdge, Precision);
            Assert.AreEqual(onEdge, DirectionalGuideSelector.NearestAtOrAfter(new[] { rightEdge }, onEdge, SlideWidth), Precision);
        }

        [TestMethod]
        public void AlignTopSnapsToTheNearestHorizontalGuideAtOrBeforeTheTopEdge()
        {
            float snapped = DirectionalGuideSelector.NearestAtOrBefore(new[] { 0f, 100f, 480f }, 250f, 0f);
            Assert.AreEqual(100f, snapped, Precision, "The nearest guide at or before the top edge must win.");
            Assert.IsTrue(snapped <= 250f, "Align Top moved the shape downwards.");

            Assert.AreEqual(0f, DirectionalGuideSelector.NearestAtOrBefore(new[] { 480f }, 250f, 0f), Precision);

            float onEdge = DirectionalGuideSelector.NearestAtOrBefore(new[] { 250f }, 250f, 0f);
            Assert.AreEqual(250f, onEdge, Precision);
            Assert.AreEqual(onEdge, DirectionalGuideSelector.NearestAtOrBefore(new[] { 250f }, onEdge, 0f), Precision);
        }

        [TestMethod]
        public void AlignBottomSnapsToTheNearestHorizontalGuideAtOrAfterTheBottomEdge()
        {
            const float height = 80f;
            float bottomEdge = 250f + height;
            float target = DirectionalGuideSelector.NearestAtOrAfter(new[] { 0f, 200f, 500f }, bottomEdge, SlideHeight);
            Assert.AreEqual(500f, target, Precision);
            Assert.IsTrue(target >= bottomEdge, "Align Bottom moved the bottom edge upwards.");
            Assert.AreEqual(420f, target - height, Precision);

            float fallback = DirectionalGuideSelector.NearestAtOrAfter(new[] { 0f, 200f }, bottomEdge, SlideHeight);
            Assert.AreEqual(SlideHeight, fallback, Precision);
            Assert.IsTrue(fallback >= bottomEdge, "The fallback moved the bottom edge upwards.");

            float onEdge = DirectionalGuideSelector.NearestAtOrAfter(new[] { bottomEdge }, bottomEdge, SlideHeight);
            Assert.AreEqual(bottomEdge, onEdge, Precision);
            Assert.AreEqual(onEdge, DirectionalGuideSelector.NearestAtOrAfter(new[] { bottomEdge }, onEdge, SlideHeight), Precision);
        }

        [TestMethod]
        public void GuideSnapRejectsAMissingGuideCollection()
        {
            // The empty/invalid scope contract: the feature must not silently invent a coordinate when the
            // host handed over no guide data at all (null is the only invalid input the selector can see).
            Assert.ThrowsException<System.ArgumentNullException>(() =>
                DirectionalGuideSelector.NearestAtOrBefore(null, 300f, 0f));
            Assert.ThrowsException<System.ArgumentNullException>(() =>
                DirectionalGuideSelector.NearestAtOrAfter(null, 300f, SlideWidth));
            Assert.ThrowsException<System.ArgumentNullException>(() =>
                DirectionalGuideSelector.NearestByDistance(null, 300f, SlideWidth / 2f));

            // An empty presentation (no guides at all) still yields the slide edge, not "no move".
            Assert.AreEqual(0f, DirectionalGuideSelector.NearestAtOrBefore(Enumerable.Empty<float>(), 305f, 0f), Precision);
        }

        [TestMethod]
        public void AlignCenterSnapsTheHorizontalCenterToTheNearestVerticalGuide()
        {
            // Horizontal center 305: the nearest vertical guide by absolute distance is 400 - center
            // snapping may move the shape in either direction, unlike Align Left/Right.
            const float width = 100f;
            const float center = 305f;
            float best = DirectionalGuideSelector.NearestByDistance(new[] { 0f, 120f, 400f }, center, SlideWidth / 2f);
            Assert.AreEqual(400f, best, Precision);
            Assert.AreEqual(350f, best - width / 2f, Precision, "The shape's left must land so its center sits on the guide.");

            // No guide is closer than the slide center: the fallback wins and the shape centers on the slide.
            float fallback = DirectionalGuideSelector.NearestByDistance(new[] { 0f, 100f }, center, SlideWidth / 2f);
            Assert.AreEqual(SlideWidth / 2f, fallback, Precision);

            // A guide exactly on the center keeps the shape in place, and the alignment stays idempotent.
            float onCenter = DirectionalGuideSelector.NearestByDistance(new[] { 305f }, center, SlideWidth / 2f);
            Assert.AreEqual(305f, onCenter, Precision);
            Assert.AreEqual(onCenter, DirectionalGuideSelector.NearestByDistance(new[] { 305f }, onCenter, SlideWidth / 2f), Precision);

            // At equal distance a real guide wins over the slide-center fallback (the replaced COM route's
            // iteration order: guides first, fallback last).
            float tie = DirectionalGuideSelector.NearestByDistance(new[] { 130f }, center, SlideWidth / 2f);
            Assert.AreEqual(130f, tie, Precision);
        }

        [TestMethod]
        public void AlignCenterVerticalSnapsTheVerticalCenterToTheNearestHorizontalGuide()
        {
            const float height = 80f;
            const float middle = 250f;
            float best = DirectionalGuideSelector.NearestByDistance(new[] { 0f, 240f, 480f }, middle, SlideHeight / 2f);
            Assert.AreEqual(240f, best, Precision);
            Assert.AreEqual(200f, best - height / 2f, Precision, "The shape's top must land so its middle sits on the guide.");

            // The slide's middle participates as the fallback when no guide is closer.
            float fallback = DirectionalGuideSelector.NearestByDistance(new[] { 0f, 100f }, middle, SlideHeight / 2f);
            Assert.AreEqual(SlideHeight / 2f, fallback, Precision);

            float onMiddle = DirectionalGuideSelector.NearestByDistance(new[] { 250f }, middle, SlideHeight / 2f);
            Assert.AreEqual(250f, onMiddle, Precision);
            Assert.AreEqual(onMiddle, DirectionalGuideSelector.NearestByDistance(new[] { 250f }, onMiddle, SlideHeight / 2f), Precision);

            // A real guide at equal distance beats the slide-middle fallback.
            float tie = DirectionalGuideSelector.NearestByDistance(new[] { 230f }, middle, SlideHeight / 2f);
            Assert.AreEqual(230f, tie, Precision);
        }

        private const float Precision = 1e-5f;
    }
}
