using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave 1 acceptance coverage for the incremental spacing features
    /// (BtnIncreaseHorizontalSpacing, BtnDecreaseHorizontalSpacing, BtnIncreaseVerticalSpacing,
    /// BtnDecreaseVerticalSpacing). Each of them hands the selected geometry and the configured spacing
    /// increment (sidebar setting, default 10) to <c>ShapeGeometryPlanner.AdjustSpacing</c>, which is the
    /// host-neutral production code driven here. The COM leg (reading the selection, writing shape.Left /
    /// shape.Top, restoring the selection) stays office-automated.
    /// </summary>
    [TestClass]
    public sealed class SpacingAdjustmentAcceptanceTests
    {
        private const float Increment = 10f;
        private const float Precision = 1e-5f;

        [TestMethod]
        public void IncreaseHorizontalSpacingAddsTheIncrementBetweenNeighbours()
        {
            var source = HorizontalSource();

            var planned = ShapeGeometryPlanner.AdjustSpacing(source, Increment, horizontal: true, preserveCenter: false);

            // Documented behaviour: the configured increment is applied between each neighbouring pair, from
            // the left-most shape outwards - so the first spatial shape keeps its position.
            Assert.AreEqual(20f, ById(planned, 1).Left, Precision, "The left-most shape moved.");
            AssertGap(planned, 1, 2, 60f);
            AssertGap(planned, 2, 3, 70f);
            Assert.AreEqual(50f + Increment, Gap(planned, 1, 2), Precision);
            Assert.AreEqual(60f + Increment, Gap(planned, 2, 3), Precision);

            // Only the driven axis may move, and the input must stay untouched.
            Assert.AreEqual(5f, ById(planned, 1).Top, Precision);
            Assert.AreEqual(30f, ById(planned, 2).Top, Precision);
            Assert.AreEqual(60f, ById(planned, 3).Top, Precision);
            Assert.AreEqual(80f, source[1].Left, Precision, "Planning mutated the source geometry.");

            // Increase and decrease are exact inverses: the documented decrement route restores the layout.
            var restored = ShapeGeometryPlanner.AdjustSpacing(planned, -Increment, horizontal: true, preserveCenter: false);
            AssertGeometry(restored, ById(source, 1), ById(source, 2), ById(source, 3));
        }

        [TestMethod]
        public void DecreaseHorizontalSpacingRemovesTheIncrementBetweenNeighbours()
        {
            var source = HorizontalSource();

            var planned = ShapeGeometryPlanner.AdjustSpacing(source, -Increment, horizontal: true, preserveCenter: false);

            Assert.AreEqual(20f, ById(planned, 1).Left, Precision, "The left-most shape moved.");
            Assert.AreEqual(40f, Gap(planned, 1, 2), Precision, "The decrement was not applied between the first pair.");
            Assert.AreEqual(50f, Gap(planned, 2, 3), Precision, "The decrement was not applied between the second pair.");
            Assert.AreEqual(5f, ById(planned, 1).Top, Precision);
            Assert.AreEqual(60f, ById(planned, 3).Top, Precision);
        }

        [TestMethod]
        public void IncreaseVerticalSpacingAddsTheIncrementBetweenNeighbours()
        {
            var source = VerticalSource();

            var planned = ShapeGeometryPlanner.AdjustSpacing(source, Increment, horizontal: false, preserveCenter: false);

            Assert.AreEqual(20f, ById(planned, 1).Top, Precision, "The top-most shape moved.");
            Assert.AreEqual(50f + Increment, VerticalGap(planned, 1, 2), Precision);
            Assert.AreEqual(60f + Increment, VerticalGap(planned, 2, 3), Precision);
            Assert.AreEqual(5f, ById(planned, 1).Left, Precision);
            Assert.AreEqual(30f, ById(planned, 2).Left, Precision);
            Assert.AreEqual(60f, ById(planned, 3).Left, Precision);

            var restored = ShapeGeometryPlanner.AdjustSpacing(planned, -Increment, horizontal: false, preserveCenter: false);
            AssertGeometry(restored, ById(source, 1), ById(source, 2), ById(source, 3));
        }

        [TestMethod]
        public void DecreaseVerticalSpacingRemovesTheIncrementBetweenNeighbours()
        {
            var source = VerticalSource();

            var planned = ShapeGeometryPlanner.AdjustSpacing(source, -Increment, horizontal: false, preserveCenter: false);

            Assert.AreEqual(20f, ById(planned, 1).Top, Precision, "The top-most shape moved.");
            Assert.AreEqual(40f, VerticalGap(planned, 1, 2), Precision);
            Assert.AreEqual(50f, VerticalGap(planned, 2, 3), Precision);
            Assert.AreEqual(5f, ById(planned, 1).Left, Precision);
            Assert.AreEqual(60f, ById(planned, 3).Left, Precision);
        }

        [TestMethod]
        public void CentredSpacingMovesTheOuterShapesEquallyAndKeepsTheCentre()
        {
            var source = HorizontalSource();
            float originalCentre = Centre(source, horizontal: true);

            var planned = ShapeGeometryPlanner.AdjustSpacing(source, -Increment, horizontal: true, preserveCenter: true);

            // The documented centred route moves the shapes equally towards the selection midpoint instead of
            // anchoring the first spatial shape.
            Assert.AreEqual(30f, ById(planned, 1).Left, Precision);
            Assert.AreEqual(80f, ById(planned, 2).Left, Precision, "The middle shape of an odd count must not move.");
            Assert.AreEqual(150f, ById(planned, 3).Left, Precision);
            Assert.AreEqual(originalCentre, Centre(planned, horizontal: true), Precision, "The centred route moved the selection centre.");
            Assert.AreEqual(-(ById(planned, 1).Left - ById(source, 1).Left),
                            ById(planned, 3).Left - ById(source, 3).Left, Precision,
                            "The outer shapes did not move symmetrically.");

            // Refusal paths: no source at all is refused, a single shape and a zero increment change nothing.
            Assert.ThrowsException<ArgumentNullException>(() =>
                ShapeGeometryPlanner.AdjustSpacing(null, Increment, horizontal: true, preserveCenter: false));
            Assert.ThrowsException<ArgumentException>(() =>
                ShapeGeometryPlanner.AdjustSpacing(new List<ShapeGeometry>(), Increment, horizontal: true, preserveCenter: false));
            var single = ShapeGeometryPlanner.AdjustSpacing(
                new List<ShapeGeometry> { Shape(7, 40, 50, 20, 10) }, Increment, horizontal: true, preserveCenter: false);
            Assert.AreEqual(40f, single[0].Left, Precision, "A single shape must not be moved.");
            var unchanged = ShapeGeometryPlanner.AdjustSpacing(source, 0f, horizontal: true, preserveCenter: false);
            AssertGeometry(unchanged, ById(source, 1), ById(source, 2), ById(source, 3));
        }

        private static List<ShapeGeometry> HorizontalSource() => Shapes(
            Shape(1, 20, 5, 10, 5),
            Shape(2, 80, 30, 20, 10),
            Shape(3, 160, 60, 30, 20));

        private static List<ShapeGeometry> VerticalSource() => Shapes(
            Shape(1, 5, 20, 5, 10),
            Shape(2, 30, 80, 10, 20),
            Shape(3, 60, 160, 20, 30));

        private static void AssertGap(IEnumerable<ShapeGeometry> shapes, int firstId, int secondId, float expected)
        {
            Assert.AreEqual(expected, Gap(shapes, firstId, secondId), Precision);
        }

        private static float Gap(IEnumerable<ShapeGeometry> shapes, int firstId, int secondId)
        {
            var first = ById(shapes, firstId);
            return ById(shapes, secondId).Left - first.Left - first.Width;
        }

        private static float VerticalGap(IEnumerable<ShapeGeometry> shapes, int firstId, int secondId)
        {
            var first = ById(shapes, firstId);
            return ById(shapes, secondId).Top - first.Top - first.Height;
        }

        private static float Centre(IEnumerable<ShapeGeometry> shapes, bool horizontal)
        {
            float minimum = shapes.Min(shape => horizontal ? shape.Left : shape.Top);
            float maximum = shapes.Max(shape => horizontal ? shape.Left + shape.Width : shape.Top + shape.Height);
            return (minimum + maximum) / 2f;
        }

        private static void AssertGeometry(IEnumerable<ShapeGeometry> actual, params ShapeGeometry[] expected)
        {
            foreach (var shape in expected)
            {
                var other = ById(actual, shape.Id);
                Assert.AreEqual(shape.Left, other.Left, Precision);
                Assert.AreEqual(shape.Top, other.Top, Precision);
                Assert.AreEqual(shape.Width, other.Width, Precision);
                Assert.AreEqual(shape.Height, other.Height, Precision);
            }
        }

        private static List<ShapeGeometry> Shapes(params ShapeGeometry[] shapes) => shapes.ToList();

        private static ShapeGeometry Shape(int id, float left, float top, float width, float height)
            => new ShapeGeometry { Id = id, Left = left, Top = top, Width = width, Height = height };

        private static ShapeGeometry ById(IEnumerable<ShapeGeometry> shapes, int id)
            => shapes.Single(shape => shape.Id == id);
    }
}
