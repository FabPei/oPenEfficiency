using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;
using static oPenEfficiency.Core.Services.ConnectorGeometryPolicy;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 arrange-wave coverage for Object Connector (QA9-BtnObjectConnector-002):
    /// the nearest side pair, the bounding-box side points, the connector quad winding and
    /// the analytic regular-polygon vertices used by ObjectConnectorFeature.
    /// </summary>
    [TestClass]
    public sealed class ConnectorGeometryPolicyTests
    {
        [TestMethod]
        public void TheDominantAxisPicksTheSidePair()
        {
            Assert.AreEqual(("Right", "Left"), DetectClosestSides(0f, 0f, 100f, 10f));
            Assert.AreEqual(("Left", "Right"), DetectClosestSides(100f, 10f, 0f, 0f));
            Assert.AreEqual(("Bottom", "Top"), DetectClosestSides(0f, 0f, 10f, 100f));
            Assert.AreEqual(("Top", "Bottom"), DetectClosestSides(0f, 0f, 10f, -100f));
        }

        [TestMethod]
        public void AnExactAxisTieGoesToTheVerticalPair()
        {
            Assert.AreEqual(("Top", "Bottom"), DetectClosestSides(0f, 0f, 5f, -5f));
            Assert.AreEqual(("Bottom", "Top"), DetectClosestSides(0f, 0f, -5f, 5f));
        }

        [TestMethod]
        public void EverySideYieldsItsTwoCorners()
        {
            // Box: left 10, top 20, 100x50.
            CollectionAssert.AreEqual(
                new List<(float, float)> { (110f, 20f), (110f, 70f) },
                GetBoundingBoxPoints(10f, 20f, 100f, 50f, "Right"));
            CollectionAssert.AreEqual(
                new List<(float, float)> { (10f, 20f), (10f, 70f) },
                GetBoundingBoxPoints(10f, 20f, 100f, 50f, "Left"));
            CollectionAssert.AreEqual(
                new List<(float, float)> { (10f, 20f), (110f, 20f) },
                GetBoundingBoxPoints(10f, 20f, 100f, 50f, "Top"));
            CollectionAssert.AreEqual(
                new List<(float, float)> { (10f, 70f), (110f, 70f) },
                GetBoundingBoxPoints(10f, 20f, 100f, 50f, "Bottom"));
            Assert.AreEqual(0, GetBoundingBoxPoints(10f, 20f, 100f, 50f, "Unknown").Count);
        }

        [TestMethod]
        public void TheQuadWindsClockwiseAndNeedsTwoPointsPerSide()
        {
            var quad = BuildConnectorQuad(
                new List<(float, float)> { (1f, 2f), (3f, 4f) },
                new List<(float, float)> { (5f, 6f), (7f, 8f) });

            CollectionAssert.AreEqual(
                new List<(float, float)> { (1f, 2f), (3f, 4f), (7f, 8f), (5f, 6f) }, quad);

            Assert.AreEqual(0, BuildConnectorQuad(
                new List<(float, float)> { (1f, 2f) },
                new List<(float, float)> { (5f, 6f), (7f, 8f) }).Count);
        }

        [TestMethod]
        public void ThePentagonStartsAtTheTopOfTheEllipse()
        {
            var pts = ComputeRegularPolygonVertices(100f, 100f, 50f, 40f, 5, -System.Math.PI / 2);

            Assert.AreEqual(5, pts.Count);
            Assert.AreEqual(100f, pts[0].x, 0.001);
            Assert.AreEqual(60f, pts[0].y, 0.001);
            Assert.AreEqual(147.55f, pts[1].x, 0.01);
            Assert.AreEqual(87.64f, pts[1].y, 0.01);
        }

        [TestMethod]
        public void TheHexagonAndOctagonUseTheirHistoricalStartAngles()
        {
            var hexagon = ComputeRegularPolygonVertices(100f, 100f, 50f, 40f, 6, 0);
            Assert.AreEqual(6, hexagon.Count);
            Assert.AreEqual(150f, hexagon[0].x, 0.001);
            Assert.AreEqual(100f, hexagon[0].y, 0.001);

            var octagon = ComputeRegularPolygonVertices(100f, 100f, 50f, 40f, 8, System.Math.PI / 8);
            Assert.AreEqual(8, octagon.Count);
            Assert.AreEqual(146.194f, octagon[0].x, 0.01);
            Assert.AreEqual(115.307f, octagon[0].y, 0.01);
        }
    }
}
