using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 arrange-wave coverage for Optimize FreeForm (QA9-BtnOptimizeFreeForm-002): the
    /// vertex decimation rule — which nodes are deleted when adjacent points sit closer than
    /// the minimum distance — used by OptimizeFreeFormFeature.
    /// </summary>
    [TestClass]
    public sealed class FreeformDecimationPolicyTests
    {
        private static List<(float X, float Y)> Points(params (float X, float Y)[] coordinates)
        {
            return new List<(float X, float Y)>(coordinates);
        }

        [TestMethod]
        public void CloseNodesAreDeletedInDescendingOrder()
        {
            // Nodes (1-based): 1=(0,0), 2=(1,0) close to 1, 3=(50,0), 4=(51,0) close to 3,
            // 5=(100,0). Nodes 2 and 4 sit below the 2-point minimum; node 3 does not.
            var deletions = FreeformDecimationPolicy.ComputeDeletableNodeIndices(
                Points((0f, 0f), (1f, 0f), (50f, 0f), (51f, 0f), (100f, 0f)), 2f);
            CollectionAssert.AreEqual(new[] { 4, 2 }, deletions); // descending
        }

        [TestMethod]
        public void ThresholdIsStrictAndComparesOriginalNeighbors()
        {
            // Exactly at the minimum distance is kept (strict <); a chain of close nodes
            // pairs each node against its ORIGINAL predecessor, not a decimated one.
            var deletions = FreeformDecimationPolicy.ComputeDeletableNodeIndices(
                Points((0f, 0f), (2f, 0f), (4f, 0f), (6f, 0f), (100f, 0f)), 2f);
            // All three inter-node distances are exactly 2 -> none deleted.
            Assert.AreEqual(0, deletions.Count);

            var chain = FreeformDecimationPolicy.ComputeDeletableNodeIndices(
                Points((0f, 0f), (1f, 0f), (2f, 0f), (3f, 0f), (100f, 0f)), 1.5f);
            // Original pairs (2,1), (3,2), (4,3) are all distance 1 -> nodes 2,3,4 deleted.
            CollectionAssert.AreEqual(new[] { 4, 3, 2 }, chain);
        }

        [TestMethod]
        public void EndpointsAreNeverDeleted()
        {
            var deletions = FreeformDecimationPolicy.ComputeDeletableNodeIndices(
                Points((0f, 0f), (0.5f, 0f), (1f, 0f), (100f, 0f)), 5f);
            CollectionAssert.AreEqual(new[] { 3, 2 }, deletions); // nodes 1 and 4 survive
        }

        [TestMethod]
        public void FreeformsWithThreeOrFewerNodesAreNeverOptimized()
        {
            Assert.AreEqual(0, FreeformDecimationPolicy.ComputeDeletableNodeIndices(
                Points((0f, 0f), (0.1f, 0f), (0.2f, 0f)), 10f).Count);
            Assert.AreEqual(0, FreeformDecimationPolicy.ComputeDeletableNodeIndices(
                Points((0f, 0f), (0.1f, 0f)), 10f).Count);
        }

        [TestMethod]
        public void TwoDimensionalDistancesAreConsidered()
        {
            // Node 2 sits sqrt(1^2+1^2) from node 1 -> below a sqrt(2.5) minimum? No:
            // distSq = 2 < 2.5 -> deleted. Node 3 sits distSq = 8 from node 2 -> kept.
            var deletions = FreeformDecimationPolicy.ComputeDeletableNodeIndices(
                Points((0f, 0f), (1f, 1f), (3f, 3f), (100f, 100f)), 2.5f);
            CollectionAssert.AreEqual(new[] { 2 }, deletions);
        }
    }
}
