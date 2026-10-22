using System;
using System.Collections.Generic;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free vertex decimation for Optimize FreeForm (QA-S9 arrange wave): which
    /// freeform nodes are deleted when adjacent points sit closer than the minimum
    /// distance. Production implementation used by OptimizeFreeFormFeature.
    /// </summary>
    public static class FreeformDecimationPolicy
    {
        /// <summary>
        /// Computes the 1-based node indices to delete: walking from the second-to-last
        /// node down to the third, a node is deleted when its distance to its immediate
        /// predecessor is below the minimum distance. Because deletions happen at strictly
        /// descending indices, every comparison pairs original adjacent points and node 1
        /// and the last node are never deleted. Freeforms with three or fewer nodes are
        /// never optimized. The returned indices are descending, so the caller can delete
        /// them in order against the live collection.
        /// </summary>
        public static List<int> ComputeDeletableNodeIndices(
            IReadOnlyList<(float X, float Y)> points,
            float minimumDistance)
        {
            var deletions = new List<int>();
            int count = points.Count;
            if (count <= 3) return deletions;

            float minimumDistanceSq = minimumDistance * minimumDistance;
            for (int i = count - 1; i >= 2; i--)
            {
                float dx = points[i - 1].X - points[i - 2].X;
                float dy = points[i - 1].Y - points[i - 2].Y;
                if (dx * dx + dy * dy < minimumDistanceSq) deletions.Add(i);
            }
            return deletions;
        }
    }
}
