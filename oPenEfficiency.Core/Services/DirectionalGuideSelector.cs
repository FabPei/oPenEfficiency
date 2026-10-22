using System;
using System.Collections.Generic;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Selects a guide on the requested side of a moving edge.
    /// Kept host-neutral so directional alignment can be regression-tested without PowerPoint.
    /// </summary>
    public static class DirectionalGuideSelector
    {
        public static float NearestAtOrBefore(IEnumerable<float> guides, float movingEdge, float fallback)
        {
            if (guides == null) throw new ArgumentNullException(nameof(guides));

            float best = fallback;
            float bestDistance = fallback <= movingEdge ? movingEdge - fallback : float.MaxValue;
            foreach (float guide in guides)
            {
                if (guide > movingEdge) continue;
                float distance = movingEdge - guide;
                if (distance < bestDistance)
                {
                    bestDistance = distance;
                    best = guide;
                }
            }
            return best;
        }

        public static float NearestAtOrAfter(IEnumerable<float> guides, float movingEdge, float fallback)
        {
            if (guides == null) throw new ArgumentNullException(nameof(guides));

            float best = fallback;
            float bestDistance = fallback >= movingEdge ? fallback - movingEdge : float.MaxValue;
            foreach (float guide in guides)
            {
                if (guide < movingEdge) continue;
                float distance = guide - movingEdge;
                if (distance < bestDistance)
                {
                    bestDistance = distance;
                    best = guide;
                }
            }
            return best;
        }

        /// <summary>
        /// Selects the guide closest to a moving center by absolute distance. The fallback (the slide's
        /// center line) participates last, so a real guide at equal distance wins - the iteration order of
        /// the COM route this replaces. Kept host-neutral so center alignment can be regression-tested
        /// without PowerPoint.
        /// </summary>
        public static float NearestByDistance(IEnumerable<float> guides, float center, float fallback)
        {
            if (guides == null) throw new ArgumentNullException(nameof(guides));

            float best = fallback;
            float bestDistance = float.MaxValue;
            foreach (float guide in guides)
            {
                float distance = Math.Abs(guide - center);
                if (distance < bestDistance)
                {
                    bestDistance = distance;
                    best = guide;
                }
            }
            float fallbackDistance = Math.Abs(fallback - center);
            if (fallbackDistance < bestDistance)
            {
                best = fallback;
            }
            return best;
        }
    }
}
