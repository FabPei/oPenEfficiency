using System;
using System.Linq;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free diagonal placement for Align Diagonally (QA-S9 align-and-size wave):
    /// projecting the selected shapes' original centers onto the requested direction,
    /// ordering them along it, deriving the even spacing (with the historical degenerate
    /// fallback) and computing every shape's new top-left around the chosen anchor.
    /// Production implementation used by DiagonalAlignFeature (derived spacing) and
    /// DiagonalAlignToolbar (the user's spacing from the toolbar slider).
    /// </summary>
    public static class DiagonalPlacementPolicy
    {
        /// <summary>The shape that stays put: the first / last projected shape, or the selection center.</summary>
        public enum DiagonalAnchor
        {
            First,
            Center,
            Last
        }

        /// <summary>
        /// The historical anchor mapping: exactly "First" and "Last" select those shapes;
        /// every other value (including null) anchors on the selection center.
        /// </summary>
        public static DiagonalAnchor ResolveAnchor(string anchorType)
        {
            if (anchorType == "First") return DiagonalAnchor.First;
            if (anchorType == "Last") return DiagonalAnchor.Last;
            return DiagonalAnchor.Center;
        }

        /// <summary>The unit direction of the diagonal (degrees, PowerPoint's downward-positive Y).</summary>
        public static void GetDirection(float angleDeg, out float dirX, out float dirY)
        {
            float angleRad = (float)(angleDeg * Math.PI / 180f);
            dirX = (float)Math.Cos(angleRad);
            dirY = (float)Math.Sin(angleRad);
        }

        /// <summary>
        /// The new top-left of every shape, returned in selection order. The shapes are
        /// ordered by the projection of their original center onto the direction; the
        /// spacing is the projected span divided by (count - 1), falling back to the
        /// average extent along the direction when that spacing collapses below 0.01 pt
        /// (shapes stacked on top of each other). The anchor shape keeps its center.
        /// </summary>
        public static void ComputePlacements(
            float[] centerX, float[] centerY, float[] width, float[] height,
            float angleDeg, DiagonalAnchor anchor,
            out float[] newLeft, out float[] newTop)
        {
            GetDirection(angleDeg, out float dirX, out float dirY);

            var projected = BuildProjection(centerX, centerY, dirX, dirY, out float avgX, out float avgY);

            float projectedSpan = projected[projected.Length - 1].Proj - projected[0].Proj;
            float spacing = projectedSpan / (projected.Length - 1);
            if (spacing < 0.01f)
            {
                spacing = (float)Enumerable.Range(0, centerX.Length)
                    .Average(i => Math.Abs(dirX) * width[i] + Math.Abs(dirY) * height[i]);
            }

            PlaceAlongDirection(projected, centerX, centerY, width, height, dirX, dirY, spacing, anchor, avgX, avgY,
                out newLeft, out newTop);
        }

        /// <summary>
        /// The same placement with the spacing the user set (the Align Diagonally toolbar
        /// slider): projection, ordering and anchor behave exactly as in the command path.
        /// </summary>
        public static void ComputePlacementsWithSpacing(
            float[] centerX, float[] centerY, float[] width, float[] height,
            float angleDeg, float spacing, DiagonalAnchor anchor,
            out float[] newLeft, out float[] newTop)
        {
            GetDirection(angleDeg, out float dirX, out float dirY);

            var projected = BuildProjection(centerX, centerY, dirX, dirY, out float avgX, out float avgY);

            PlaceAlongDirection(projected, centerX, centerY, width, height, dirX, dirY, spacing, anchor, avgX, avgY,
                out newLeft, out newTop);
        }

        private struct ProjectedShape
        {
            public int Index;
            public float Proj;
        }

        /// <summary>
        /// The shapes ordered by the projection of their center onto the direction.
        /// Enumerable.Average over floats accumulates in double and the historical code
        /// assigned that result to a float, so the same widening cast is kept.
        /// The OrderBy is stable: equal projections keep selection order.
        /// </summary>
        private static ProjectedShape[] BuildProjection(
            float[] centerX, float[] centerY, float dirX, float dirY,
            out float avgX, out float avgY)
        {
            // Enumerable.Average over floats accumulates in double; the historical code
            // assigned that result to a float, so the same widening cast is kept here.
            float localAvgX = (float)centerX.Average();
            float localAvgY = (float)centerY.Average();

            var projected = Enumerable.Range(0, centerX.Length)
                .Select(i => new ProjectedShape
                {
                    Index = i,
                    Proj = (centerX[i] - localAvgX) * dirX + (centerY[i] - localAvgY) * dirY
                })
                .OrderBy(x => x.Proj)   // stable: equal projections keep selection order
                .ToArray();

            avgX = localAvgX;
            avgY = localAvgY;
            return projected;
        }

        /// <summary>
        /// Every shape's new top-left around the anchor (the first/last projected shape or
        /// the selection center, which keeps the middle of the ordered line in place).
        /// </summary>
        private static void PlaceAlongDirection(
            ProjectedShape[] projected,
            float[] centerX, float[] centerY, float[] width, float[] height,
            float dirX, float dirY, float spacing, DiagonalAnchor anchor, float avgX, float avgY,
            out float[] newLeft, out float[] newTop)
        {
            int count = projected.Length;
            newLeft = new float[count];
            newTop = new float[count];

            float anchorOrigX = avgX;
            float anchorOrigY = avgY;
            float anchorIndex = (count - 1) / 2f;

            if (anchor == DiagonalAnchor.First)
            {
                anchorOrigX = centerX[projected[0].Index];
                anchorOrigY = centerY[projected[0].Index];
                anchorIndex = 0;
            }
            else if (anchor == DiagonalAnchor.Last)
            {
                anchorOrigX = centerX[projected[count - 1].Index];
                anchorOrigY = centerY[projected[count - 1].Index];
                anchorIndex = count - 1;
            }

            for (int i = 0; i < count; i++)
            {
                int index = projected[i].Index;
                float offset = (i - anchorIndex) * spacing;

                float newCenterX = anchorOrigX + offset * dirX;
                float newCenterY = anchorOrigY + offset * dirY;

                newLeft[index] = newCenterX - (width[index] / 2f);
                newTop[index] = newCenterY - (height[index] / 2f);
            }
        }
    }
}
