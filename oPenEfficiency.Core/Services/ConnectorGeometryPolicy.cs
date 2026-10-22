using System;
using System.Collections.Generic;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free connector geometry for Object Connector (QA-S9 arrange wave): the nearest
    /// side pair between two shapes, the bounding-box points of a chosen side, the winding
    /// order of the connector quad and the analytic vertices of a regular polygon.
    /// Production implementation used by ObjectConnectorFeature.
    /// </summary>
    public static class ConnectorGeometryPolicy
    {
        /// <summary>
        /// The nearest side pair as (side of the first shape, side of the second shape).
        /// The horizontal axis wins whenever it dominates; an exact tie goes to the
        /// vertical axis, and each axis follows the sign of its delta.
        /// </summary>
        public static (string side1, string side2) DetectClosestSides(
            float fromCenterX, float fromCenterY, float toCenterX, float toCenterY)
        {
            float dx = toCenterX - fromCenterX;
            float dy = toCenterY - fromCenterY;

            if (Math.Abs(dx) > Math.Abs(dy))
            {
                return dx > 0 ? ("Right", "Left") : ("Left", "Right");
            }

            return dy > 0 ? ("Bottom", "Top") : ("Top", "Bottom");
        }

        /// <summary>The two corner points of the chosen side of a bounding box.</summary>
        public static List<(float x, float y)> GetBoundingBoxPoints(
            float left, float top, float width, float height, string side)
        {
            var pts = new List<(float, float)>();
            if (side == "Right")
            {
                pts.Add((left + width, top));
                pts.Add((left + width, top + height));
            }
            else if (side == "Left")
            {
                pts.Add((left, top));
                pts.Add((left, top + height));
            }
            else if (side == "Top")
            {
                pts.Add((left, top));
                pts.Add((left + width, top));
            }
            else if (side == "Bottom")
            {
                pts.Add((left, top + height));
                pts.Add((left + width, top + height));
            }
            return pts;
        }

        /// <summary>
        /// The connector quad in clockwise winding (first shape's side top → bottom, then
        /// second shape's side bottom → top), or an empty list when either side carries
        /// fewer than two points.
        /// </summary>
        public static List<(float x, float y)> BuildConnectorQuad(
            List<(float x, float y)> side1Points, List<(float x, float y)> side2Points)
        {
            if (side1Points.Count < 2 || side2Points.Count < 2)
                return new List<(float x, float y)>();

            return new List<(float x, float y)>
            {
                side1Points[0],                              // top of the first side
                side1Points[side1Points.Count - 1],          // bottom of the first side
                side2Points[side2Points.Count - 1],          // bottom of the second side
                side2Points[0]                               // top of the second side
            };
        }

        /// <summary>
        /// The vertices of a regular polygon laid over an ellipse (used for the
        /// analytically known AutoShape silhouettes).
        /// </summary>
        public static List<(float x, float y)> ComputeRegularPolygonVertices(
            float cx, float cy, float rx, float ry, int sides, double startAngle)
        {
            var pts = new List<(float, float)>();
            for (int i = 0; i < sides; i++)
            {
                double angle = startAngle + 2 * Math.PI * i / sides;
                pts.Add((cx + (float)(rx * Math.Cos(angle)),
                         cy + (float)(ry * Math.Sin(angle))));
            }
            return pts;
        }
    }
}
