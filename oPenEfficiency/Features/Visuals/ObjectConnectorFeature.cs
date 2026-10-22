using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnObjectConnector",
        Name = "Connect Shapes",
        SearchAliases = new[] { "Object Connector" },
        Tooltip = "Create a filled connector between two selected shapes",
        IconData = "M1,6H9V18H1Z M3,8V16H7V8Z M15,6H23V18H15Z M17,8V16H21V8Z M9,9H15V11H9Z M9,13H15V15H9Z",
        Color = "#06B6D4",
        Description = "Creates a filled connector shape between two selected objects. Auto-detects closest sides.",
        DetailedHelpText = "### Connect Shapes\nCreates a filled connector polygon between the first two selected shapes. The primary action detects the nearest sides and follows supported shape silhouettes. Right-click to use automatic smart routing or force right-to-left, left-to-right, bottom-to-top, or top-to-bottom attachment sides.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class ObjectConnectorFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - creates connector with auto-detected sides.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, "Right", "Left", autoDetect: true, smart: true);
        }

        public static bool Execute(PowerPointManager manager, string side1 = "Right", string side2 = "Left", bool autoDetect = false, bool smart = false)
        {
            try
            {
                var shapes = manager.GetSelectedShapes();
                if (shapes == null || shapes.Count < 2) return false;

                var shape1 = shapes[1];
                var shape2 = shapes[2];

                if (autoDetect)
                {
                    (side1, side2) = DetectClosestSides(shape1, shape2);
                }

                var pts = CalculatePoints(shape1, side1, shape2, side2, smart);
                if (pts.Count < 4) return false;

                var parentSlide = (Slide)((dynamic)shape1).Parent;
                var builder = parentSlide.Shapes.BuildFreeform(Office.MsoEditingType.msoEditingCorner, pts[0].x, pts[0].y);
                
                for (int i = 1; i < pts.Count; i++)
                {
                    builder.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingAuto, pts[i].x, pts[i].y);
                }
                builder.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingAuto, pts[0].x, pts[0].y);

                var newShape = builder.ConvertToShape();

                try
                {
                    newShape.Fill.ForeColor.RGB = shape1.Fill.ForeColor.RGB;
                    newShape.Fill.Transparency = 0.5f;
                    newShape.Line.Visible = Office.MsoTriState.msoFalse;
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"ObjectConnector: Error styling shape: {ex.Message}");
                }

                return true;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ObjectConnector Error: {ex}");
                return false;
            }
        }

        private static (string, string) DetectClosestSides(Shape s1, Shape s2)
        {
            return oPenEfficiency.Core.Services.ConnectorGeometryPolicy.DetectClosestSides(
                s1.Left + s1.Width / 2, s1.Top + s1.Height / 2,
                s2.Left + s2.Width / 2, s2.Top + s2.Height / 2);
        }

        private static List<(float x, float y)> CalculatePoints(Shape s1, string side1, Shape s2, string side2, bool smart)
        {
            // p1: two points on s1's chosen side, sorted top→bottom (or left→right)
            // p2: two points on s2's chosen side, sorted top→bottom (or left→right)
            List<(float x, float y)> p1 = smart ? GetSilhouettePoints(s1, side1) : GetBoundingBoxPoints(s1, side1);
            List<(float x, float y)> p2 = smart ? GetSilhouettePoints(s2, side2) : GetBoundingBoxPoints(s2, side2);

            // Polygon winds clockwise: p1-top → p1-bottom → p2-bottom → p2-top
            // This ensures a proper closed quad without crossed edges.
            return oPenEfficiency.Core.Services.ConnectorGeometryPolicy.BuildConnectorQuad(p1, p2);
        }

        private static List<(float x, float y)> GetBoundingBoxPoints(Shape s, string side)
        {
            return oPenEfficiency.Core.Services.ConnectorGeometryPolicy.GetBoundingBoxPoints(
                s.Left, s.Top, s.Width, s.Height, side);
        }

        private static List<(float x, float y)> GetSilhouettePoints(Shape s, string side)
        {
            try
            {
                var vertices = GetShapeVertices(s);
                if (vertices == null || vertices.Count < 2)
                    return GetBoundingBoxPoints(s, side);

                float centerX = s.Left + s.Width / 2;
                float centerY = s.Top + s.Height / 2;

                List<(float x, float y)> candidates;
                if      (side == "Right")  candidates = vertices.FindAll(n => n.x >= centerX - 1f);
                else if (side == "Left")   candidates = vertices.FindAll(n => n.x <= centerX + 1f);
                else if (side == "Top")    candidates = vertices.FindAll(n => n.y <= centerY + 1f);
                else                       candidates = vertices.FindAll(n => n.y >= centerY - 1f);

                if (candidates.Count < 2) candidates = vertices;

                (float x, float y) edgeA, edgeB;
                if (side == "Right" || side == "Left")
                {
                    candidates.Sort((a, b) => a.y.CompareTo(b.y));
                    edgeA = candidates[0];
                    edgeB = candidates[candidates.Count - 1];
                }
                else
                {
                    candidates.Sort((a, b) => a.x.CompareTo(b.x));
                    edgeA = candidates[0];
                    edgeB = candidates[candidates.Count - 1];
                }

                if (System.Math.Abs(edgeA.x - edgeB.x) < 0.5f && System.Math.Abs(edgeA.y - edgeB.y) < 0.5f)
                    return GetBoundingBoxPoints(s, side);

                return new List<(float x, float y)> { edgeA, edgeB };
            }
            catch
            {
                return GetBoundingBoxPoints(s, side);
            }
        }

        /// <summary>
        /// Returns the key vertices of a shape in slide-absolute coordinates.
        /// Handles known AutoShapes analytically, freeforms via node API, others return null (→ bounding box fallback).
        /// </summary>
        private static List<(float x, float y)> GetShapeVertices(Shape s)
        {
            float l = s.Left, t = s.Top;
            float r = s.Left + s.Width, b = s.Top + s.Height;
            float cx = l + s.Width / 2f, cy = t + s.Height / 2f;

            // --- Known AutoShape types: compute analytically ---
            try
            {
                switch (s.AutoShapeType)
                {
                    case Office.MsoAutoShapeType.msoShapeIsoscelesTriangle:
                        return new List<(float, float)> { (cx, t), (l, b), (r, b) };

                    case Office.MsoAutoShapeType.msoShapeRightTriangle:
                        return new List<(float, float)> { (l, t), (l, b), (r, b) };

                    case Office.MsoAutoShapeType.msoShapeDiamond:
                        return new List<(float, float)> { (cx, t), (r, cy), (cx, b), (l, cy) };

                    case Office.MsoAutoShapeType.msoShapeParallelogram:
                        // Parallelogram: top-right and bottom-left are offset by ~25% of width
                        float offset = s.Width * 0.25f;
                        return new List<(float, float)> { (l + offset, t), (r, t), (r - offset, b), (l, b) };

                    case Office.MsoAutoShapeType.msoShapeTrapezoid:
                        float trap = s.Width * 0.2f;
                        return new List<(float, float)> { (l + trap, t), (r - trap, t), (r, b), (l, b) };

                    case Office.MsoAutoShapeType.msoShapeRegularPentagon:
                        // Approximate: use 5 equally spaced points on an ellipse
                        return oPenEfficiency.Core.Services.ConnectorGeometryPolicy.ComputeRegularPolygonVertices(cx, cy, s.Width / 2f, s.Height / 2f, 5, -System.Math.PI / 2);

                    case Office.MsoAutoShapeType.msoShapeHexagon:
                        return oPenEfficiency.Core.Services.ConnectorGeometryPolicy.ComputeRegularPolygonVertices(cx, cy, s.Width / 2f, s.Height / 2f, 6, 0);

                    case Office.MsoAutoShapeType.msoShapeOctagon:
                        return oPenEfficiency.Core.Services.ConnectorGeometryPolicy.ComputeRegularPolygonVertices(cx, cy, s.Width / 2f, s.Height / 2f, 8, System.Math.PI / 8);
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ObjectConnector: Error getting shape vertices: {ex.Message}");
            }

            // --- Actual freeform: read nodes directly (no conversion needed) ---
            if (s.Type == Office.MsoShapeType.msoFreeform)
            {
                try
                {
                    var nodes = new List<(float x, float y)>();
                    for (int i = 1; i <= s.Nodes.Count; i++)
                    {
                        dynamic p = s.Nodes[i].Points;
                        float x = p[1, 1], y = p[1, 2];
                        bool dup = false;
                        foreach (var n in nodes)
                            if (System.Math.Abs(n.x - x) < 2f && System.Math.Abs(n.y - y) < 2f) { dup = true; break; }
                        if (!dup) nodes.Add((x, y));
                    }
                    if (nodes.Count >= 2) return nodes;
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"ObjectConnector: Error reading freeform nodes: {ex.Message}");
                }
            }

            return null; // triggers bounding-box fallback
        }
    }
}
