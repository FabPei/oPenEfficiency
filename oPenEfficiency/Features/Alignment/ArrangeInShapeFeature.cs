using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    public enum ArrangeProAnchorMode
    {
        SelectionCenter,
        FirstShape,
        LastShape
    }

    public enum ArrangeProOrder
    {
        Selection,
        LeftToRight,
        TopToBottom
    }

    /// <summary>
    /// Feature: Arrange selected shapes in various shapes (Circle, Triangle, Heart, etc.).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnArrangeInShape",
        Name = "Arrange Shapes in Pattern",
        SearchAliases = new[] { "Arrange in Shape" },
        Tooltip = "Arrange shapes in a geometric pattern",
        IconData = "M12,2C6.48,2 2,6.48 2,12C2,17.52 6.48,22 12,22C17.52,22 22,17.52 22,12C22,6.48 17.52,2 12,2M12,4C16.42,4 20,7.58 20,12C20,16.42 16.42,20 12,20C7.58,20 4,16.42 4,12C4,7.58 7.58,4 12,4Z",
        Color = "#F43F5E",
        Description = "Arranges selected shapes in a circular, triangular, or other pattern.",
        DetailedHelpText = "### Arrange Shapes in Pattern\nDistributes selected shapes into a geometric pattern within their total bounding area.\n\n**Right-Click Options:**\n* **Patterns**: Choose between Circle, Triangle, Square, Star, Octagon, or Pentagon.\n* **Pro Mode**: Opens a dialog for fine-tuned control over radius, start angle, and auto-rotation towards center.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class ArrangeInShapeFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - arranges shapes in a circle pattern.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, shapeType: "circle");
        }

        public static bool Execute(PowerPointManager manager, string shapeType)
        {
            if (manager == null) return false;
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count < 2) return false;

            float minL = float.MaxValue, minT = float.MaxValue, maxR = float.MinValue, maxB = float.MinValue;
            for (int i = 1; i <= shapeRange.Count; i++)
            {
                var s = shapeRange[i];
                if (s.Left < minL) minL = s.Left;
                if (s.Top < minT) minT = s.Top;
                if (s.Left + s.Width > maxR) maxR = s.Left + s.Width;
                if (s.Top + s.Height > maxB) maxB = s.Top + s.Height;
            }
            float radius = (Math.Max(maxR - minL, maxB - minT)) / 2f;
            if (radius < 10) radius = 100;

            return ArrangeInShapePro(manager, shapeType, shapeRange.Count, radius, 0, false, false, false);
        }

        public static bool ArrangeInShapePro(
            PowerPointManager manager,
            string shapeType,
            int instances,
            float radius,
            float startAngleDegree,
            bool centerOnSlide,
            bool rotateToCenter,
            bool fillShape = false,
            float? anchorX = null,
            float? anchorY = null,
            ArrangeProAnchorMode anchorMode = ArrangeProAnchorMode.SelectionCenter,
            ArrangeProOrder order = ArrangeProOrder.Selection,
            bool clockwise = true,
            float arcSpanDegrees = 270f,
            bool avoidOverlap = false,
            IDictionary<int, Tuple<float, float>> originalOrderPositions = null)
        {
            if (manager == null) return false;
            var originalStates = new List<Tuple<Shape, float, float, float>>();
            var createdInstances = new List<Shape>();
            try
            {
                var app = manager.GetApplication();
                var selection = manager.GetSelectedShapes();
                if (app == null || selection == null || selection.Count == 0 || instances < 1 || radius <= 0 || float.IsNaN(radius) || float.IsInfinity(radius)) return false;

                string normalizedShapeType = (shapeType ?? string.Empty).Trim().ToLowerInvariant();
                if (normalizedShapeType != "circle" && normalizedShapeType != "arc" &&
                    normalizedShapeType != "triangle" && normalizedShapeType != "square" &&
                    normalizedShapeType != "pentagon" && normalizedShapeType != "octagon" &&
                    normalizedShapeType != "star" && normalizedShapeType != "heart") return false;
                if (normalizedShapeType == "arc" &&
                    (arcSpanDegrees <= 0 || arcSpanDegrees > 360 ||
                     float.IsNaN(arcSpanDegrees) || float.IsInfinity(arcSpanDegrees))) return false;

                try { app.StartNewUndoEntry(); } catch { /* Arrangement remains valid when undo grouping is unavailable. */ }

                List<Shape> shapesToArrange = new List<Shape>();
                Shape anchorShape = null;
                const string InstanceTag = "oPE_ArrangePro_Instance";
                const string SourceTag = "oPE_ArrangePro_Source";
                var instancesToDelete = new List<Shape>();

                if (selection.Count == 1)
                {
                    var baseShape = selection[1];
                    originalStates.Add(Tuple.Create(baseShape, baseShape.Left, baseShape.Top, baseShape.Rotation));
                    bool useBaseAsAnchor = anchorMode == ArrangeProAnchorMode.FirstShape ||
                                           anchorMode == ArrangeProAnchorMode.LastShape;
                    if (useBaseAsAnchor)
                        anchorShape = baseShape;
                    else
                        shapesToArrange.Add(baseShape);

                    var slide = (Slide)baseShape.Parent;
                    var existingInstances = new List<Shape>();
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        var s = slide.Shapes[i];
                        if (s.Tags[InstanceTag] == "true" && s.Tags[SourceTag] == baseShape.Id.ToString())
                        {
                            existingInstances.Add(s);
                            originalStates.Add(Tuple.Create(s, s.Left, s.Top, s.Rotation));
                        }
                    }

                    int needed = useBaseAsAnchor ? instances : instances - 1;
                    while (existingInstances.Count > needed && existingInstances.Count > 0)
                    {
                        instancesToDelete.Add(existingInstances[existingInstances.Count - 1]);
                        existingInstances.RemoveAt(existingInstances.Count - 1);
                    }

                    while (existingInstances.Count < needed)
                    {
                        var dup = baseShape.Duplicate();
                        var newShape = dup[1];
                        newShape.Tags.Add(InstanceTag, "true");
                        newShape.Tags.Add(SourceTag, baseShape.Id.ToString());
                        existingInstances.Add(newShape);
                        createdInstances.Add(newShape);
                    }

                    shapesToArrange.AddRange(existingInstances);
                }
                else
                {
                    for (int i = 1; i <= selection.Count; i++)
                    {
                        var selectedShape = selection[i];
                        originalStates.Add(Tuple.Create(selectedShape, selectedShape.Left, selectedShape.Top, selectedShape.Rotation));
                        if ((anchorMode == ArrangeProAnchorMode.FirstShape && i == 1) ||
                            (anchorMode == ArrangeProAnchorMode.LastShape && i == selection.Count))
                            anchorShape = selectedShape;
                        else
                            shapesToArrange.Add(selectedShape);
                    }
                }

                SortShapes(shapesToArrange, order, originalOrderPositions);
                int count = shapesToArrange.Count;
                if (count == 0) return false;
                float centerX, centerY;

                if (anchorShape != null)
                {
                    centerX = anchorShape.Left + anchorShape.Width / 2f;
                    centerY = anchorShape.Top + anchorShape.Height / 2f;
                }
                else if (anchorX.HasValue && anchorY.HasValue)
                {
                    centerX = anchorX.Value;
                    centerY = anchorY.Value;
                }
                else if (centerOnSlide)
                {
                    centerX = app.ActivePresentation.PageSetup.SlideWidth / 2f;
                    centerY = app.ActivePresentation.PageSetup.SlideHeight / 2f;
                }
                else
                {
                    float minL = float.MaxValue, minT = float.MaxValue, maxR = float.MinValue, maxB = float.MinValue;
                    foreach (var s in shapesToArrange)
                    {
                        if (s.Left < minL) minL = s.Left;
                        if (s.Top < minT) minT = s.Top;
                        if (s.Left + s.Width > maxR) maxR = s.Left + s.Width;
                        if (s.Top + s.Height > maxB) maxB = s.Top + s.Height;
                    }
                    centerX = (minL + maxR) / 2f;
                    centerY = (minT + maxB) / 2f;
                }

                float effectiveRadius = radius;
                List<Tuple<float, float>> targets = null;
                int attempts = avoidOverlap ? 100 : 1;
                for (int attempt = 0; attempt < attempts; attempt++)
                {
                    targets = PatternPlacementPlanner.BuildTargets(
                        normalizedShapeType,
                        count,
                        effectiveRadius,
                        centerX,
                        centerY,
                        startAngleDegree,
                        fillShape,
                        clockwise,
                        arcSpanDegrees,
                        anchorShape != null);

                    if (!avoidOverlap || !HasOverlap(shapesToArrange, targets, anchorShape, 4f))
                        break;

                    effectiveRadius *= 1.08f;
                    targets = null;
                }

                if (targets == null)
                    throw new InvalidOperationException("Arrange Shapes Geometrically could not find a collision-free radius.");

                for (int i = 0; i < count; i++)
                {
                    var shape = shapesToArrange[i];
                    float targetX = targets[i].Item1;
                    float targetY = targets[i].Item2;
                    shape.Left = targetX - shape.Width / 2f;
                    shape.Top = targetY - shape.Height / 2f;

                    if (rotateToCenter)
                    {
                        double rotation = Math.Atan2(targetY - centerY, targetX - centerX);
                        shape.Rotation = (float)(rotation * 180.0 / Math.PI) + 90f;
                    }
                }

                foreach (var extraInstance in instancesToDelete) extraInstance.Delete();
                return true;
            }
            catch (Exception ex)
            {
                foreach (var state in originalStates)
                {
                    try { state.Item1.Left = state.Item2; state.Item1.Top = state.Item3; state.Item1.Rotation = state.Item4; } catch { /* Preserve the primary error if a stale shape cannot be restored. */ }
                }
                foreach (var created in createdInstances)
                {
                    try { created.Delete(); } catch { /* Rollback shape may already have been removed by PowerPoint. */ }
                }
                ExceptionLogger.Log(ex, "ArrangeInShapeFeature.ArrangeInShapePro");
                return false;
            }
        }

        private static void SortShapes(
            List<Shape> shapes,
            ArrangeProOrder order,
            IDictionary<int, Tuple<float, float>> originalOrderPositions)
        {
            if (order == ArrangeProOrder.LeftToRight)
            {
                shapes.Sort((left, right) =>
                {
                    var leftPosition = GetOrderPosition(left, originalOrderPositions);
                    var rightPosition = GetOrderPosition(right, originalOrderPositions);
                    int horizontal = leftPosition.Item1.CompareTo(rightPosition.Item1);
                    return horizontal != 0
                        ? horizontal
                        : leftPosition.Item2.CompareTo(rightPosition.Item2);
                });
            }
            else if (order == ArrangeProOrder.TopToBottom)
            {
                shapes.Sort((left, right) =>
                {
                    var leftPosition = GetOrderPosition(left, originalOrderPositions);
                    var rightPosition = GetOrderPosition(right, originalOrderPositions);
                    int vertical = leftPosition.Item2.CompareTo(rightPosition.Item2);
                    return vertical != 0
                        ? vertical
                        : leftPosition.Item1.CompareTo(rightPosition.Item1);
                });
            }
        }

        private static Tuple<float, float> GetOrderPosition(
            Shape shape,
            IDictionary<int, Tuple<float, float>> originalOrderPositions)
        {
            Tuple<float, float> position;
            if (originalOrderPositions != null &&
                originalOrderPositions.TryGetValue(shape.Id, out position))
                return position;

            return Tuple.Create(shape.Left + shape.Width / 2f, shape.Top + shape.Height / 2f);
        }

        private static bool HasOverlap(
            IList<Shape> shapes,
            IList<Tuple<float, float>> targets,
            Shape anchorShape,
            float gap)
        {
            for (int i = 0; i < shapes.Count; i++)
            {
                if (anchorShape != null &&
                    RectanglesOverlap(
                        targets[i].Item1,
                        targets[i].Item2,
                        shapes[i].Width,
                        shapes[i].Height,
                        anchorShape.Left + anchorShape.Width / 2f,
                        anchorShape.Top + anchorShape.Height / 2f,
                        anchorShape.Width,
                        anchorShape.Height,
                        gap))
                    return true;

                for (int j = i + 1; j < shapes.Count; j++)
                {
                    if (RectanglesOverlap(
                        targets[i].Item1,
                        targets[i].Item2,
                        shapes[i].Width,
                        shapes[i].Height,
                        targets[j].Item1,
                        targets[j].Item2,
                        shapes[j].Width,
                        shapes[j].Height,
                        gap))
                        return true;
                }
            }

            return false;
        }

        private static bool RectanglesOverlap(
            float firstX,
            float firstY,
            float firstWidth,
            float firstHeight,
            float secondX,
            float secondY,
            float secondWidth,
            float secondHeight,
            float gap)
        {
            return Math.Abs(firstX - secondX) < (firstWidth + secondWidth) / 2f + gap &&
                   Math.Abs(firstY - secondY) < (firstHeight + secondHeight) / 2f + gap;
        }

    }
}
