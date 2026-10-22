using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Internal helper providing alignment, match-size, and dock logic used by
    /// AlignToFirst*, AlignShapes, MatchSize, etc. feature classes.
    /// </summary>
    internal static class AlignmentHelpers
    {
        private const string GuideH = "Guide_H_";
        private const string GuideV = "Guide_V_";
        private const string OpenEfficiencyGuide = "oPE_Guide_";
        private const float GuideTolerance    = 0.5f; // snapping window for guide detection
        private const float CollisionTolerance = 1f;  // snapping window for collision detection

        // ──────────────────────────────────────────────────────
        // AlignShapes — standard PowerPoint alignment directions
        // If only one shape is selected, align to slide center/edges
        // ──────────────────────────────────────────────────────
        public static bool AlignShapes(PowerPointManager manager, string direction)
        {
            if (manager == null) return false;
            var app = manager.GetApplication();
            if (app.ActiveWindow?.Selection?.Type != PpSelectionType.ppSelectionShapes) return false;
            try
            {
                var sr = app.ActiveWindow.Selection.ShapeRange;
                if (sr == null || sr.Count == 0) return false;

                // If only one shape selected, align to slide
                if (sr.Count == 1)
                {
                    return AlignShapeToSlide(app, sr[1], direction);
                }

                // Multiple shapes - use standard PowerPoint alignment
                Microsoft.Office.Core.MsoAlignCmd cmd;
                switch (direction.ToLower())
                {
                    case "left":   cmd = Microsoft.Office.Core.MsoAlignCmd.msoAlignLefts;   break;
                    case "right":  cmd = Microsoft.Office.Core.MsoAlignCmd.msoAlignRights;  break;
                    case "top":    cmd = Microsoft.Office.Core.MsoAlignCmd.msoAlignTops;    break;
                    case "bottom": cmd = Microsoft.Office.Core.MsoAlignCmd.msoAlignBottoms; break;
                    case "center": cmd = Microsoft.Office.Core.MsoAlignCmd.msoAlignCenters; break;
                    case "middle": cmd = Microsoft.Office.Core.MsoAlignCmd.msoAlignMiddles; break;
                    default: return false;
                }
                sr.Align(cmd, Microsoft.Office.Core.MsoTriState.msoFalse);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AlignmentHelpers.AlignShapes");
                return false;
            }
        }

        public static bool AlignShapesToSlide(PowerPointManager manager, string direction)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                var shapeRange = manager.GetSelectedShapes();
                if (app == null || shapeRange == null || shapeRange.Count == 0) return false;

                bool aligned = false;
                for (int index = 1; index <= shapeRange.Count; index++)
                    aligned |= AlignShapeToSlide(app, shapeRange[index], direction);
                return aligned;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AlignmentHelpers.AlignShapesToSlide");
                return false;
            }
        }

        public static bool SnapSelectionToNearestGuide(PowerPointManager manager, string direction)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                var shapeRange = manager.GetSelectedShapes();
                if (app == null || shapeRange == null || shapeRange.Count == 0) return false;

                bool snapped = false;
                for (int index = 1; index <= shapeRange.Count; index++)
                    snapped |= SnapToNearestGuide(app, shapeRange[index], direction);
                return snapped;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AlignmentHelpers.SnapSelectionToNearestGuide");
                return false;
            }
        }

        public static bool AlignToSelectedAnchor(PowerPointManager manager, string edge, bool useLastSelected)
        {
            if (manager == null) return false;
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count < 2) return false;

            try
            {
                var anchor = shapeRange[useLastSelected ? shapeRange.Count : 1];
                for (int index = 1; index <= shapeRange.Count; index++)
                {
                    var shape = shapeRange[index];
                    if (shape.Id == anchor.Id) continue;
                    switch (edge.ToLowerInvariant())
                    {
                        case "left": shape.Left = anchor.Left; break;
                        case "right": shape.Left = anchor.Left + anchor.Width - shape.Width; break;
                        case "top": shape.Top = anchor.Top; break;
                        case "bottom": shape.Top = anchor.Top + anchor.Height - shape.Height; break;
                        default: return false;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AlignmentHelpers.AlignToSelectedAnchor");
                return false;
            }
        }

        // ──────────────────────────────────────────────────────
        // Align single shape to slide center/edges
        // ──────────────────────────────────────────────────────
        private static bool AlignShapeToSlide(Application app, Shape shape, string direction)
        {
            try
            {
                // View.Slide returns object, needs explicit cast
                var slideObj = app.ActiveWindow.View.Slide;
                if (slideObj == null) return false;

                Slide slide = slideObj as Slide;
                if (slide == null) return false;

                // Cast Parent to Presentation to access PageSetup
                Presentation presentation = slide.Parent as Presentation;
                if (presentation == null) return false;

                float slideWidth = (float)presentation.PageSetup.SlideWidth;
                float slideHeight = (float)presentation.PageSetup.SlideHeight;

                switch (direction.ToLower())
                {
                    case "left":
                        shape.Left = 0;
                        break;
                    case "right":
                        shape.Left = slideWidth - shape.Width;
                        break;
                    case "top":
                        shape.Top = 0;
                        break;
                    case "bottom":
                        shape.Top = slideHeight - shape.Height;
                        break;
                    case "center":
                        shape.Left = (slideWidth - shape.Width) / 2f;
                        break;
                    case "middle":
                        shape.Top = (slideHeight - shape.Height) / 2f;
                        break;
                    default:
                        return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AlignmentHelpers.AlignShapeToSlide");
                return false;
            }
        }

        // ──────────────────────────────────────────────────────
        // AlignToFirst — align shapes to first selected shape
        // ──────────────────────────────────────────────────────
        public static bool AlignToFirst(PowerPointManager manager, string edge)
        {
            if (manager == null) return false;
            var app = manager.GetApplication();
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count == 0) return false;

            // If only one shape selected, align to slide
            if (shapeRange.Count == 1)
            {
                string direction = "";
                switch (edge.ToLower())
                {
                    case "left": direction = "left"; break;
                    case "right": direction = "right"; break;
                    case "top": direction = "top"; break;
                    case "bottom": direction = "bottom"; break;
                    case "centerh": direction = "middle"; break; // Center horizontally aligned shapes = vertical center of slide
                    case "centerv": direction = "center"; break; // Center vertically aligned shapes = horizontal center of slide
                    default: return false;
                }
                return AlignShapeToSlide(app, shapeRange[1], direction);
            }

            var refShape = manager.GetReferenceShape();
            if (refShape == null) return false;

            try
            {
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var s = shapeRange[i];
                    if (s.Id == refShape.Id) continue;
                    switch (edge.ToLower())
                    {
                        case "left":            s.Left = refShape.Left; break;
                        case "right":           s.Left = refShape.Left + refShape.Width - s.Width; break;
                        case "top":             s.Top  = refShape.Top; break;
                        case "bottom":          s.Top  = refShape.Top + refShape.Height - s.Height; break;
                        case "centerh":         s.Top  = refShape.Top + (refShape.Height - s.Height) / 2f; break;
                        case "centerv":         s.Left = refShape.Left + (refShape.Width  - s.Width)  / 2f; break;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AlignmentHelpers.AlignToFirst");
                return false;
            }
        }

        // ──────────────────────────────────────────────────────
        // SnapToNearestGuide — snap single shape to nearest native guide line
        // Falls back to the slide edge if no guide is found.
        // ──────────────────────────────────────────────────────
        public static bool SnapToNearestGuide(Application app, Shape shape, string direction)
        {
            try
            {
                var slideObj = app.ActiveWindow.View.Slide;
                if (slideObj == null) return false;

                Slide slide = slideObj as Slide;
                if (slide == null) return false;

                Presentation presentation = slide.Parent as Presentation;
                if (presentation == null) return false;

                float slideWidth  = (float)presentation.PageSetup.SlideWidth;
                float slideHeight = (float)presentation.PageSetup.SlideHeight;

                System.Collections.Generic.List<float> vGuides = new System.Collections.Generic.List<float>();
                System.Collections.Generic.List<float> hGuides = new System.Collections.Generic.List<float>();

                // Collect native guides
                try
                {
                    if (presentation.Guides != null)
                    {
                        foreach (Microsoft.Office.Interop.PowerPoint.Guide g in presentation.Guides)
                        {
                            if (g.Orientation == PpGuideOrientation.ppVerticalGuide) vGuides.Add(g.Position);
                            else if (g.Orientation == PpGuideOrientation.ppHorizontalGuide) hGuides.Add(g.Position);
                        }
                    }
                }
                catch (Exception ex) { ReportOptionalAlignmentFailure(ex, "AlignmentHelpers.SnapToNearestGuide.PresentationGuides"); }

                try
                {
                    if (slide.Design != null && slide.Design.SlideMaster != null && slide.Design.SlideMaster.Guides != null)
                    {
                        foreach (Microsoft.Office.Interop.PowerPoint.Guide g in slide.Design.SlideMaster.Guides)
                        {
                            if (g.Orientation == PpGuideOrientation.ppVerticalGuide) vGuides.Add(g.Position);
                            else if (g.Orientation == PpGuideOrientation.ppHorizontalGuide) hGuides.Add(g.Position);
                        }
                    }
                }
                catch (Exception ex) { ReportOptionalAlignmentFailure(ex, "AlignmentHelpers.SnapToNearestGuide.MasterGuides"); }

                try
                {
                    if (slide.CustomLayout != null && slide.CustomLayout.Guides != null)
                    {
                        foreach (Microsoft.Office.Interop.PowerPoint.Guide g in slide.CustomLayout.Guides)
                        {
                            if (g.Orientation == PpGuideOrientation.ppVerticalGuide) vGuides.Add(g.Position);
                            else if (g.Orientation == PpGuideOrientation.ppHorizontalGuide) hGuides.Add(g.Position);
                        }
                    }
                }
                catch (Exception ex) { ReportOptionalAlignmentFailure(ex, "AlignmentHelpers.SnapToNearestGuide.LayoutGuides"); }

                switch (direction.ToLower())
                {
                    case "left":
                    {
                        float shapeLeft = shape.Left;
                        shape.Left = DirectionalGuideSelector.NearestAtOrBefore(vGuides, shapeLeft, 0f);
                        return true;
                    }
                    case "right":
                    {
                        float shapeRight = shape.Left + shape.Width;
                        float best = DirectionalGuideSelector.NearestAtOrAfter(vGuides, shapeRight, slideWidth);
                        shape.Left = best - shape.Width;
                        return true;
                    }
                    case "top":
                    {
                        float shapeTop = shape.Top;
                        shape.Top = DirectionalGuideSelector.NearestAtOrBefore(hGuides, shapeTop, 0f);
                        return true;
                    }
                    case "bottom":
                    {
                        float shapeBottom = shape.Top + shape.Height;
                        float best = DirectionalGuideSelector.NearestAtOrAfter(hGuides, shapeBottom, slideHeight);
                        shape.Top = best - shape.Height;
                        return true;
                    }
                    case "centerv": // Horizontal center of shape to vertical guide
                    {
                        float shapeCenter = shape.Left + shape.Width / 2f;
                        float best = DirectionalGuideSelector.NearestByDistance(vGuides, shapeCenter, slideWidth / 2f);
                        shape.Left = best - shape.Width / 2f;
                        return true;
                    }
                    case "centerh": // Vertical center of shape to horizontal guide
                    {
                        float shapeMiddle = shape.Top + shape.Height / 2f;
                        float best = DirectionalGuideSelector.NearestByDistance(hGuides, shapeMiddle, slideHeight / 2f);
                        shape.Top = best - shape.Height / 2f;
                        return true;
                    }
                    default:
                        return false;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AlignmentHelpers.SnapToNearestGuide");
                return false;
            }
        }

        // ──────────────────────────────────────────────────────
        // MoveUntilCollision — move single shape until it hits
        // another shape or the slide edge (used by context menu)
        // ──────────────────────────────────────────────────────
        public static bool MoveUntilCollision(Application app, Shape shape, string direction)
        {
            try
            {
                var slideObj = app.ActiveWindow.View.Slide;
                if (slideObj == null) return false;

                Slide slide = slideObj as Slide;
                if (slide == null) return false;

                Presentation presentation = slide.Parent as Presentation;
                if (presentation == null) return false;

                float slideWidth  = (float)presentation.PageSetup.SlideWidth;
                float slideHeight = (float)presentation.PageSetup.SlideHeight;

                switch (direction.ToLower())
                {
                    case "left":
                    {
                        float shapeTop    = shape.Top;
                        float shapeBottom = shape.Top + shape.Height;
                        float obstacle    = 0f;
                        for (int i = 1; i <= slide.Shapes.Count; i++)
                        {
                            var other = slide.Shapes[i];
                            if (ShouldIgnoreCollisionShape(other, shape)) continue;
                            float otherBottom = other.Top + other.Height;
                            float otherRight  = other.Left + other.Width;
                            bool vOverlap = (shapeTop < otherBottom + CollisionTolerance) && (shapeBottom > other.Top - CollisionTolerance);
                            if (vOverlap && otherRight <= shape.Left + CollisionTolerance && otherRight > obstacle)
                                obstacle = otherRight;
                        }
                        shape.Left = obstacle;
                        return true;
                    }
                    case "right":
                    {
                        float shapeTop    = shape.Top;
                        float shapeBottom = shape.Top + shape.Height;
                        float shapeRight  = shape.Left + shape.Width;
                        float obstacle    = slideWidth;
                        for (int i = 1; i <= slide.Shapes.Count; i++)
                        {
                            var other = slide.Shapes[i];
                            if (ShouldIgnoreCollisionShape(other, shape)) continue;
                            float otherBottom = other.Top + other.Height;
                            bool vOverlap = (shapeTop < otherBottom + CollisionTolerance) && (shapeBottom > other.Top - CollisionTolerance);
                            if (vOverlap && other.Left >= shapeRight - CollisionTolerance && other.Left < obstacle)
                                obstacle = other.Left;
                        }
                        shape.Left = obstacle - shape.Width;
                        return true;
                    }
                    case "top":
                    {
                        float shapeLeft  = shape.Left;
                        float shapeRight = shape.Left + shape.Width;
                        float obstacle   = 0f;
                        for (int i = 1; i <= slide.Shapes.Count; i++)
                        {
                            var other = slide.Shapes[i];
                            if (ShouldIgnoreCollisionShape(other, shape)) continue;
                            float otherRight  = other.Left + other.Width;
                            float otherBottom = other.Top + other.Height;
                            bool hOverlap = (shapeLeft < otherRight + CollisionTolerance) && (shapeRight > other.Left - CollisionTolerance);
                            if (hOverlap && otherBottom <= shape.Top + CollisionTolerance && otherBottom > obstacle)
                                obstacle = otherBottom;
                        }
                        shape.Top = obstacle;
                        return true;
                    }
                    case "bottom":
                    {
                        float shapeLeft   = shape.Left;
                        float shapeRight  = shape.Left + shape.Width;
                        float shapeBottom = shape.Top + shape.Height;
                        float obstacle    = slideHeight;
                        for (int i = 1; i <= slide.Shapes.Count; i++)
                        {
                            var other = slide.Shapes[i];
                            if (ShouldIgnoreCollisionShape(other, shape)) continue;
                            float otherRight = other.Left + other.Width;
                            bool hOverlap = (shapeLeft < otherRight + CollisionTolerance) && (shapeRight > other.Left - CollisionTolerance);
                            if (hOverlap && other.Top >= shapeBottom - CollisionTolerance && other.Top < obstacle)
                                obstacle = other.Top;
                        }
                        shape.Top = obstacle - shape.Height;
                        return true;
                    }
                    default:
                        return false;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AlignmentHelpers.MoveUntilCollision");
                return false;
            }
        }

        /// <summary>
        /// Centralized execution logic for alignment features.
        /// Handles selection validation, single-object snapping, and multi-object alignment.
        /// </summary>
        public static bool ExecuteAlignment(PowerPointManager manager, string direction)
        {
            if (manager == null) return false;

            var app = manager.GetApplication();
            if (app.ActiveWindow?.Selection?.Type != PpSelectionType.ppSelectionShapes) return false;

            var shapeRange = app.ActiveWindow.Selection.ShapeRange;
            if (shapeRange == null || shapeRange.Count == 0) return false;

            try { app.StartNewUndoEntry(); }
            catch (Exception ex) { ReportOptionalAlignmentFailure(ex, "AlignmentHelpers.MoveUntilCollision.StartUndoEntry"); }

            if (shapeRange.Count == 1)
                return SnapToNearestGuide(app, shapeRange[1], direction);

            return AlignToFirst(manager, direction);
        }

        public static bool ExecuteMoveUntilCollision(PowerPointManager manager, string direction)
        {
            if (manager == null) return false;
            var app = manager.GetApplication();
            var range = manager.GetSelectedShapes();
            if (range == null || range.Count != 1)
            {
                System.Windows.Forms.MessageBox.Show("Select exactly one shape for Move Until Object.", "Move Until Object", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return false;
            }
            try { app.StartNewUndoEntry(); }
            catch (Exception ex) { ReportOptionalAlignmentFailure(ex, "AlignmentHelpers.MoveSelectionUntilCollision.StartUndoEntry"); }
            return MoveUntilCollision(app, range[1], direction);
        }

        private static bool ShouldIgnoreCollisionShape(Shape candidate, Shape movingShape)
        {
            if (candidate.Id == movingShape.Id || candidate.Visible == Microsoft.Office.Core.MsoTriState.msoFalse) return true;
            string name = candidate.Name ?? string.Empty;
            return name.StartsWith(GuideH, StringComparison.OrdinalIgnoreCase) ||
                   name.StartsWith(GuideV, StringComparison.OrdinalIgnoreCase) ||
                   name.StartsWith(OpenEfficiencyGuide, StringComparison.OrdinalIgnoreCase);
        }

        // ──────────────────────────────────────────────────────
        // MatchSize — match width/height of selected to first
        // ──────────────────────────────────────────────────────
        public static bool MatchSize(PowerPointManager manager, bool matchWidth, bool matchHeight)
        {
            if (manager == null) return false;
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count < 2) return false;
            var refShape = manager.GetReferenceShape();
            if (refShape == null) return false;

            var originalStates = new System.Collections.Generic.List<Tuple<Shape, float, float, Microsoft.Office.Core.MsoTriState>>();
            try
            {
                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (Exception ex) { ReportOptionalAlignmentFailure(ex, "AlignmentHelpers.MatchSize.StartUndoEntry"); }
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var s = shapeRange[i];
                    if (s.Id == refShape.Id) continue;
                    originalStates.Add(Tuple.Create(s, s.Width, s.Height, s.LockAspectRatio));
                    s.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                    if (matchWidth)  s.Width  = refShape.Width;
                    if (matchHeight) s.Height = refShape.Height;
                    s.LockAspectRatio = originalStates[originalStates.Count - 1].Item4;
                }
                return true;
            }
            catch (Exception ex)
            {
                foreach (var state in originalStates)
                {
                    try
                    {
                        state.Item1.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        state.Item1.Width = state.Item2;
                        state.Item1.Height = state.Item3;
                        state.Item1.LockAspectRatio = state.Item4;
                    }
                    catch (Exception rollbackEx)
                    {
                        ExceptionLogger.Log(rollbackEx, "AlignmentHelpers.MatchSize.Rollback");
                    }
                }
                ExceptionLogger.Log(ex, "AlignmentHelpers.MatchSize");
                return false;
            }
        }

        private static void ReportOptionalAlignmentFailure(Exception exception, string context)
        {
            if (!ExceptionPolicy.IsExpectedComFailure(exception))
                ExceptionLogger.Log(exception, context);
            // Missing native guides or undo support falls back without blocking the alignment operation.
        }

        private const float StretchChangeTolerance = 0.01f;

        public static bool IsStretchFeature(string featureId)
        {
            switch (featureId)
            {
                case "BtnStretchLeft":
                case "BtnStretchRight":
                case "BtnStretchTop":
                case "BtnStretchBottom":
                case "BtnStretchLeftEdge":
                case "BtnStretchRightEdge":
                case "BtnStretchTopEdge":
                case "BtnStretchBottomEdge":
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Returns true when at least one non-anchor shape can be changed by the
        /// requested stretch operation without producing an inverted dimension.
        /// </summary>
        public static bool CanStretch(PowerPointManager manager, string featureId)
        {
            if (manager == null || !IsStretchFeature(featureId)) return false;

            try
            {
                var shapes = manager.GetSelectedShapes();
                if (shapes == null || shapes.Count < 2) return false;

                var anchor = manager.GetReferenceShape();
                if (anchor == null) return false;

                for (int i = 1; i <= shapes.Count; i++)
                {
                    var shape = shapes[i];
                    if (shape.Id != anchor.Id && IsStretchTargetEligible(featureId, shape, anchor))
                        return true;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AlignmentHelpers.CanStretch");
            }

            return false;
        }

        public static bool IsStretchTargetEligible(string featureId, Shape shape, Shape anchor)
        {
            if (shape == null || anchor == null || shape.Id == anchor.Id) return false;

            float shapeRight = shape.Left + shape.Width;
            float shapeBottom = shape.Top + shape.Height;
            float anchorRight = anchor.Left + anchor.Width;
            float anchorBottom = anchor.Top + anchor.Height;

            switch (featureId)
            {
                case "BtnStretchLeft":
                    return shape.Left > anchorRight;
                case "BtnStretchRight":
                    return shapeRight < anchor.Left;
                case "BtnStretchTop":
                    return shape.Top > anchorBottom;
                case "BtnStretchBottom":
                    return shapeBottom < anchor.Top;
                case "BtnStretchLeftEdge":
                    return shapeRight > anchor.Left &&
                           Math.Abs(shape.Left - anchor.Left) > StretchChangeTolerance;
                case "BtnStretchRightEdge":
                    return anchorRight > shape.Left &&
                           Math.Abs(shapeRight - anchorRight) > StretchChangeTolerance;
                case "BtnStretchTopEdge":
                    return shapeBottom > anchor.Top &&
                           Math.Abs(shape.Top - anchor.Top) > StretchChangeTolerance;
                case "BtnStretchBottomEdge":
                    return anchorBottom > shape.Top &&
                           Math.Abs(shapeBottom - anchorBottom) > StretchChangeTolerance;
                default:
                    return false;
            }
        }
    }
}
