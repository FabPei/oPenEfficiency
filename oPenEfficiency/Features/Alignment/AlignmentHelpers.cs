using System;
using Microsoft.Office.Interop.PowerPoint;
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
                catch { }

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
                catch { }

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
                catch { }

                switch (direction.ToLower())
                {
                    case "left":
                    {
                        vGuides.Add(0f); // Fallback to slide left edge
                        float shapeLeft = shape.Left;
                        float center = shape.Left + shape.Width / 2f;
                        float best = 0f;
                        float minDiff = float.MaxValue;
                        foreach (float gx in vGuides)
                        {
                            if (gx > center && gx != 0f) continue;
                            float diff = Math.Abs(gx - shapeLeft);
                            if (diff < minDiff)
                            {
                                minDiff = diff;
                                best = gx;
                            }
                        }
                        shape.Left = best;
                        return true;
                    }
                    case "right":
                    {
                        vGuides.Add(slideWidth); // Fallback to slide right edge
                        float shapeRight = shape.Left + shape.Width;
                        float center = shape.Left + shape.Width / 2f;
                        float best = slideWidth;
                        float minDiff = float.MaxValue;
                        foreach (float gx in vGuides)
                        {
                            if (gx < center && gx != slideWidth) continue;
                            float diff = Math.Abs(gx - shapeRight);
                            if (diff < minDiff)
                            {
                                minDiff = diff;
                                best = gx;
                            }
                        }
                        shape.Left = best - shape.Width;
                        return true;
                    }
                    case "top":
                    {
                        hGuides.Add(0f); // Fallback to slide top edge
                        float shapeTop = shape.Top;
                        float center = shape.Top + shape.Height / 2f;
                        float best = 0f;
                        float minDiff = float.MaxValue;
                        foreach (float gy in hGuides)
                        {
                            if (gy > center && gy != 0f) continue;
                            float diff = Math.Abs(gy - shapeTop);
                            if (diff < minDiff)
                            {
                                minDiff = diff;
                                best = gy;
                            }
                        }
                        shape.Top = best;
                        return true;
                    }
                    case "bottom":
                    {
                        hGuides.Add(slideHeight); // Fallback to slide bottom edge
                        float shapeBottom = shape.Top + shape.Height;
                        float center = shape.Top + shape.Height / 2f;
                        float best = slideHeight;
                        float minDiff = float.MaxValue;
                        foreach (float gy in hGuides)
                        {
                            if (gy < center && gy != slideHeight) continue;
                            float diff = Math.Abs(gy - shapeBottom);
                            if (diff < minDiff)
                            {
                                minDiff = diff;
                                best = gy;
                            }
                        }
                        shape.Top = best - shape.Height;
                        return true;
                    }
                    case "centerv": // Horizontal center of shape to vertical guide
                    {
                        vGuides.Add(slideWidth / 2f); // Fallback to slide center
                        float shapeCenter = shape.Left + shape.Width / 2f;
                        float best = slideWidth / 2f;
                        float minDiff = float.MaxValue;
                        foreach (float gx in vGuides)
                        {
                            float diff = Math.Abs(gx - shapeCenter);
                            if (diff < minDiff)
                            {
                                minDiff = diff;
                                best = gx;
                            }
                        }
                        shape.Left = best - shape.Width / 2f;
                        return true;
                    }
                    case "centerh": // Vertical center of shape to horizontal guide
                    {
                        hGuides.Add(slideHeight / 2f); // Fallback to slide middle
                        float shapeMiddle = shape.Top + shape.Height / 2f;
                        float best = slideHeight / 2f;
                        float minDiff = float.MaxValue;
                        foreach (float gy in hGuides)
                        {
                            float diff = Math.Abs(gy - shapeMiddle);
                            if (diff < minDiff)
                            {
                                minDiff = diff;
                                best = gy;
                            }
                        }
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
                            if (other.Id == shape.Id) continue;
                            if (other.Name.StartsWith(GuideH) || other.Name.StartsWith(GuideV)) continue;
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
                            if (other.Id == shape.Id) continue;
                            if (other.Name.StartsWith(GuideH) || other.Name.StartsWith(GuideV)) continue;
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
                            if (other.Id == shape.Id) continue;
                            if (other.Name.StartsWith(GuideH) || other.Name.StartsWith(GuideV)) continue;
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
                            if (other.Id == shape.Id) continue;
                            if (other.Name.StartsWith(GuideH) || other.Name.StartsWith(GuideV)) continue;
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

            if (shapeRange.Count == 1)
                return SnapToNearestGuide(app, shapeRange[1], direction);

            return AlignToFirst(manager, direction);
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

            try
            {
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var s = shapeRange[i];
                    if (s.Id == refShape.Id) continue;
                    if (matchWidth)  s.Width  = refShape.Width;
                    if (matchHeight) s.Height = refShape.Height;
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AlignmentHelpers.MatchSize");
                return false;
            }
        }
    }
}
