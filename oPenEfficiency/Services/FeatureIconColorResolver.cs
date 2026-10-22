using System.Collections.Generic;
using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Text.RegularExpressions;
using oPenEfficiency.Models;
using oPenEfficiency.UI;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Applies the host-neutral icon color policy to WPF paths. Neutral decisions use dynamic
    /// resources so open windows follow Light/Dark theme changes without rebuilding their icons.
    /// Supports two-layer rendering: base layer uses Office-neutral foreground; accent layer
    /// uses a semantic accent resource.
    /// </summary>
    internal static class FeatureIconColorResolver
    {
        public static void ApplyToPath(
            Path path,
            ParsedIconElement element,
            string mode,
            FeatureIconVisualState state,
            Brush storedFeatureColor,
            double viewportScale = 1d)
        {
            if (path == null || element == null) return;

            FeatureIconColorDecision decision = FeatureIconColorPolicy.ResolveForLayer(
                mode,
                state,
                storedFeatureColor?.ToString(),
                element.IsAccent,
                element.AccentRole);

            if (element.HasFill)
            {
                var contentBrush = ResolveContentBrush(element, element.FillPaint, mode, state);
                if (contentBrush != null) path.Fill = contentBrush;
                else ApplyBrush(path, Shape.FillProperty, decision, storedFeatureColor);
            }
            else
            {
                path.Fill = Brushes.Transparent;
            }

            if (element.HasStroke)
            {
                var contentBrush = ResolveContentBrush(element, element.StrokePaint, mode, state);
                if (contentBrush != null) path.Stroke = contentBrush;
                else ApplyBrush(path, Shape.StrokeProperty, decision, storedFeatureColor);
            }
            else
            {
                path.Stroke = null;
            }

            path.StrokeThickness = element.HasStroke ? element.StrokeThickness * viewportScale : 0d;
            path.StrokeLineJoin = element.StrokeLineJoin;
            path.StrokeStartLineCap = element.StrokeLineCap;
            path.StrokeEndLineCap = element.StrokeLineCap;
        }

        private static Brush ResolveContentBrush(ParsedIconElement element, string paint,
            string mode, FeatureIconVisualState state)
        {
            // Only explicitly reviewed content samples retain literal colors. Tool/frame
            // geometry and disabled/Feature Colors modes still follow the shared policy.
            if (!element.IsAccent || !element.HasExplicitAccentRole ||
                element.AccentRole != FeatureIconAccentRole.ContentColor ||
                state == FeatureIconVisualState.Disabled ||
                IconColorModes.Normalize(mode) != IconColorModes.PowerPoint ||
                string.IsNullOrEmpty(paint) ||
                !Regex.IsMatch(paint, @"\A#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})\z"))
                return null;
            var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(paint));
            brush.Freeze();
            return brush;
        }

        private static void ApplyBrush(
            FrameworkElement target,
            DependencyProperty property,
            FeatureIconColorDecision decision,
            Brush storedFeatureColor)
        {
            if (decision.UsesStoredFeatureColor && storedFeatureColor != null)
            {
                target.SetValue(property, storedFeatureColor);
                return;
            }

            target.SetResourceReference(property, decision.ResourceKey);
        }
    }

    internal static class FeatureIconPathFactory
    {
        public static IReadOnlyList<Path> CreatePaths(
            ParsedIconGeometry parsed,
            string mode,
            FeatureIconVisualState state,
            Brush storedFeatureColor)
        {
            var paths = new List<Path>();
            if (parsed == null) return paths;

            foreach (ParsedIconElement element in parsed.Elements)
            {
                paths.Add(CreatePath(element, parsed.ViewBox));
            }

            ApplyState(paths, parsed, mode, state, storedFeatureColor);
            return paths;
        }

        /// <summary>
        /// Creates the geometry for one element. Paths produced here are state-independent and can
        /// be re-skinned with <see cref="ApplyState"/>; visual-state changes must never recreate
        /// them because that destroys hit elements mid mouse-dispatch (PQ-015/PQ-016).
        /// </summary>
        public static Path CreatePath(ParsedIconElement element, Rect viewBox)
        {
            double scale = ResolveViewportScale(viewBox);
            return new Path
            {
                Data = NormalizeGeometry(element.Geometry, viewBox, scale),
                Width = 24d,
                Height = 24d,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                Stretch = Stretch.None
            };
        }

        /// <summary>
        /// Re-resolves brushes for existing paths. Keeps the same Path instances, so it is safe to
        /// call from mouse-driven state changes while input is in dispatch.
        /// </summary>
        public static void ApplyState(
            IReadOnlyList<Path> paths,
            ParsedIconGeometry parsed,
            string mode,
            FeatureIconVisualState state,
            Brush storedFeatureColor)
        {
            if (parsed == null) return;

            double scale = ResolveViewportScale(parsed.ViewBox);
            for (int i = 0; i < paths.Count && i < parsed.Elements.Count; i++)
            {
                FeatureIconColorResolver.ApplyToPath(
                    paths[i],
                    parsed.Elements[i],
                    mode,
                    state,
                    storedFeatureColor,
                    scale);
            }
        }

        private static double ResolveViewportScale(Rect viewBox)
        {
            if (viewBox.Width <= 0d || viewBox.Height <= 0d ||
                double.IsNaN(viewBox.Width) || double.IsNaN(viewBox.Height))
            {
                return 1d;
            }

            return Math.Min(24d / viewBox.Width, 24d / viewBox.Height);
        }

        private static Geometry NormalizeGeometry(Geometry geometry, Rect viewBox, double scale)
        {
            if (geometry == null) return null;

            double offsetX = (24d - (viewBox.Width * scale)) / 2d - (viewBox.X * scale);
            double offsetY = (24d - (viewBox.Height * scale)) / 2d - (viewBox.Y * scale);
            var group = new GeometryGroup
            {
                Transform = new MatrixTransform(scale, 0d, 0d, scale, offsetX, offsetY)
            };
            group.Children.Add(geometry);
            return group;
        }
    }
}
