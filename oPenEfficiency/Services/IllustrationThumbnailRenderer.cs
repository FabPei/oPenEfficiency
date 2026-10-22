using System;
using System.Windows;
using System.Windows.Media;
using oPenEfficiency.UI;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Creates a frozen WPF vector preview from a screened illustration SVG. The renderer deliberately
    /// keeps previews in memory only: bundled and imported artwork remains the single source of truth.
    /// Unsupported SVG paints such as gradients fall back to a neutral brush while the original SVG is
    /// still used for PowerPoint insertion.
    /// </summary>
    internal static class IllustrationThumbnailRenderer
    {
        private static readonly Brush FallbackFill = Freeze(new SolidColorBrush(Color.FromRgb(210, 215, 222)));
        private static readonly Brush FallbackStroke = Freeze(new SolidColorBrush(Color.FromRgb(79, 87, 98)));

        public static ImageSource Create(string path, double size = 72d)
        {
            ParsedIconGeometry parsed = SvgParser.ParseSvgWithRoles(path);
            if (parsed == null || parsed.Elements.Count == 0) return null;

            Rect viewBox = parsed.ViewBox;
            if (viewBox.Width <= 0d || viewBox.Height <= 0d) return null;

            double scale = Math.Min(size / viewBox.Width, size / viewBox.Height);
            double offsetX = (size - (viewBox.Width * scale)) / 2d - (viewBox.X * scale);
            double offsetY = (size - (viewBox.Height * scale)) / 2d - (viewBox.Y * scale);
            var transform = new MatrixTransform(scale, 0d, 0d, scale, offsetX, offsetY);
            transform.Freeze();

            var group = new DrawingGroup { Transform = transform };
            foreach (ParsedIconElement element in parsed.Elements)
            {
                Brush fill = element.HasFill ? ResolveBrush(element.FillPaint, FallbackFill) : null;
                Pen pen = null;
                if (element.HasStroke)
                {
                    pen = new Pen(
                        ResolveBrush(element.StrokePaint, FallbackStroke),
                        Math.Max(0.5d / scale, element.StrokeThickness));
                    pen.StartLineCap = element.StrokeLineCap;
                    pen.EndLineCap = element.StrokeLineCap;
                    pen.LineJoin = element.StrokeLineJoin;
                    pen.Freeze();
                }

                group.Children.Add(new GeometryDrawing(fill, pen, element.Geometry));
            }

            group.ClipGeometry = new RectangleGeometry(viewBox);
            group.Freeze();
            var image = new DrawingImage(group);
            image.Freeze();
            return image;
        }

        private static Brush ResolveBrush(string paint, Brush fallback)
        {
            if (string.IsNullOrWhiteSpace(paint) ||
                paint.StartsWith("url(", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(paint, "currentColor", StringComparison.OrdinalIgnoreCase))
            {
                return fallback;
            }

            try
            {
                var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(paint));
                brush.Freeze();
                return brush;
            }
            catch (FormatException)
            {
                return fallback;
            }
            catch (NotSupportedException)
            {
                return fallback;
            }
        }

        private static Brush Freeze(Brush brush)
        {
            brush.Freeze();
            return brush;
        }
    }
}
