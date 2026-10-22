using System;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Xml.Linq;
using oPenEfficiency.UI;

namespace oPenEfficiency.Utils
{
    public static class SvgParser
    {
        public static GeometryGroup ParseLucideSvg(string filePath, out bool isFilled)
        {
            isFilled = false;
            if (!File.Exists(filePath))
                return null;

            try
            {
                var svgDoc = XDocument.Load(filePath);
                if (svgDoc.Root == null) return null;

                var fillAttr = svgDoc.Root.Attribute("fill");
                if (fillAttr != null && !string.IsNullOrEmpty(fillAttr.Value))
                {
                    string fillVal = fillAttr.Value.ToLower().Trim();
                    if (fillVal != "none")
                    {
                        isFilled = true;
                    }
                }

                var group = new GeometryGroup();
                ParseElements(svgDoc.Root.Elements(), group);
                
                return group.Children.Count > 0 ? group : null;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SvgParser.ParseLucideSvg");
                return null;
            }
        }

        /// <summary>
        /// Role-aware SVG parser that splits geometry into Base and Accent layers based on
        /// <c>data-oe-role</c> attributes. Elements with <c>data-oe-role="accent"</c> become
        /// the accent layer; all other drawable elements become the base layer.
        /// Icons without any role metadata render entirely as base (backward compatible).
        /// </summary>
        public static ParsedIconGeometry ParseSvgWithRoles(string filePath)
        {
            if (!File.Exists(filePath))
                return null;

            try
            {
                var svgDoc = XDocument.Load(filePath);
                if (svgDoc.Root == null) return null;

                var elements = new System.Collections.Generic.List<ParsedIconElement>();
                ParseElementsWithRoles(
                    svgDoc.Root.Elements(),
                    elements,
                    SvgPaintStyle.Resolve(SvgPaintStyle.Default, svgDoc.Root),
                    inheritedAccent: false,
                    FeatureIconAccentRole.None,
                    inheritedAccentRoleIsExplicit: false,
                    new System.Collections.Generic.List<Transform>());

                var baseGroup = new GeometryGroup();
                var accentGroup = new GeometryGroup();
                foreach (ParsedIconElement element in elements)
                {
                    (element.IsAccent ? accentGroup : baseGroup).Children.Add(element.Geometry);
                }

                var result = new ParsedIconGeometry
                {
                    ViewBox = ParseViewBox(svgDoc.Root),
                    BaseGeometry = baseGroup.Children.Count > 0 ? baseGroup : null,
                    AccentGeometry = accentGroup.Children.Count > 0 ? accentGroup : null
                };
                foreach (ParsedIconElement element in elements)
                {
                    result.Elements.Add(element);
                }

                return result.BaseGeometry != null ? result : null;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SvgParser.ParseSvgWithRoles");
                return null;
            }
        }

        private static void ParseElements(System.Collections.Generic.IEnumerable<XElement> elements, GeometryGroup group)
        {
            var geomConverter = new GeometryConverter();

            foreach (var el in elements)
            {
                try
                {
                    // Skip hidden elements
                    if (IsHidden(el)) continue;

                    Geometry geom = null;
                    string name = el.Name.LocalName.ToLower();

                    if (name == "path")
                    {
                        var d = el.Attribute("d")?.Value;
                        if (!string.IsNullOrEmpty(d))
                        {
                            geom = (Geometry)geomConverter.ConvertFromInvariantString(d);
                        }
                    }
                    else if (name == "rect")
                    {
                        double x = GetDouble(el, "x");
                        double y = GetDouble(el, "y");
                        double w = GetDouble(el, "width");
                        double h = GetDouble(el, "height");
                        double rx = GetDouble(el, "rx");
                        double ry = GetDouble(el, "ry", rx);
                        geom = new RectangleGeometry(new Rect(x, y, w, h), rx, ry);
                    }
                    else if (name == "circle")
                    {
                        double cx = GetDouble(el, "cx");
                        double cy = GetDouble(el, "cy");
                        double r = GetDouble(el, "r");
                        geom = new EllipseGeometry(new Point(cx, cy), r, r);
                    }
                    else if (name == "ellipse")
                    {
                        double cx = GetDouble(el, "cx");
                        double cy = GetDouble(el, "cy");
                        double rx = GetDouble(el, "rx");
                        double ry = GetDouble(el, "ry");
                        geom = new EllipseGeometry(new Point(cx, cy), rx, ry);
                    }
                    else if (name == "line")
                    {
                        double x1 = GetDouble(el, "x1");
                        double y1 = GetDouble(el, "y1");
                        double x2 = GetDouble(el, "x2");
                        double y2 = GetDouble(el, "y2");
                        geom = new LineGeometry(new Point(x1, y1), new Point(x2, y2));
                    }
                    else if (name == "polyline" || name == "polygon")
                    {
                        var pointsAttr = el.Attribute("points")?.Value;
                        if (!string.IsNullOrEmpty(pointsAttr))
                        {
                            var pointCol = PointCollection.Parse(pointsAttr);
                            if (pointCol.Count > 0)
                            {
                                var figure = new PathFigure { StartPoint = pointCol[0] };
                                var segment = new PolyLineSegment();
                                for (int i = 1; i < pointCol.Count; i++)
                                    segment.Points.Add(pointCol[i]);
                                figure.Segments.Add(segment);
                                if (name == "polygon") figure.IsClosed = true;
                                var pGeom = new PathGeometry();
                                pGeom.Figures.Add(figure);
                                geom = pGeom;
                            }
                        }
                    }
                    else if (name == "g" || name == "a" || name == "svg")
                    {
                        Transform groupTransform = ParseElementTransform(el);
                        if (groupTransform != null)
                        {
                            var childGroup = new GeometryGroup();
                            childGroup.Transform = groupTransform;
                            ParseElements(el.Elements(), childGroup);
                            if (childGroup.Children.Count > 0) group.Children.Add(childGroup);
                            continue;
                        }

                        ParseElements(el.Elements(), group);
                        continue;
                    }
                    else if (name == "defs" || name == "metadata" || name == "title" || name == "desc")
                    {
                        // Explicitly ignore non-visual containers
                        continue;
                    }

                    if (geom != null)
                    {
                        Transform elementTransform = ParseElementTransform(el);
                        if (elementTransform != null)
                        {
                            geom.Transform = elementTransform;
                        }
                        group.Children.Add(geom);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"SvgParser: Failed to parse element {el.Name.LocalName}: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Role-aware element parser. Checks <c>data-oe-role</c> on each element and routes
        /// geometry to the base or accent target group. For <c>&lt;g&gt;</c> elements, the role
        /// attribute on the group applies to all its children.
        /// </summary>
        private static void ParseElementsWithRoles(
            System.Collections.Generic.IEnumerable<XElement> elements,
            System.Collections.Generic.ICollection<ParsedIconElement> parsedElements,
            SvgPaintStyle inheritedStyle,
            bool inheritedAccent,
            FeatureIconAccentRole inheritedAccentRole,
            bool inheritedAccentRoleIsExplicit,
            System.Collections.Generic.IReadOnlyList<Transform> inheritedTransforms)
        {
            var geomConverter = new GeometryConverter();

            foreach (var el in elements)
            {
                try
                {
                    if (IsHidden(el)) continue;

                    string role = GetOeRole(el);
                    bool isAccent = role == "accent" || (role != "base" && inheritedAccent);
                    FeatureIconAccentRole accentRole = GetAccentRole(el, inheritedAccentRole, isAccent);
                    bool hasExplicitAccentRole = isAccent &&
                        (el.Attribute("data-oe-accent-role") != null || inheritedAccentRoleIsExplicit);
                    SvgPaintStyle style = SvgPaintStyle.Resolve(inheritedStyle, el);

                    Geometry geom = null;
                    string name = el.Name.LocalName.ToLower();

                    if (name == "path")
                    {
                        var d = el.Attribute("d")?.Value;
                        if (!string.IsNullOrEmpty(d))
                        {
                            geom = (Geometry)geomConverter.ConvertFromInvariantString(d);
                        }
                    }
                    else if (name == "rect")
                    {
                        double x = GetDouble(el, "x");
                        double y = GetDouble(el, "y");
                        double w = GetDouble(el, "width");
                        double h = GetDouble(el, "height");
                        double rx = GetDouble(el, "rx");
                        double ry = GetDouble(el, "ry", rx);
                        geom = new RectangleGeometry(new Rect(x, y, w, h), rx, ry);
                    }
                    else if (name == "circle")
                    {
                        double cx = GetDouble(el, "cx");
                        double cy = GetDouble(el, "cy");
                        double r = GetDouble(el, "r");
                        geom = new EllipseGeometry(new Point(cx, cy), r, r);
                    }
                    else if (name == "ellipse")
                    {
                        double cx = GetDouble(el, "cx");
                        double cy = GetDouble(el, "cy");
                        double rx = GetDouble(el, "rx");
                        double ry = GetDouble(el, "ry");
                        geom = new EllipseGeometry(new Point(cx, cy), rx, ry);
                    }
                    else if (name == "line")
                    {
                        double x1 = GetDouble(el, "x1");
                        double y1 = GetDouble(el, "y1");
                        double x2 = GetDouble(el, "x2");
                        double y2 = GetDouble(el, "y2");
                        geom = new LineGeometry(new Point(x1, y1), new Point(x2, y2));
                    }
                    else if (name == "polyline" || name == "polygon")
                    {
                        var pointsAttr = el.Attribute("points")?.Value;
                        if (!string.IsNullOrEmpty(pointsAttr))
                        {
                            var pointCol = PointCollection.Parse(pointsAttr);
                            if (pointCol.Count > 0)
                            {
                                var figure = new PathFigure { StartPoint = pointCol[0] };
                                var segment = new PolyLineSegment();
                                for (int i = 1; i < pointCol.Count; i++)
                                    segment.Points.Add(pointCol[i]);
                                figure.Segments.Add(segment);
                                if (name == "polygon") figure.IsClosed = true;
                                var pGeom = new PathGeometry();
                                pGeom.Figures.Add(figure);
                                geom = pGeom;
                            }
                        }
                    }
                    else if (name == "g" || name == "a" || name == "svg")
                    {
                        var childTransforms = new System.Collections.Generic.List<Transform>(inheritedTransforms);
                        Transform groupTransform = ParseElementTransform(el);
                        if (groupTransform != null)
                        {
                            childTransforms.Add(groupTransform);
                        }

                        ParseElementsWithRoles(
                            el.Elements(),
                            parsedElements,
                            style,
                            isAccent,
                            accentRole,
                            hasExplicitAccentRole,
                            childTransforms);
                        continue;
                    }
                    else if (name == "defs" || name == "metadata" || name == "title" || name == "desc")
                    {
                        continue;
                    }

                    if (geom != null)
                    {
                        var transforms = new System.Collections.Generic.List<Transform>(inheritedTransforms);
                        Transform elementTransform = ParseElementTransform(el);
                        if (elementTransform != null)
                        {
                            transforms.Add(elementTransform);
                        }

                        parsedElements.Add(new ParsedIconElement
                        {
                            Geometry = ApplyTransforms(geom, transforms),
                            IsAccent = isAccent,
                            AccentRole = accentRole,
                            HasExplicitAccentRole = hasExplicitAccentRole,
                            HasFill = HasPaint(style.Fill),
                            HasStroke = HasPaint(style.Stroke),
                            FillPaint = style.Fill,
                            StrokePaint = style.Stroke,
                            StrokeThickness = style.StrokeWidth,
                            StrokeLineCap = style.StrokeLineCap,
                            StrokeLineJoin = style.StrokeLineJoin
                        });
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"SvgParser: Failed to parse element {el.Name.LocalName}: {ex.Message}");
                }
            }
        }

        private static string GetOeRole(XElement el)
        {
            // Check data-oe-role attribute (namespaced and non-namespaced)
            var attr = el.Attribute("data-oe-role");
            if (attr != null) return attr.Value.ToLower().Trim();
            return null;
        }

        private static FeatureIconAccentRole GetAccentRole(
            XElement element,
            FeatureIconAccentRole inheritedRole,
            bool isAccent)
        {
            if (!isAccent) return FeatureIconAccentRole.None;

            string value = element.Attribute("data-oe-accent-role")?.Value;
            if (string.IsNullOrWhiteSpace(value))
            {
                return inheritedRole == FeatureIconAccentRole.None
                    ? FeatureIconAccentRole.ActionBlue
                    : inheritedRole;
            }

            return Enum.TryParse(value.Trim(), ignoreCase: true, out FeatureIconAccentRole role)
                ? role
                : FeatureIconAccentRole.ActionBlue;
        }

        private static Geometry ApplyTransforms(
            Geometry geometry,
            System.Collections.Generic.IReadOnlyList<Transform> transforms)
        {
            Geometry transformed = geometry;
            for (int index = transforms.Count - 1; index >= 0; index--)
            {
                var group = new GeometryGroup { Transform = transforms[index] };
                group.Children.Add(transformed);
                transformed = group;
            }

            return transformed;
        }

        private static bool HasPaint(string value)
        {
            return !string.IsNullOrWhiteSpace(value) &&
                   !string.Equals(value, "none", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(value, "transparent", StringComparison.OrdinalIgnoreCase);
        }

        private static Rect ParseViewBox(XElement root)
        {
            string value = root?.Attribute("viewBox")?.Value;
            if (string.IsNullOrWhiteSpace(value))
                return ParsedIconGeometry.DefaultViewBox;

            string[] parts = value.Split(new[] { ',', ' ', '\t', '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 4)
                return ParsedIconGeometry.DefaultViewBox;

            if (!double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double x) ||
                !double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double y) ||
                !double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out double width) ||
                !double.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out double height) ||
                width <= 0d || height <= 0d)
            {
                return ParsedIconGeometry.DefaultViewBox;
            }

            return new Rect(x, y, width, height);
        }

        private sealed class SvgPaintStyle
        {
            public static SvgPaintStyle Default => new SvgPaintStyle
            {
                Fill = "black",
                Stroke = "none",
                StrokeWidth = 1d,
                StrokeLineCap = PenLineCap.Flat,
                StrokeLineJoin = PenLineJoin.Miter
            };

            public string Fill { get; private set; }
            public string Stroke { get; private set; }
            public double StrokeWidth { get; private set; }
            public PenLineCap StrokeLineCap { get; private set; }
            public PenLineJoin StrokeLineJoin { get; private set; }

            public static SvgPaintStyle Resolve(SvgPaintStyle inherited, XElement element)
            {
                var resolved = new SvgPaintStyle
                {
                    Fill = inherited.Fill,
                    Stroke = inherited.Stroke,
                    StrokeWidth = inherited.StrokeWidth,
                    StrokeLineCap = inherited.StrokeLineCap,
                    StrokeLineJoin = inherited.StrokeLineJoin
                };

                resolved.Fill = GetPresentationValue(element, "fill") ?? resolved.Fill;
                resolved.Stroke = GetPresentationValue(element, "stroke") ?? resolved.Stroke;

                string strokeWidth = GetPresentationValue(element, "stroke-width");
                if (double.TryParse(
                    strokeWidth,
                    NumberStyles.Any,
                    CultureInfo.InvariantCulture,
                    out double parsedWidth))
                {
                    resolved.StrokeWidth = parsedWidth;
                }

                string lineCap = GetPresentationValue(element, "stroke-linecap");
                if (!string.IsNullOrWhiteSpace(lineCap))
                {
                    resolved.StrokeLineCap = string.Equals(lineCap, "round", StringComparison.OrdinalIgnoreCase)
                        ? PenLineCap.Round
                        : string.Equals(lineCap, "square", StringComparison.OrdinalIgnoreCase)
                            ? PenLineCap.Square
                            : PenLineCap.Flat;
                }

                string lineJoin = GetPresentationValue(element, "stroke-linejoin");
                if (!string.IsNullOrWhiteSpace(lineJoin))
                {
                    resolved.StrokeLineJoin = string.Equals(lineJoin, "round", StringComparison.OrdinalIgnoreCase)
                        ? PenLineJoin.Round
                        : string.Equals(lineJoin, "bevel", StringComparison.OrdinalIgnoreCase)
                            ? PenLineJoin.Bevel
                            : PenLineJoin.Miter;
                }

                return resolved;
            }

            private static string GetPresentationValue(XElement element, string name)
            {
                string value = element.Attribute(name)?.Value;
                string inlineStyle = element.Attribute("style")?.Value;
                if (string.IsNullOrWhiteSpace(inlineStyle)) return value;

                foreach (string declaration in inlineStyle.Split(';'))
                {
                    string[] parts = declaration.Split(new[] { ':' }, 2);
                    if (parts.Length == 2 &&
                        string.Equals(parts[0].Trim(), name, StringComparison.OrdinalIgnoreCase))
                    {
                        return parts[1].Trim();
                    }
                }

                return value;
            }
        }

        private static bool IsHidden(XElement el)
        {
            // Check visibility attribute
            var visibility = el.Attribute("visibility")?.Value?.ToLower();
            if (visibility == "hidden" || visibility == "collapse") return true;

            // Check display attribute
            var display = el.Attribute("display")?.Value?.ToLower();
            if (display == "none") return true;

            // Check style attribute for inline CSS
            var style = el.Attribute("style")?.Value?.ToLower();
            if (style != null)
            {
                if (style.Contains("visibility:hidden") || style.Contains("visibility: hidden")) return true;
                if (style.Contains("display:none") || style.Contains("display: none")) return true;
            }

            return false;
        }

        private static Transform ParseTransform(string transformStr)
        {
            if (string.IsNullOrEmpty(transformStr)) return null;
            try
            {
                transformStr = transformStr.Trim();
                
                // Handle matrix(...) specifically to ensure invariant parsing
                if (transformStr.StartsWith("matrix(", StringComparison.OrdinalIgnoreCase))
                {
                    string values = transformStr.Substring(7).TrimEnd(')');
                    // Matrix.Parse is also culture-sensitive, so we use Matrix.Parse with invariant if possible
                    // but WPF Matrix doesn't have an invariant Parse. We use the converter.
                    var matrix = (Matrix)new MatrixConverter().ConvertFromInvariantString(values);
                    return new MatrixTransform(matrix);
                }

                if (transformStr.StartsWith("translate(", StringComparison.OrdinalIgnoreCase))
                {
                    string values = transformStr.Substring(10).TrimEnd(')');
                    string[] parts = values.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    double x = double.Parse(parts[0], CultureInfo.InvariantCulture);
                    double y = parts.Length > 1
                        ? double.Parse(parts[1], CultureInfo.InvariantCulture)
                        : 0d;
                    return new TranslateTransform(x, y);
                }

                // Fallback for other transform types (translate, rotate, etc.)
                // TransformConverter is generally good at handling these invariantly
                return (Transform)new TransformConverter().ConvertFromInvariantString(transformStr);
            }
            catch
            {
                return null;
            }
        }

        private static Transform ParseElementTransform(XElement element)
        {
            Transform transform = ParseTransform(element?.Attribute("transform")?.Value);
            if (transform == null) return null;

            Point? origin = ParseTransformOrigin(element);
            if (!origin.HasValue) return transform;

            Matrix matrix = transform.Value;
            Point point = origin.Value;
            matrix.OffsetX += point.X - ((matrix.M11 * point.X) + (matrix.M21 * point.Y));
            matrix.OffsetY += point.Y - ((matrix.M12 * point.X) + (matrix.M22 * point.Y));
            return new MatrixTransform(matrix);
        }

        private static Point? ParseTransformOrigin(XElement element)
        {
            string value = element?.Attribute("transform-origin")?.Value;
            string inlineStyle = element?.Attribute("style")?.Value;
            if (!string.IsNullOrWhiteSpace(inlineStyle))
            {
                foreach (string declaration in inlineStyle.Split(';'))
                {
                    string[] parts = declaration.Split(new[] { ':' }, 2);
                    if (parts.Length == 2 &&
                        string.Equals(parts[0].Trim(), "transform-origin", StringComparison.OrdinalIgnoreCase))
                    {
                        value = parts[1].Trim();
                        break;
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(value)) return null;
            string[] coordinates = value.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (coordinates.Length < 2 ||
                !TryParseSvgLength(coordinates[0], out double x) ||
                !TryParseSvgLength(coordinates[1], out double y))
            {
                return null;
            }

            return new Point(x, y);
        }

        private static bool TryParseSvgLength(string value, out double result)
        {
            string normalized = value?.Trim();
            if (!string.IsNullOrEmpty(normalized) && normalized.EndsWith("px", StringComparison.OrdinalIgnoreCase))
                normalized = normalized.Substring(0, normalized.Length - 2);
            return double.TryParse(
                normalized,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out result);
        }

        private static double GetDouble(XElement el, string attributeName, double defaultValue = 0)
        {
            var attr = el.Attribute(attributeName);
            if (attr != null && double.TryParse(attr.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }
            return defaultValue;
        }
    }
}
