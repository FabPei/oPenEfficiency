using System;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Xml.Linq;

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
                        var groupTransform = el.Attribute("transform")?.Value;
                        if (!string.IsNullOrEmpty(groupTransform))
                        {
                            var childGroup = new GeometryGroup();
                            childGroup.Transform = ParseTransform(groupTransform);
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
                        var transformAttr = el.Attribute("transform")?.Value;
                        if (!string.IsNullOrEmpty(transformAttr))
                        {
                            geom.Transform = ParseTransform(transformAttr);
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

                // Fallback for other transform types (translate, rotate, etc.)
                // TransformConverter is generally good at handling these invariantly
                return (Transform)new TransformConverter().ConvertFromInvariantString(transformStr);
            }
            catch
            {
                return null;
            }
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
