using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Media;
using System.Xml;
using System.Xml.Linq;

/// <summary>
/// Flattens the deliberately small SVG subset emitted by Icon Composer into filled paths.
/// Strokes and binary luminance masks are resolved with WPF geometry so the result does not
/// depend on mask, clip, filter, metadata, or identifier support in the production SVG reader.
/// </summary>
public static class IconComposerSvgExporter
{
    private const double GeometryTolerance = 0.001d;
    private const string SvgNamespace = "http://www.w3.org/2000/svg";

    private static readonly Regex NumberRegex = new Regex(
        @"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?",
        RegexOptions.CultureInvariant);

    private static readonly Regex TransformRegex = new Regex(
        @"([A-Za-z]+)\s*\(([^)]*)\)",
        RegexOptions.CultureInvariant);

    private static readonly HashSet<string> CommonAttributes = new HashSet<string>(StringComparer.Ordinal)
    {
        "id", "transform", "mask", "style", "fill", "stroke", "color", "fill-rule",
        "stroke-width", "stroke-linecap", "stroke-linejoin", "stroke-miterlimit",
        "stroke-dasharray", "stroke-dashoffset", "opacity", "fill-opacity", "stroke-opacity",
        "data-oe-role", "data-oe-accent-role", "data-oe-knockout", "data-oe-layer",
        "data-oe-part", "data-oe-composer"
    };

    private static readonly HashSet<string> StyleProperties = new HashSet<string>(StringComparer.Ordinal)
    {
        "fill", "stroke", "color", "fill-rule", "stroke-width", "stroke-linecap",
        "stroke-linejoin", "stroke-miterlimit", "stroke-dasharray", "stroke-dashoffset",
        "opacity", "fill-opacity", "stroke-opacity"
    };

    private static readonly Dictionary<string, string> ApprovedAccentRoles =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "ActionBlue", "ActionBlue" },
            { "CreateGreen", "CreateGreen" },
            { "SuccessGreen", "SuccessGreen" },
            { "AttentionYellow", "AttentionYellow" },
            { "DestructiveRed", "DestructiveRed" },
            { "ContentColor", "ContentColor" }
        };

    /// <summary>
    /// Exports a composer SVG as a deterministic SVG containing only explicit filled paths.
    /// </summary>
    public static string Export(string svg)
    {
        if (svg == null)
            throw new ArgumentNullException("svg");

        XDocument document;
        XmlReaderSettings settings = new XmlReaderSettings();
        settings.DtdProcessing = DtdProcessing.Prohibit;
        settings.XmlResolver = null;
        settings.IgnoreWhitespace = false;
        settings.MaxCharactersFromEntities = 0L;

        using (StringReader textReader = new StringReader(svg))
        using (XmlReader reader = XmlReader.Create(textReader, settings))
        {
            document = XDocument.Load(reader, LoadOptions.None);
        }

        if (document.Root == null || !IsSvgElement(document.Root, "svg"))
            throw new NotSupportedException("The document root must be an SVG <svg> element.");

        string viewBoxText = RequiredAttribute(document.Root, "viewBox").Trim();
        double[] rootViewBox = ParseNumberList(viewBoxText, 4, "viewBox");
        if (rootViewBox[2] <= 0d || rootViewBox[3] <= 0d)
            throw new NotSupportedException("The root SVG viewBox must have positive width and height.");

        Style defaults = Style.CreateDefaults();
        Semantic defaultsSemantic = new Semantic();
        HashSet<MaskDefinition> activeMasks = new HashSet<MaskDefinition>();
        List<PaintPath> paths = RenderSvg(
            document.Root,
            defaults,
            defaultsSemantic,
            true,
            activeMasks);

        StringBuilder output = new StringBuilder();
        output.Append("<svg xmlns=\"");
        output.Append(SvgNamespace);
        output.Append("\" viewBox=\"");
        output.Append(EscapeAttribute(viewBoxText));
        output.Append("\">");

        int emittedPaths = 0;
        int emittedBasePaths = 0;
        for (int index = 0; index < paths.Count; index++)
        {
            PaintPath path = paths[index];
            string data;
            string fillRule;
            if (!TrySerializeGeometry(path.Geometry, out data, out fillRule))
                continue;
            emittedPaths++;
            if (!path.Semantic.IsAccent)
                emittedBasePaths++;

            output.Append("\n  <path d=\"");
            output.Append(EscapeAttribute(data));
            output.Append("\" fill=\"");
            output.Append(EscapeAttribute(GetOutputPaint(path)));
            output.Append("\"");
            output.Append(" fill-rule=\"");
            output.Append(fillRule);
            output.Append("\"");
            if (path.Semantic.IsAccent)
            {
                output.Append(" data-oe-role=\"accent\" data-oe-accent-role=\"");
                output.Append(EscapeAttribute(path.Semantic.AccentRole));
                output.Append("\"");
            }
            else if (path.Semantic.IsExplicitBase)
            {
                output.Append(" data-oe-role=\"base\"");
            }
            output.Append("/>");
        }

        if (emittedPaths == 0 || emittedBasePaths == 0)
            throw new NotSupportedException("The production renderer requires at least one visible neutral base path.");
        output.Append("\n</svg>");
        return output.ToString();
    }

    private static List<PaintPath> RenderSvg(
        XElement element,
        Style inheritedStyle,
        Semantic inheritedSemantic,
        bool isRoot,
        HashSet<MaskDefinition> activeMasks)
    {
        ValidateElementAttributes(element, "svg");
        ValidateSvgViewport(element, isRoot);
        Style style = ResolveStyle(inheritedStyle, element);
        Semantic semantic = ResolveSemantic(inheritedSemantic, element);
        MaskScope scope = new MaskScope();
        CollectMaskDefinitions(element, style, semantic, scope);

        List<PaintPath> result = RenderContentChildren(
            element,
            style,
            semantic,
            scope,
            activeMasks);

        result = ApplyContentMask(result, element, scope, activeMasks);

        if (!isRoot)
        {
            Rect viewportClip;
            Matrix viewport = ParseNestedSvgViewport(element, out viewportClip);
            TransformPaintPaths(result, viewport);
            result = ClipPaintPaths(result, new RectangleGeometry(viewportClip));
            TransformPaintPaths(result, ParseTransformAttribute(element));
        }
        else
        {
            TransformPaintPaths(result, ParseTransformAttribute(element));
            double[] viewBox = ParseNumberList(RequiredAttribute(element, "viewBox"), 4, "viewBox");
            result = ClipPaintPaths(
                result,
                new RectangleGeometry(new Rect(viewBox[0], viewBox[1], viewBox[2], viewBox[3])));
        }
        return result;
    }

    private static void ValidateSvgViewport(XElement element, bool isRoot)
    {
        string overflow = AttributeValue(element, "overflow");
        if (overflow != null && !string.Equals(overflow.Trim(), "hidden", StringComparison.Ordinal))
            throw new NotSupportedException("Only clipped SVG viewports (overflow=\"hidden\") are supported.");
        if (!isRoot)
            return;

        // The production path factory fits the viewBox uniformly into its 24x24 canvas.
        // Do not silently erase source viewport instructions that demand another layout.
        string preserve = AttributeValue(element, "preserveAspectRatio");
        if (preserve != null)
        {
            preserve = Regex.Replace(preserve.Trim(), @"\s+", " ");
            if (preserve != "xMidYMid" && preserve != "xMidYMid meet")
                throw new NotSupportedException("The root SVG must use preserveAspectRatio=\"xMidYMid meet\".");
        }
        string widthText = AttributeValue(element, "width");
        string heightText = AttributeValue(element, "height");
        double width = widthText == null ? 0d : ParseLength(widthText, "width");
        double height = heightText == null ? 0d : ParseLength(heightText, "height");
        if ((widthText != null && width <= 0d) || (heightText != null && height <= 0d))
            throw new NotSupportedException("The root SVG viewport must have positive width and height.");
        if (widthText != null && heightText != null)
        {
            double[] viewBox = ParseNumberList(RequiredAttribute(element, "viewBox"), 4, "viewBox");
            double ratioDifference = Math.Log(width) - Math.Log(height) - Math.Log(viewBox[2]) + Math.Log(viewBox[3]);
            if (Math.Abs(ratioDifference) > 1e-9d)
                throw new NotSupportedException("The root SVG viewport aspect ratio must match its viewBox.");
        }
    }

    private static List<PaintPath> RenderContentChildren(
        XElement parent,
        Style inheritedStyle,
        Semantic inheritedSemantic,
        MaskScope scope,
        HashSet<MaskDefinition> activeMasks)
    {
        List<PaintPath> result = new List<PaintPath>();
        foreach (XElement child in parent.Elements())
        {
            string name = SvgName(child);
            if (name == "defs" || name == "metadata" || name == "title" || name == "desc")
                continue;

            List<PaintPath> childPaths;
            if (name == "g")
            {
                childPaths = RenderGroup(child, inheritedStyle, inheritedSemantic, scope, activeMasks);
            }
            else if (name == "svg")
            {
                childPaths = RenderSvg(child, inheritedStyle, inheritedSemantic, false, activeMasks);
            }
            else if (IsDrawableName(name))
            {
                childPaths = RenderDrawable(child, inheritedStyle, inheritedSemantic, scope, activeMasks);
            }
            else
            {
                throw UnsupportedElement(child);
            }

            result.AddRange(childPaths);
        }
        return result;
    }

    private static List<PaintPath> RenderGroup(
        XElement element,
        Style inheritedStyle,
        Semantic inheritedSemantic,
        MaskScope scope,
        HashSet<MaskDefinition> activeMasks)
    {
        ValidateElementAttributes(element, "g");
        Style style = ResolveStyle(inheritedStyle, element);
        Semantic semantic = ResolveSemantic(inheritedSemantic, element);
        List<PaintPath> result = RenderContentChildren(element, style, semantic, scope, activeMasks);
        result = ApplyContentMask(result, element, scope, activeMasks);
        TransformPaintPaths(result, ParseTransformAttribute(element));
        return result;
    }

    private static List<PaintPath> RenderDrawable(
        XElement element,
        Style inheritedStyle,
        Semantic inheritedSemantic,
        MaskScope scope,
        HashSet<MaskDefinition> activeMasks)
    {
        string name = SvgName(element);
        ValidateElementAttributes(element, name);
        Style style = ResolveStyle(inheritedStyle, element);
        Semantic semantic = ResolveSemantic(inheritedSemantic, element);
        Geometry source = CreateGeometry(element, name, style.FillRule);
        List<PaintPath> result = new List<PaintPath>();

        if (!IsNone(style.Fill))
        {
            ValidateRegularPaint(style.Fill, style.Color, element, "fill");
            result.Add(new PaintPath(CloneGeometry(source), style.Fill, style.Color, semantic.Clone()));
        }

        if (!IsNone(style.Stroke) && style.StrokeWidth > 0d)
        {
            ValidateRegularPaint(style.Stroke, style.Color, element, "stroke");
            Geometry widened = WidenGeometry(source, style);
            result.Add(new PaintPath(widened, style.Stroke, style.Color, semantic.Clone()));
        }

        result = ApplyContentMask(result, element, scope, activeMasks);
        TransformPaintPaths(result, ParseTransformAttribute(element));
        return result;
    }

    private static List<PaintPath> ApplyContentMask(
        List<PaintPath> paths,
        XElement element,
        MaskScope scope,
        HashSet<MaskDefinition> activeMasks)
    {
        string reference = AttributeValue(element, "mask");
        if (string.IsNullOrWhiteSpace(reference))
            return paths;

        Geometry mask = ResolveMaskGeometry(reference, scope, activeMasks);
        List<PaintPath> result = new List<PaintPath>();
        for (int index = 0; index < paths.Count; index++)
        {
            Geometry clipped = Combine(paths[index].Geometry, mask, GeometryCombineMode.Intersect);
            if (!IsEmptyGeometry(clipped))
            {
                PaintPath source = paths[index];
                result.Add(new PaintPath(clipped, source.Paint, source.Color, source.Semantic));
            }
        }
        return result;
    }

    private static void CollectMaskDefinitions(
        XElement svg,
        Style svgStyle,
        Semantic svgSemantic,
        MaskScope scope)
    {
        foreach (XElement child in svg.Elements())
            CollectMaskDefinitionsFromNode(child, svgStyle, svgSemantic, scope);
    }

    private static void CollectMaskDefinitionsFromNode(
        XElement element,
        Style inheritedStyle,
        Semantic inheritedSemantic,
        MaskScope scope)
    {
        string name = SvgName(element);
        if (name == "metadata" || name == "title" || name == "desc" || name == "svg")
            return;

        if (name == "defs")
        {
            ValidateElementAttributes(element, "defs");
            Style definitionStyle = ResolveStyle(inheritedStyle, element);
            Semantic definitionSemantic = ResolveSemantic(inheritedSemantic, element);
            foreach (XElement definition in element.Elements())
            {
                if (SvgName(definition) != "mask")
                    throw new NotSupportedException(
                        "Unsupported <" + SvgName(definition) + "> definition; only local <mask> definitions are supported.");
                RegisterMaskDefinition(definition, definitionStyle, definitionSemantic, scope);
            }
            return;
        }

        if (name == "g")
        {
            ValidateElementAttributes(element, "g");
            Style style = ResolveStyle(inheritedStyle, element);
            Semantic semantic = ResolveSemantic(inheritedSemantic, element);
            foreach (XElement child in element.Elements())
                CollectMaskDefinitionsFromNode(child, style, semantic, scope);
        }
    }

    private static void RegisterMaskDefinition(
        XElement element,
        Style inheritedStyle,
        Semantic inheritedSemantic,
        MaskScope scope)
    {
        ValidateElementAttributes(element, "mask");
        string id = RequiredAttribute(element, "id").Trim();
        if (id.Length == 0)
            throw new NotSupportedException("A mask id cannot be empty.");
        if (scope.Masks.ContainsKey(id))
            throw new NotSupportedException("Duplicate local mask id '#" + id + "'.");

        string maskUnits = AttributeValue(element, "maskUnits");
        if (maskUnits == null)
            throw new NotSupportedException("Mask '#" + id + "' must declare maskUnits=\"userSpaceOnUse\".");
        if (!string.Equals(maskUnits.Trim(), "userSpaceOnUse", StringComparison.Ordinal))
            throw new NotSupportedException("Mask '#" + id + "' uses unsupported maskUnits '" + maskUnits + "'.");

        string contentUnits = AttributeValue(element, "maskContentUnits");
        if (contentUnits != null &&
            !string.Equals(contentUnits.Trim(), "userSpaceOnUse", StringComparison.Ordinal))
        {
            throw new NotSupportedException(
                "Mask '#" + id + "' uses unsupported maskContentUnits '" + contentUnits + "'.");
        }

        ValidateOptionalLengths(element, new string[] { "x", "y", "width", "height" });
        Style style = ResolveStyle(inheritedStyle, element);
        Semantic semantic = ResolveSemantic(inheritedSemantic, element);
        MaskDefinition definition = new MaskDefinition(element, style, semantic, scope, id);
        scope.Masks.Add(id, definition);
    }

    private static Geometry ResolveMaskGeometry(
        string reference,
        MaskScope scope,
        HashSet<MaskDefinition> activeMasks)
    {
        string id = ParseLocalUrl(reference, "mask");
        MaskDefinition definition;
        if (!scope.Masks.TryGetValue(id, out definition))
            throw new NotSupportedException("Unresolved local mask reference '#" + id + "'.");
        if (activeMasks.Contains(definition))
            throw new NotSupportedException("Cyclic local mask reference involving '#" + id + "'.");

        activeMasks.Add(definition);
        try
        {
            List<MaskOperation> operations = RenderMaskChildren(
                definition.Element,
                definition.Style,
                definition.Semantic,
                definition.Scope,
                activeMasks);

            Geometry result = null;
            for (int index = 0; index < operations.Count; index++)
            {
                MaskOperation operation = operations[index];
                if (operation.IsWhite)
                    result = result == null
                        ? CloneGeometry(operation.Geometry)
                        : Combine(result, operation.Geometry, GeometryCombineMode.Union);
                else if (result != null)
                    result = Combine(result, operation.Geometry, GeometryCombineMode.Exclude);
            }

            if (result == null)
                return Geometry.Empty;

            bool hasX = AttributeValue(definition.Element, "x") != null;
            bool hasY = AttributeValue(definition.Element, "y") != null;
            bool hasWidth = AttributeValue(definition.Element, "width") != null;
            bool hasHeight = AttributeValue(definition.Element, "height") != null;
            if (hasX || hasY || hasWidth || hasHeight)
            {
                if (!hasX || !hasY || !hasWidth || !hasHeight)
                {
                    throw new NotSupportedException(
                        "Mask '#" + definition.Id +
                        "' must provide x, y, width, and height together because percentage defaults are unsupported.");
                }
                double x = RequiredLength(definition.Element, "x");
                double y = RequiredLength(definition.Element, "y");
                double width = RequiredLength(definition.Element, "width");
                double height = RequiredLength(definition.Element, "height");
                if (width < 0d || height < 0d)
                    throw new NotSupportedException("Mask region width and height cannot be negative.");
                Geometry region = new RectangleGeometry(new Rect(x, y, width, height));
                result = Combine(result, region, GeometryCombineMode.Intersect);
            }

            string nestedReference = AttributeValue(definition.Element, "mask");
            if (!string.IsNullOrWhiteSpace(nestedReference))
            {
                Geometry nested = ResolveMaskGeometry(nestedReference, scope, activeMasks);
                result = Combine(result, nested, GeometryCombineMode.Intersect);
            }

            return TransformGeometry(result, ParseTransformAttribute(definition.Element));
        }
        finally
        {
            activeMasks.Remove(definition);
        }
    }

    private static List<MaskOperation> RenderMaskChildren(
        XElement parent,
        Style inheritedStyle,
        Semantic inheritedSemantic,
        MaskScope scope,
        HashSet<MaskDefinition> activeMasks)
    {
        List<MaskOperation> result = new List<MaskOperation>();
        foreach (XElement child in parent.Elements())
        {
            string name = SvgName(child);
            if (name == "metadata" || name == "title" || name == "desc")
                continue;
            if (name == "defs")
                throw new NotSupportedException("Nested <defs> inside a mask is unsupported.");
            if (name == "g")
                result.AddRange(RenderMaskGroup(child, inheritedStyle, inheritedSemantic, scope, activeMasks));
            else if (IsDrawableName(name))
                result.AddRange(RenderMaskDrawable(child, inheritedStyle, inheritedSemantic, scope, activeMasks));
            else
                throw UnsupportedElement(child);
        }
        return result;
    }

    private static List<MaskOperation> RenderMaskGroup(
        XElement element,
        Style inheritedStyle,
        Semantic inheritedSemantic,
        MaskScope scope,
        HashSet<MaskDefinition> activeMasks)
    {
        ValidateElementAttributes(element, "g");
        Style style = ResolveStyle(inheritedStyle, element);
        Semantic semantic = ResolveSemantic(inheritedSemantic, element);
        List<MaskOperation> result = RenderMaskChildren(element, style, semantic, scope, activeMasks);
        result = ApplyMaskToMaskOperations(result, element, scope, activeMasks);
        TransformMaskOperations(result, ParseTransformAttribute(element));
        return result;
    }

    private static List<MaskOperation> RenderMaskDrawable(
        XElement element,
        Style inheritedStyle,
        Semantic inheritedSemantic,
        MaskScope scope,
        HashSet<MaskDefinition> activeMasks)
    {
        string name = SvgName(element);
        ValidateElementAttributes(element, name);
        Style style = ResolveStyle(inheritedStyle, element);
        ResolveSemantic(inheritedSemantic, element);
        Geometry source = CreateGeometry(element, name, style.FillRule);
        List<MaskOperation> result = new List<MaskOperation>();

        if (!IsNone(style.Fill))
        {
            bool fillIsWhite = ResolveMaskPaint(style.Fill, style.Color, element, "fill");
            result.Add(new MaskOperation(CloneGeometry(source), fillIsWhite));
        }
        if (!IsNone(style.Stroke) && style.StrokeWidth > 0d)
        {
            bool strokeIsWhite = ResolveMaskPaint(style.Stroke, style.Color, element, "stroke");
            result.Add(new MaskOperation(WidenGeometry(source, style), strokeIsWhite));
        }

        result = ApplyMaskToMaskOperations(result, element, scope, activeMasks);
        TransformMaskOperations(result, ParseTransformAttribute(element));
        return result;
    }

    private static List<MaskOperation> ApplyMaskToMaskOperations(
        List<MaskOperation> operations,
        XElement element,
        MaskScope scope,
        HashSet<MaskDefinition> activeMasks)
    {
        string reference = AttributeValue(element, "mask");
        if (string.IsNullOrWhiteSpace(reference))
            return operations;

        Geometry mask = ResolveMaskGeometry(reference, scope, activeMasks);
        List<MaskOperation> result = new List<MaskOperation>();
        for (int index = 0; index < operations.Count; index++)
        {
            Geometry clipped = Combine(operations[index].Geometry, mask, GeometryCombineMode.Intersect);
            if (!IsEmptyGeometry(clipped))
                result.Add(new MaskOperation(clipped, operations[index].IsWhite));
        }
        return result;
    }

    private static Geometry CreateGeometry(XElement element, string name, string fillRule)
    {
        Geometry geometry;
        if (name == "path")
        {
            string data = RequiredAttribute(element, "d").Trim();
            if (Regex.IsMatch(data, @"\AF[01](?:\s|,|[Mm])", RegexOptions.CultureInvariant))
                throw new NotSupportedException("WPF F0/F1 path prefixes are not valid source SVG path data.");
            if (data.Length == 0)
                throw new NotSupportedException("A <path> d attribute cannot be empty.");
            try
            {
                geometry = Geometry.Parse(data);
            }
            catch (Exception exception)
            {
                throw new NotSupportedException("Invalid SVG path data: " + exception.Message, exception);
            }
        }
        else if (name == "rect")
        {
            double x = OptionalLength(element, "x", 0d);
            double y = OptionalLength(element, "y", 0d);
            double width = RequiredLength(element, "width");
            double height = RequiredLength(element, "height");
            if (width < 0d || height < 0d)
                throw new NotSupportedException("Rectangle width and height cannot be negative.");
            string rxText = AttributeValue(element, "rx");
            string ryText = AttributeValue(element, "ry");
            double rx = rxText == null ? 0d : ParseLength(rxText, "rx");
            double ry = ryText == null ? rx : ParseLength(ryText, "ry");
            if (rxText == null && ryText != null)
                rx = ry;
            if (rx < 0d || ry < 0d)
                throw new NotSupportedException("Rectangle corner radii cannot be negative.");
            rx = Math.Min(rx, width / 2d);
            ry = Math.Min(ry, height / 2d);
            geometry = new RectangleGeometry(new Rect(x, y, width, height), rx, ry);
        }
        else if (name == "circle")
        {
            double radius = RequiredLength(element, "r");
            if (radius < 0d)
                throw new NotSupportedException("Circle radius cannot be negative.");
            geometry = new EllipseGeometry(
                new Point(OptionalLength(element, "cx", 0d), OptionalLength(element, "cy", 0d)),
                radius,
                radius);
        }
        else if (name == "ellipse")
        {
            double rx = RequiredLength(element, "rx");
            double ry = RequiredLength(element, "ry");
            if (rx < 0d || ry < 0d)
                throw new NotSupportedException("Ellipse radii cannot be negative.");
            geometry = new EllipseGeometry(
                new Point(OptionalLength(element, "cx", 0d), OptionalLength(element, "cy", 0d)),
                rx,
                ry);
        }
        else if (name == "line")
        {
            geometry = new LineGeometry(
                new Point(OptionalLength(element, "x1", 0d), OptionalLength(element, "y1", 0d)),
                new Point(OptionalLength(element, "x2", 0d), OptionalLength(element, "y2", 0d)));
        }
        else if (name == "polygon" || name == "polyline")
        {
            double[] values = ParseFlexibleNumberList(RequiredAttribute(element, "points"), "points");
            if (values.Length < 2 || values.Length % 2 != 0)
                throw new NotSupportedException("The points attribute must contain complete coordinate pairs.");
            StreamGeometry stream = new StreamGeometry();
            using (StreamGeometryContext context = stream.Open())
            {
                context.BeginFigure(new Point(values[0], values[1]), true, name == "polygon");
                for (int index = 2; index < values.Length; index += 2)
                    context.LineTo(new Point(values[index], values[index + 1]), true, false);
            }
            geometry = stream;
        }
        else
        {
            throw UnsupportedElement(element);
        }

        PathGeometry pathGeometry = PathGeometry.CreateFromGeometry(geometry);
        pathGeometry.FillRule = fillRule == "evenodd" ? FillRule.EvenOdd : FillRule.Nonzero;
        return pathGeometry;
    }

    private static Geometry WidenGeometry(Geometry geometry, Style style)
    {
        Pen pen = new Pen(Brushes.Black, style.StrokeWidth);
        pen.StartLineCap = style.LineCap;
        pen.EndLineCap = style.LineCap;
        pen.DashCap = style.LineCap;
        pen.LineJoin = style.LineJoin;
        pen.MiterLimit = style.MiterLimit;

        if (style.DashArray != null)
        {
            DoubleCollection dashes = new DoubleCollection();
            for (int index = 0; index < style.DashArray.Count; index++)
                dashes.Add(style.DashArray[index] / style.StrokeWidth);
            if (dashes.Count % 2 != 0)
            {
                int originalCount = dashes.Count;
                for (int index = 0; index < originalCount; index++)
                    dashes.Add(dashes[index]);
            }
            pen.DashStyle = new DashStyle(dashes, style.DashOffset / style.StrokeWidth);
        }

        return geometry.GetWidenedPathGeometry(pen, GeometryTolerance, ToleranceType.Absolute);
    }

    private static Style ResolveStyle(Style inherited, XElement element)
    {
        Style result = inherited.Clone();
        string[] properties = new string[]
        {
            "fill", "stroke", "color", "fill-rule", "stroke-width", "stroke-linecap",
            "stroke-linejoin", "stroke-miterlimit", "stroke-dasharray", "stroke-dashoffset",
            "opacity", "fill-opacity", "stroke-opacity"
        };

        for (int index = 0; index < properties.Length; index++)
        {
            string value = AttributeValue(element, properties[index]);
            if (value != null)
                ApplyStyleProperty(result, properties[index], value, element);
        }

        string styleText = AttributeValue(element, "style");
        if (!string.IsNullOrWhiteSpace(styleText))
        {
            string[] declarations = styleText.Split(';');
            for (int index = 0; index < declarations.Length; index++)
            {
                string declaration = declarations[index].Trim();
                if (declaration.Length == 0)
                    continue;
                int separator = declaration.IndexOf(':');
                if (separator <= 0 || separator == declaration.Length - 1)
                    throw new NotSupportedException("Invalid style declaration '" + declaration + "'.");
                string property = declaration.Substring(0, separator).Trim().ToLowerInvariant();
                string value = declaration.Substring(separator + 1).Trim();
                if (!StyleProperties.Contains(property))
                    throw new NotSupportedException("Unsupported SVG style property '" + property + "'.");
                ApplyStyleProperty(result, property, value, element);
            }
        }

        return result;
    }

    private static void ApplyStyleProperty(Style style, string property, string value, XElement element)
    {
        value = value.Trim();
        if (string.Equals(value, "inherit", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("Explicit CSS inherit is unsupported; resolve the inherited source style before export.");

        if (property == "fill")
            style.Fill = NormalizePaint(value, element, property);
        else if (property == "stroke")
            style.Stroke = NormalizePaint(value, element, property);
        else if (property == "color")
            style.Color = NormalizePaint(value, element, property);
        else if (property == "fill-rule")
        {
            string normalized = value.ToLowerInvariant();
            if (normalized != "nonzero" && normalized != "evenodd")
                throw new NotSupportedException("Unsupported fill-rule '" + value + "'.");
            style.FillRule = normalized;
        }
        else if (property == "stroke-width")
        {
            style.StrokeWidth = ParseLength(value, property);
            if (style.StrokeWidth < 0d)
                throw new NotSupportedException("stroke-width cannot be negative.");
        }
        else if (property == "stroke-linecap")
        {
            string normalized = value.ToLowerInvariant();
            if (normalized == "butt") style.LineCap = PenLineCap.Flat;
            else if (normalized == "round") style.LineCap = PenLineCap.Round;
            else if (normalized == "square") style.LineCap = PenLineCap.Square;
            else throw new NotSupportedException("Unsupported stroke-linecap '" + value + "'.");
        }
        else if (property == "stroke-linejoin")
        {
            string normalized = value.ToLowerInvariant();
            if (normalized == "miter") style.LineJoin = PenLineJoin.Miter;
            else if (normalized == "round") style.LineJoin = PenLineJoin.Round;
            else if (normalized == "bevel") style.LineJoin = PenLineJoin.Bevel;
            else throw new NotSupportedException("Unsupported stroke-linejoin '" + value + "'.");
        }
        else if (property == "stroke-miterlimit")
        {
            style.MiterLimit = ParseLength(value, property);
            if (style.MiterLimit < 1d)
                throw new NotSupportedException("stroke-miterlimit must be at least 1.");
        }
        else if (property == "stroke-dasharray")
        {
            if (string.Equals(value, "none", StringComparison.OrdinalIgnoreCase))
            {
                style.DashArray = null;
            }
            else
            {
                double[] values = ParseFlexibleNumberList(value, property);
                if (values.Length == 0)
                    throw new NotSupportedException("stroke-dasharray cannot be empty.");
                double sum = 0d;
                style.DashArray = new List<double>();
                for (int index = 0; index < values.Length; index++)
                {
                    if (values[index] < 0d)
                        throw new NotSupportedException("stroke-dasharray values cannot be negative.");
                    style.DashArray.Add(values[index]);
                    sum += values[index];
                }
                if (sum == 0d)
                    throw new NotSupportedException("stroke-dasharray cannot contain only zeros.");
            }
        }
        else if (property == "stroke-dashoffset")
        {
            style.DashOffset = ParseLength(value, property);
        }
        else if (property == "opacity" || property == "fill-opacity" || property == "stroke-opacity")
        {
            double opacity = ParseLength(value, property);
            if (Math.Abs(opacity - 1d) > 0.000000001d)
                throw new NotSupportedException(property + " values other than 1 are unsupported.");
        }
    }

    private static Semantic ResolveSemantic(Semantic inherited, XElement element)
    {
        Semantic result = inherited.Clone();
        string role = AttributeValue(element, "data-oe-role");
        if (role != null)
        {
            string normalized = role.Trim().ToLowerInvariant();
            if (normalized == "accent")
            {
                result.IsAccent = true;
                result.IsExplicitBase = false;
                if (string.IsNullOrEmpty(result.AccentRole))
                    result.AccentRole = "ActionBlue";
            }
            else if (normalized == "base")
            {
                result.IsAccent = false;
                result.IsExplicitBase = true;
                result.AccentRole = null;
            }
            else
            {
                throw new NotSupportedException("Unsupported data-oe-role '" + role + "'.");
            }
        }

        string accentRole = AttributeValue(element, "data-oe-accent-role");
        if (accentRole != null)
        {
            if (!result.IsAccent)
                throw new NotSupportedException("data-oe-accent-role requires data-oe-role=\"accent\" in its inheritance chain.");
            string canonical;
            if (!ApprovedAccentRoles.TryGetValue(accentRole.Trim(), out canonical))
                throw new NotSupportedException("Unsupported data-oe-accent-role '" + accentRole + "'.");
            result.AccentRole = canonical;
        }

        if (result.IsAccent && string.IsNullOrEmpty(result.AccentRole))
            result.AccentRole = "ActionBlue";
        return result;
    }

    private static string NormalizePaint(string value, XElement element, string property)
    {
        string normalized = value.Trim();
        if (string.Equals(normalized, "none", StringComparison.OrdinalIgnoreCase))
            return "none";
        if (string.Equals(normalized, "currentColor", StringComparison.OrdinalIgnoreCase))
            return "currentColor";
        if (string.Equals(normalized, "black", StringComparison.OrdinalIgnoreCase))
            return "#000000";
        if (string.Equals(normalized, "white", StringComparison.OrdinalIgnoreCase))
            return "#FFFFFF";
        if (Regex.IsMatch(normalized, @"\A#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})\z", RegexOptions.CultureInvariant))
            return ExpandHex(normalized).ToUpperInvariant();
        throw new NotSupportedException(
            "Unsupported " + property + " paint '" + value + "' on <" + SvgName(element) + ">.");
    }

    private static void ValidateRegularPaint(string paint, string color, XElement element, string property)
    {
        if (paint == "currentColor")
        {
            if (color == "none")
                throw new NotSupportedException("currentColor cannot resolve from color=\"none\".");
            return;
        }
        if (!Regex.IsMatch(paint, @"\A#[0-9A-F]{6}\z", RegexOptions.CultureInvariant))
            throw new NotSupportedException("Unsupported " + property + " paint on <" + SvgName(element) + ">.");
    }

    private static bool ResolveMaskPaint(
        string paint,
        string color,
        XElement element,
        string property)
    {
        string resolved = paint == "currentColor" ? color : paint;
        if (resolved == "#FFFFFF")
            return true;
        if (resolved == "#000000")
            return false;
        throw new NotSupportedException(
            "Mask " + property + " paint on <" + SvgName(element) +
            "> must resolve to pure white or pure black, but resolved to '" + resolved + "'.");
    }

    private static string GetOutputPaint(PaintPath path)
    {
        if (!path.Semantic.IsAccent || path.Semantic.AccentRole != "ContentColor")
            return "currentColor";
        if (path.Paint == "currentColor")
        {
            if (Regex.IsMatch(path.Color, @"\A#[0-9A-F]{6}\z", RegexOptions.CultureInvariant))
                return path.Color;
            return "currentColor";
        }
        return path.Paint;
    }

    private static Matrix ParseTransformAttribute(XElement element)
    {
        string value = AttributeValue(element, "transform");
        if (string.IsNullOrWhiteSpace(value))
            return Matrix.Identity;
        return ParseTransform(value);
    }

    private static Matrix ParseTransform(string text)
    {
        Matrix result = Matrix.Identity;
        int consumed = 0;
        MatchCollection matches = TransformRegex.Matches(text);
        if (matches.Count == 0)
            throw new NotSupportedException("Invalid SVG transform '" + text + "'.");

        for (int index = 0; index < matches.Count; index++)
        {
            Match match = matches[index];
            ValidateTransformSeparator(text.Substring(consumed, match.Index - consumed), text);
            consumed = match.Index + match.Length;
            string name = match.Groups[1].Value;
            double[] values = ParseFlexibleNumberList(match.Groups[2].Value, "transform " + name);
            Matrix operation = Matrix.Identity;

            if (string.Equals(name, "matrix", StringComparison.OrdinalIgnoreCase))
            {
                RequireCount(values, 6, name);
                operation = new Matrix(values[0], values[1], values[2], values[3], values[4], values[5]);
            }
            else if (string.Equals(name, "translate", StringComparison.OrdinalIgnoreCase))
            {
                if (values.Length != 1 && values.Length != 2)
                    throw TransformCountException(name, "one or two");
                operation.Translate(values[0], values.Length == 2 ? values[1] : 0d);
            }
            else if (string.Equals(name, "scale", StringComparison.OrdinalIgnoreCase))
            {
                if (values.Length != 1 && values.Length != 2)
                    throw TransformCountException(name, "one or two");
                operation.Scale(values[0], values.Length == 2 ? values[1] : values[0]);
            }
            else if (string.Equals(name, "rotate", StringComparison.OrdinalIgnoreCase))
            {
                if (values.Length == 1)
                    operation.Rotate(values[0]);
                else if (values.Length == 3)
                    operation.RotateAt(values[0], values[1], values[2]);
                else
                    throw TransformCountException(name, "one or three");
            }
            else if (string.Equals(name, "skewX", StringComparison.OrdinalIgnoreCase))
            {
                RequireCount(values, 1, name);
                operation.Skew(values[0], 0d);
            }
            else if (string.Equals(name, "skewY", StringComparison.OrdinalIgnoreCase))
            {
                RequireCount(values, 1, name);
                operation.Skew(0d, values[0]);
            }
            else
            {
                throw new NotSupportedException("Unsupported SVG transform function '" + name + "'.");
            }

            // SVG transform lists are matrix-multiplied in textual order. WPF uses row vectors,
            // therefore Prepend gives the corresponding SVG point mapping.
            result.Prepend(operation);
        }

        ValidateTransformSeparator(text.Substring(consumed), text);
        if (!IsFiniteMatrix(result))
            throw new NotSupportedException("SVG transform produced a non-finite matrix.");
        return result;
    }

    private static Matrix ParseNestedSvgViewport(XElement element, out Rect viewportClip)
    {
        double x = OptionalLength(element, "x", 0d);
        double y = OptionalLength(element, "y", 0d);
        string widthText = AttributeValue(element, "width");
        string heightText = AttributeValue(element, "height");
        if (widthText == null || heightText == null)
        {
            throw new NotSupportedException(
                "A nested SVG must declare numeric width and height so its viewport clipping can be preserved.");
        }
        double width = ParseLength(widthText, "width");
        double height = ParseLength(heightText, "height");
        if (width <= 0d || height <= 0d)
            throw new NotSupportedException("A nested SVG viewport must have positive width and height.");
        viewportClip = new Rect(x, y, width, height);

        string viewBoxText = AttributeValue(element, "viewBox");
        if (viewBoxText == null)
        {
            if (AttributeValue(element, "preserveAspectRatio") != null)
            {
                throw new NotSupportedException(
                    "A nested SVG without a viewBox cannot declare preserveAspectRatio.");
            }
            Matrix translation = Matrix.Identity;
            translation.Translate(x, y);
            return translation;
        }

        double[] viewBox = ParseNumberList(viewBoxText, 4, "viewBox");
        if (viewBox[2] <= 0d || viewBox[3] <= 0d)
            throw new NotSupportedException("A nested SVG viewBox must have positive width and height.");

        string preserve = AttributeValue(element, "preserveAspectRatio");
        if (string.IsNullOrWhiteSpace(preserve))
            preserve = "xMidYMid meet";
        preserve = Regex.Replace(preserve.Trim(), @"\s+", " ");

        double scaleX = width / viewBox[2];
        double scaleY = height / viewBox[3];
        double offsetX = x;
        double offsetY = y;
        if (string.Equals(preserve, "none", StringComparison.Ordinal))
        {
            offsetX -= viewBox[0] * scaleX;
            offsetY -= viewBox[1] * scaleY;
        }
        else if (string.Equals(preserve, "xMidYMid meet", StringComparison.Ordinal))
        {
            double scale = Math.Min(scaleX, scaleY);
            offsetX += (width - viewBox[2] * scale) / 2d - viewBox[0] * scale;
            offsetY += (height - viewBox[3] * scale) / 2d - viewBox[1] * scale;
            scaleX = scale;
            scaleY = scale;
        }
        else
        {
            throw new NotSupportedException("Unsupported preserveAspectRatio '" + preserve + "'.");
        }

        return new Matrix(scaleX, 0d, 0d, scaleY, offsetX, offsetY);
    }

    private static List<PaintPath> ClipPaintPaths(List<PaintPath> paths, Geometry clip)
    {
        List<PaintPath> result = new List<PaintPath>();
        for (int index = 0; index < paths.Count; index++)
        {
            Geometry clipped = Combine(paths[index].Geometry, clip, GeometryCombineMode.Intersect);
            if (!IsEmptyGeometry(clipped))
            {
                PaintPath source = paths[index];
                result.Add(new PaintPath(clipped, source.Paint, source.Color, source.Semantic));
            }
        }
        return result;
    }

    private static void TransformPaintPaths(List<PaintPath> paths, Matrix transform)
    {
        if (transform.IsIdentity)
            return;
        for (int index = 0; index < paths.Count; index++)
            paths[index].Geometry = TransformGeometry(paths[index].Geometry, transform);
    }

    private static void TransformMaskOperations(List<MaskOperation> operations, Matrix transform)
    {
        if (transform.IsIdentity)
            return;
        for (int index = 0; index < operations.Count; index++)
            operations[index].Geometry = TransformGeometry(operations[index].Geometry, transform);
    }

    private static Geometry TransformGeometry(Geometry geometry, Matrix transform)
    {
        if (transform.IsIdentity || IsEmptyGeometry(geometry))
            return geometry;
        GeometryGroup wrapper = new GeometryGroup();
        wrapper.FillRule = FillRule.Nonzero;
        wrapper.Children.Add(geometry);
        wrapper.Transform = new MatrixTransform(transform);
        return wrapper;
    }

    private static Geometry Combine(Geometry first, Geometry second, GeometryCombineMode mode)
    {
        return Geometry.Combine(
            first,
            second,
            mode,
            null,
            GeometryTolerance,
            ToleranceType.Absolute);
    }

    private static Geometry CloneGeometry(Geometry geometry)
    {
        return geometry == null ? null : geometry.Clone();
    }

    private static bool IsEmptyGeometry(Geometry geometry)
    {
        return geometry == null || geometry == Geometry.Empty || geometry.Bounds.IsEmpty;
    }

    private static bool TrySerializeGeometry(Geometry geometry, out string data, out string fillRule)
    {
        data = null;
        fillRule = "nonzero";
        if (IsEmptyGeometry(geometry))
            return false;

        // Resolve the actual filled region before serialization.  The production WPF reader
        // treats unprefixed path data as EvenOdd and does not apply SVG fill-rule metadata, so a
        // source Nonzero/self-overlapping path cannot safely be emitted as its original contours.
        // GetOutlinedPathGeometry converts the filled result (including Boolean operations) into
        // boundary contours whose region is equivalent under EvenOdd in both SVG and WPF.
        PathGeometry path = geometry.GetOutlinedPathGeometry(
            GeometryTolerance,
            ToleranceType.Absolute);
        if (path.Figures.Count == 0 || path.Bounds.IsEmpty)
            return false;
        path.FillRule = FillRule.EvenOdd;
        fillRule = "evenodd";
        string serialized = path.ToString(CultureInfo.InvariantCulture);
        if (serialized.StartsWith("F0", StringComparison.Ordinal))
        {
            fillRule = "evenodd";
            serialized = serialized.Substring(2);
        }
        else if (serialized.StartsWith("F1", StringComparison.Ordinal))
        {
            fillRule = "nonzero";
            serialized = serialized.Substring(2);
        }
        serialized = serialized.TrimStart();
        if (serialized.Length == 0)
            return false;
        data = serialized;
        return true;
    }

    private static void ValidateElementAttributes(XElement element, string name)
    {
        HashSet<string> allowed = new HashSet<string>(CommonAttributes, StringComparer.Ordinal);
        if (name == "svg")
        {
            AddAllowed(allowed, "x", "y", "width", "height", "viewBox", "preserveAspectRatio", "overflow");
        }
        else if (name == "path")
        {
            AddAllowed(allowed, "d");
        }
        else if (name == "rect")
        {
            AddAllowed(allowed, "x", "y", "width", "height", "rx", "ry");
        }
        else if (name == "circle")
        {
            AddAllowed(allowed, "cx", "cy", "r");
        }
        else if (name == "ellipse")
        {
            AddAllowed(allowed, "cx", "cy", "rx", "ry");
        }
        else if (name == "line")
        {
            AddAllowed(allowed, "x1", "y1", "x2", "y2");
        }
        else if (name == "polygon" || name == "polyline")
        {
            AddAllowed(allowed, "points");
        }
        else if (name == "mask")
        {
            AddAllowed(allowed, "maskUnits", "maskContentUnits", "x", "y", "width", "height");
        }

        foreach (XAttribute attribute in element.Attributes())
        {
            if (attribute.IsNamespaceDeclaration)
                continue;
            if (attribute.Name.NamespaceName.Length != 0)
                throw new NotSupportedException("Unsupported namespaced attribute '" + attribute.Name + "'.");
            if (!allowed.Contains(attribute.Name.LocalName))
            {
                throw new NotSupportedException(
                    "Unsupported attribute '" + attribute.Name.LocalName + "' on <" + name + ">.");
            }
        }

        string filter = AttributeValue(element, "filter");
        string clip = AttributeValue(element, "clip-path");
        if (filter != null || clip != null)
            throw new NotSupportedException("Filters and clip paths are unsupported in production exports.");
    }

    private static void AddAllowed(HashSet<string> set, params string[] values)
    {
        for (int index = 0; index < values.Length; index++)
            set.Add(values[index]);
    }

    private static string SvgName(XElement element)
    {
        string namespaceName = element.Name.NamespaceName;
        if (namespaceName.Length != 0 && namespaceName != SvgNamespace)
            throw new NotSupportedException("Unsupported element namespace '" + namespaceName + "'.");
        return element.Name.LocalName;
    }

    private static bool IsSvgElement(XElement element, string localName)
    {
        string namespaceName = element.Name.NamespaceName;
        return element.Name.LocalName == localName &&
            (namespaceName.Length == 0 || namespaceName == SvgNamespace);
    }

    private static bool IsDrawableName(string name)
    {
        return name == "path" || name == "rect" || name == "circle" || name == "ellipse" ||
            name == "line" || name == "polygon" || name == "polyline";
    }

    private static NotSupportedException UnsupportedElement(XElement element)
    {
        return new NotSupportedException("Unsupported SVG element <" + SvgName(element) + ">.");
    }

    private static string AttributeValue(XElement element, string name)
    {
        XAttribute attribute = element.Attribute(name);
        return attribute == null ? null : attribute.Value;
    }

    private static string RequiredAttribute(XElement element, string name)
    {
        string value = AttributeValue(element, name);
        if (value == null)
            throw new NotSupportedException("<" + SvgName(element) + "> requires attribute '" + name + "'.");
        return value;
    }

    private static bool IsNone(string paint)
    {
        return string.Equals(paint, "none", StringComparison.Ordinal);
    }

    private static double RequiredLength(XElement element, string name)
    {
        return ParseLength(RequiredAttribute(element, name), name);
    }

    private static double OptionalLength(XElement element, string name, double defaultValue)
    {
        string value = AttributeValue(element, name);
        return value == null ? defaultValue : ParseLength(value, name);
    }

    private static void ValidateOptionalLengths(XElement element, string[] names)
    {
        for (int index = 0; index < names.Length; index++)
        {
            string value = AttributeValue(element, names[index]);
            if (value != null)
                ParseLength(value, names[index]);
        }
    }

    private static double ParseLength(string text, string description)
    {
        string value = text.Trim();
        Match match = NumberRegex.Match(value);
        if (!match.Success || match.Index != 0 || match.Length != value.Length)
            throw new NotSupportedException("Unsupported units or invalid number '" + text + "' for " + description + ".");
        double parsed;
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed) ||
            double.IsNaN(parsed) || double.IsInfinity(parsed))
        {
            throw new NotSupportedException("Invalid finite number '" + text + "' for " + description + ".");
        }
        return parsed;
    }

    private static double[] ParseNumberList(string text, int requiredCount, string description)
    {
        double[] values = ParseFlexibleNumberList(text, description);
        if (values.Length != requiredCount)
            throw new NotSupportedException(description + " requires exactly " + requiredCount + " numbers.");
        return values;
    }

    private static double[] ParseFlexibleNumberList(string text, string description)
    {
        List<double> result = new List<double>();
        MatchCollection matches = NumberRegex.Matches(text);
        int consumed = 0;
        for (int index = 0; index < matches.Count; index++)
        {
            Match match = matches[index];
            ValidateNumberSeparator(text.Substring(consumed, match.Index - consumed), description);
            result.Add(ParseLength(match.Value, description));
            consumed = match.Index + match.Length;
        }
        ValidateNumberSeparator(text.Substring(consumed), description);
        return result.ToArray();
    }

    private static void ValidateNumberSeparator(string separator, string description)
    {
        for (int index = 0; index < separator.Length; index++)
        {
            char character = separator[index];
            if (!char.IsWhiteSpace(character) && character != ',')
                throw new NotSupportedException("Invalid numeric list for " + description + ".");
        }
    }

    private static void ValidateTransformSeparator(string separator, string original)
    {
        for (int index = 0; index < separator.Length; index++)
        {
            char character = separator[index];
            if (!char.IsWhiteSpace(character) && character != ',')
                throw new NotSupportedException("Invalid SVG transform '" + original + "'.");
        }
    }

    private static void RequireCount(double[] values, int count, string transform)
    {
        if (values.Length != count)
            throw TransformCountException(transform, count.ToString(CultureInfo.InvariantCulture));
    }

    private static NotSupportedException TransformCountException(string transform, string count)
    {
        return new NotSupportedException("Transform '" + transform + "' requires " + count + " parameters.");
    }

    private static string ParseLocalUrl(string text, string description)
    {
        Match match = Regex.Match(
            text.Trim(),
            @"\Aurl\(\s*#([^\s()]+)\s*\)\z",
            RegexOptions.CultureInvariant);
        if (!match.Success)
            throw new NotSupportedException("Unsupported " + description + " reference '" + text + "'; only local url(#id) references are supported.");
        return match.Groups[1].Value;
    }

    private static bool IsFiniteMatrix(Matrix matrix)
    {
        return IsFinite(matrix.M11) && IsFinite(matrix.M12) && IsFinite(matrix.M21) &&
            IsFinite(matrix.M22) && IsFinite(matrix.OffsetX) && IsFinite(matrix.OffsetY);
    }

    private static bool IsFinite(double value)
    {
        return !double.IsNaN(value) && !double.IsInfinity(value);
    }

    private static string ExpandHex(string value)
    {
        if (value.Length == 4)
        {
            return "#" + value[1] + value[1] + value[2] + value[2] + value[3] + value[3];
        }
        return value;
    }

    private static string EscapeAttribute(string value)
    {
        return value.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;");
    }

    private sealed class Style
    {
        public string Fill;
        public string Stroke;
        public string Color;
        public string FillRule;
        public double StrokeWidth;
        public PenLineCap LineCap;
        public PenLineJoin LineJoin;
        public double MiterLimit;
        public List<double> DashArray;
        public double DashOffset;

        public static Style CreateDefaults()
        {
            Style style = new Style();
            style.Fill = "#000000";
            style.Stroke = "none";
            style.Color = "#000000";
            style.FillRule = "nonzero";
            style.StrokeWidth = 1d;
            style.LineCap = PenLineCap.Flat;
            style.LineJoin = PenLineJoin.Miter;
            style.MiterLimit = 4d;
            style.DashArray = null;
            style.DashOffset = 0d;
            return style;
        }

        public Style Clone()
        {
            Style clone = (Style)MemberwiseClone();
            clone.DashArray = DashArray == null ? null : new List<double>(DashArray);
            return clone;
        }
    }

    private sealed class Semantic
    {
        public bool IsAccent;
        public bool IsExplicitBase;
        public string AccentRole;

        public Semantic Clone()
        {
            return (Semantic)MemberwiseClone();
        }
    }

    private sealed class PaintPath
    {
        public PaintPath(Geometry geometry, string paint, string color, Semantic semantic)
        {
            Geometry = geometry;
            Paint = paint;
            Color = color;
            Semantic = semantic;
        }

        public Geometry Geometry;
        public string Paint;
        public string Color;
        public Semantic Semantic;
    }

    private sealed class MaskOperation
    {
        public MaskOperation(Geometry geometry, bool isWhite)
        {
            Geometry = geometry;
            IsWhite = isWhite;
        }

        public Geometry Geometry;
        public bool IsWhite;
    }

    private sealed class MaskScope
    {
        public MaskScope()
        {
            Masks = new Dictionary<string, MaskDefinition>(StringComparer.Ordinal);
        }

        public readonly Dictionary<string, MaskDefinition> Masks;
    }

    private sealed class MaskDefinition
    {
        public MaskDefinition(XElement element, Style style, Semantic semantic, MaskScope scope, string id)
        {
            Element = element;
            Style = style;
            Semantic = semantic;
            Scope = scope;
            Id = id;
        }

        public readonly XElement Element;
        public readonly Style Style;
        public readonly Semantic Semantic;
        public readonly MaskScope Scope;
        public readonly string Id;
    }
}
