using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;

namespace oPenEfficiency.UI
{
    /// <summary>
    /// Element-aware icon geometry produced by the SVG role-aware parser.
    /// Geometry and authored paint semantics stay together so mixed fill/stroke icons can render
    /// without being flattened into one incompatible path style.
    /// </summary>
    public sealed class ParsedIconGeometry
    {
        public static readonly Rect DefaultViewBox = new Rect(0d, 0d, 24d, 24d);

        public IList<ParsedIconElement> Elements { get; } = new List<ParsedIconElement>();

        /// <summary>
        /// Source SVG viewport. Rendering normalizes this rectangle into the shared 24×24 canvas.
        /// </summary>
        public Rect ViewBox { get; set; } = DefaultViewBox;

        /// <summary>Main icon geometry; always present for a valid icon.</summary>
        public GeometryGroup BaseGeometry { get; set; }

        /// <summary>
        /// Optional accent sub-element identified by <c>data-oe-role="accent"</c> in the SVG.
        /// Null for neutral-only icons.
        /// </summary>
        public Geometry AccentGeometry { get; set; }

        /// <summary>True when the SVG contains a <c>data-oe-role="accent"</c> element.</summary>
        public bool HasAccent => AccentGeometry != null;
    }

    /// <summary>
    /// One drawable SVG element with its geometry, layer role, and authored paint treatment.
    /// Keeping these values together prevents mixed fill/stroke icons from being flattened into
    /// one incompatible WPF Path style.
    /// </summary>
    public sealed class ParsedIconElement
    {
        public Geometry Geometry { get; set; }
        public bool IsAccent { get; set; }
        public FeatureIconAccentRole AccentRole { get; set; }
        public bool HasExplicitAccentRole { get; set; }
        public bool HasFill { get; set; }
        public bool HasStroke { get; set; }
        /// <summary>Authored SVG fill paint, retained for source-faithful previews.</summary>
        public string FillPaint { get; set; }
        /// <summary>Authored SVG stroke paint, retained for source-faithful previews.</summary>
        public string StrokePaint { get; set; }
        public double StrokeThickness { get; set; }
        public PenLineCap StrokeLineCap { get; set; }
        public PenLineJoin StrokeLineJoin { get; set; }
    }
}
