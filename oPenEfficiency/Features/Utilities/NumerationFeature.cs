using System;
using System.Drawing;
using System.Globalization;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert numbered shapes (numeration markers).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnNumeration",
        Name = "Numbered Marker",
        Tooltip = "Insert a numbered marker",
        SearchAliases = new[] { "Numeration" },
        IconData = "M5,3H19V5H5V3M5,7H19V9H5V7M5,11H19V13H5V11M5,15H19V17H5V15M3,21H17L20,21V19H3V21M7,19H9V21H7V19M11,19H13V21H11V19M15,19H17V21H15V19Z",
        Color = "#F43F5E",
        Description = "Inserts one editable numbered marker on the current slide.",
        DetailedHelpText = "### Numbered Marker\n\nInsert one circular numbered marker on the current slide. Right-click for direct number 1 through 5 presets. Select the marker to edit its number, size, color, opacity, or shape from the contextual toolbar. The marker stores its settings in openEfficiency metadata and can be changed without overwriting text in other shapes.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class NumerationFeature
    {
        public const string NumerationTag = oPenEfficiency.Core.Services.NumerationStatePolicy.NumerationTag;

        /// <summary>
        /// Wrapper for auto-discovery - creates numeration marker with default settings.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Create(manager);
        }

        public struct NumerationSettings
        {
            public int Number;
            public string ShapeType; // "Circle", "Square", "Hexagon", etc.
            public float Size; // In points (0.8cm approx 22.7pt)
            public string ColorHex;
            public int Alpha; // 0-255

            public static NumerationSettings Default()
            {
                oPenEfficiency.Core.Services.NumerationStatePolicy.Defaults(
                    out int number, out string shapeType, out float size, out string colorHex, out int alpha);
                return new NumerationSettings
                {
                    Number = number,
                    ShapeType = shapeType,
                    Size = size,
                    ColorHex = colorHex,
                    Alpha = alpha
                };
            }

            public string ToTag() =>
                oPenEfficiency.Core.Services.NumerationStatePolicy.FormatTag(Number, ShapeType, Size, ColorHex, Alpha);

            public static NumerationSettings FromTag(string tag)
            {
                oPenEfficiency.Core.Services.NumerationStatePolicy.ParseState(
                    tag, out int number, out string shapeType, out float size, out string colorHex, out int alpha);
                return new NumerationSettings
                {
                    Number = number,
                    ShapeType = shapeType,
                    Size = size,
                    ColorHex = colorHex,
                    Alpha = alpha
                };
            }
        }

        public static bool Create(PowerPointManager manager)
        {
            var settings = NumerationSettings.Default();
            return Execute(manager, null, settings);
        }

        public static bool Execute(PowerPointManager manager, PowerPoint.Shape existingShape, NumerationSettings settings)
        {
            if (manager == null) return false;
            PowerPoint.Shape shape = null;

            try
            {
                var slide = manager.GetCurrentSlide();
                if (slide == null) return false;
                if (!oPenEfficiency.Core.Services.NumerationStatePolicy.IsValidSize(settings.Size))
                    throw new ArgumentOutOfRangeException(nameof(settings.Size), "Marker size must be between 5 and 200 points.");

                settings.Alpha = oPenEfficiency.Core.Services.NumerationStatePolicy.ClampAlpha(settings.Alpha);
                Color markerColor = ColorTranslator.FromHtml(settings.ColorHex ?? "#005C94");

                float left = 100, top = 100;
                if (existingShape != null)
                {
                    oPenEfficiency.Core.Services.NumerationStatePolicy.CenterOnShape(
                        existingShape.Left, existingShape.Top, existingShape.Width, existingShape.Height, settings.Size,
                        out left, out top);
                }

                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "NumerationFeature.StartNewUndoEntry"); }

                var type = Office.MsoAutoShapeType.msoShapeOval;
                if (settings.ShapeType == "Square") type = Office.MsoAutoShapeType.msoShapeRectangle;
                else if (settings.ShapeType == "Diamond") type = Office.MsoAutoShapeType.msoShapeDiamond;
                else if (settings.ShapeType == "Hexagon") type = Office.MsoAutoShapeType.msoShapeHexagon;

                shape = slide.Shapes.AddShape(type, left, top, settings.Size, settings.Size);
                shape.Fill.Visible = Office.MsoTriState.msoTrue;
                shape.Fill.ForeColor.RGB = ColorTranslator.ToWin32(markerColor);
                shape.Fill.Transparency = 1.0f - (settings.Alpha / 255.0f);
                shape.Line.Visible = Office.MsoTriState.msoFalse;

                shape.TextFrame.TextRange.Text = settings.Number.ToString(CultureInfo.InvariantCulture);
                shape.TextFrame.TextRange.Font.Size = 10;
                shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF;
                shape.TextFrame.MarginLeft = 0;
                shape.TextFrame.MarginRight = 0;
                shape.TextFrame.MarginTop = 0;
                shape.TextFrame.MarginBottom = 0;
                shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                shape.AlternativeText = settings.ToTag();
                shape.Name = "oPE_Bullet_" + Guid.NewGuid().ToString("N").Substring(0, 8);

                // Delete the predecessor only after its fully formatted replacement exists.
                if (existingShape != null) existingShape.Delete();
                shape.Select();
                return true;
            }
            catch (Exception ex)
            {
                if (shape != null)
                {
                    try { shape.Delete(); }
                    catch (Exception cleanupEx) { ExceptionLogger.Log(cleanupEx, "NumerationFeature.Cleanup"); }
                }
                ExceptionLogger.Log(ex, "NumerationFeature.Execute");
                return false;
            }
        }
    }
}
