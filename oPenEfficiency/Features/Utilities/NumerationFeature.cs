using System;
using System.Drawing;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert numbered shapes (numeration markers).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnNumeration",
        Name = "Numeration",
        Tooltip = "Numeration",
        IconData = "M5,3H19V5H5V3M5,7H19V9H5V7M5,11H19V13H5V11M5,15H19V17H5V15M3,21H17L20,21V19H3V21M7,19H9V21H7V19M11,19H13V21H11V19M15,19H17V21H15V19Z",
        Color = "#F43F5E",
        Description = "Inserts numbered shapes for marking and organizing content.",
        DetailedHelpText = "### Numeration\nAutomatically numbers all selected shapes sequentially by injecting or overwriting their text content with the configured prefix and number format.",
        Keywords = "numbers, list, bullet points, numbering, sequence, step",
        MinSelection = 0)]
    public static class NumerationFeature
    {
        public const string NumerationTag = "oPE_Numeration";

        /// <summary>
        /// Wrapper for auto-discovery - creates numeration marker with default settings.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            Create(manager);
            return true;
        }

        public struct NumerationSettings
        {
            public int Number;
            public string ShapeType; // "Circle", "Square", "Hexagon", etc.
            public float Size; // In points (0.8cm approx 22.7pt)
            public string ColorHex;
            public int Alpha; // 0-255

            public static NumerationSettings Default() => new NumerationSettings
            {
                Number = 1,
                ShapeType = "Circle",
                Size = 22.68f, // 0.8cm
                ColorHex = "#005C94",
                Alpha = 255
            };

            public string ToTag() => $"{NumerationTag}|{Number}|{ShapeType}|{Size}|{ColorHex}|{Alpha}";

            public static NumerationSettings FromTag(string tag)
            {
                var settings = Default();
                if (string.IsNullOrEmpty(tag) || !tag.StartsWith(NumerationTag)) return settings;

                var parts = tag.Split('|');
                if (parts.Length >= 2) int.TryParse(parts[1], out settings.Number);
                if (parts.Length >= 3) settings.ShapeType = parts[2];
                if (parts.Length >= 4) float.TryParse(parts[3], out settings.Size);
                if (parts.Length >= 5) settings.ColorHex = parts[4];
                if (parts.Length >= 6) int.TryParse(parts[5], out settings.Alpha);

                return settings;
            }
        }

        public static void Create(PowerPointManager manager)
        {
            var settings = NumerationSettings.Default();
            Execute(manager, null, settings);
        }

        public static void Execute(PowerPointManager manager, PowerPoint.Shape existingShape, NumerationSettings settings)
        {
            var slide = manager.GetCurrentSlide();
            if (slide == null) return;

            float left = 100, top = 100;
            if (existingShape != null)
            {
                left = existingShape.Left + (existingShape.Width / 2) - (settings.Size / 2);
                top = existingShape.Top + (existingShape.Height / 2) - (settings.Size / 2);
            }

            var type = Office.MsoAutoShapeType.msoShapeOval;
            if (settings.ShapeType == "Square") type = Office.MsoAutoShapeType.msoShapeRectangle;
            else if (settings.ShapeType == "Hexagon") type = Office.MsoAutoShapeType.msoShapeHexagon;

            var shape = slide.Shapes.AddShape(type, left, top, settings.Size, settings.Size);
            
            // Formatting
            shape.Fill.Visible = Office.MsoTriState.msoTrue;
            shape.Fill.ForeColor.RGB = ColorTranslator.ToWin32(ColorTranslator.FromHtml(settings.ColorHex));
            shape.Fill.Transparency = 1.0f - (settings.Alpha / 255.0f);
            
            shape.Line.Visible = Office.MsoTriState.msoFalse;

            // Text
            shape.TextFrame.TextRange.Text = settings.Number.ToString();
            shape.TextFrame.TextRange.Font.Size = 10;
            shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF; // White
            shape.TextFrame.MarginLeft = 0;
            shape.TextFrame.MarginRight = 0;
            shape.TextFrame.MarginTop = 0;
            shape.TextFrame.MarginBottom = 0;
            shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
            shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;

            shape.AlternativeText = settings.ToTag();
            shape.Name = "oPE_Bullet_" + Guid.NewGuid().ToString("N").Substring(0, 8);

            if (existingShape != null)
            {
                existingShape.Delete();
            }

            shape.Select();
        }
    }
}
