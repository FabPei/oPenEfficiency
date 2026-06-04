using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnGlassHide",
        Name = "Glass Hide",
        Tooltip = "Overlay selection with glass shapes",
        IconData = "M2,16V4C2,2.89 2.89,2 4,2H20A2,2 0 0,1 22,4V16A2,2 0 0,1 20,18H4A2,2 0 0,1 2,16M4,4V16H20V4H4M18,12V14H16V12H18M18,8V10H16V8H18M14,12V14H6V12H14M14,8V10H6V8H14M2,20H22V22H2V20Z",
        Color = "#6366F1",
        Description = "Inserts semi-transparent white overlays on top of the current selection.",
        DetailedHelpText = "### Glass Hide\nCreates semi-transparent 'glass' shapes that overlay your selection. Useful for temporarily obscuring content or creating a frosted glass effect.\n\n**Right-Click Options:**\n* **Single Bounding Box:** Creates one large shape covering the entire selection area.\n* **Individual Shapes:** Creates one overlay for each individual shape in the selection.",
        Keywords = "glass, overlay, hide, conceal, frosted, transparent",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class GlassHideFeature
    {
        public static float DefaultTransparency { get; set; } = 0.5f;
        public static int DefaultColorRgb { get; set; } = 0xFFFFFF;

        public enum GlassMode
        {
            Single,
            Individual
        }

        /// <summary>
        /// Default execution - Single Bounding Box.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, GlassMode.Single, DefaultColorRgb, DefaultTransparency);
        }

        /// <summary>
        /// Executes Glass Hide with specified mode.
        /// </summary>
        public static bool Execute(PowerPointManager manager, GlassMode mode, int colorRgb = -1, float transparency = -1f)
        {
            if (colorRgb == -1) colorRgb = DefaultColorRgb;
            if (transparency < 0f) transparency = DefaultTransparency;

            if (manager == null) return false;
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count == 0) return false;

            try
            {
                // ShapeRange.Parent is the Slide object
                var slide = shapeRange.Parent as Slide;
                if (slide == null) return false;

                if (mode == GlassMode.Single)
                {
                    CreateGlassShape(slide, shapeRange.Left, shapeRange.Top, shapeRange.Width, shapeRange.Height, colorRgb, transparency);
                }
                else
                {
                    // Create individual overlays
                    for (int i = 1; i <= shapeRange.Count; i++)
                    {
                        var s = shapeRange[i];
                        CreateGlassShape(slide, s.Left, s.Top, s.Width, s.Height, colorRgb, transparency);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GlassHideFeature.Execute");
                return false;
            }
        }

        private static void CreateGlassShape(Slide slide, float left, float top, float width, float height, int colorRgb = 0xFFFFFF, float transparency = 0.5f)
        {
            var glass = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, width, height);

            // Configure Glass Appearance
            glass.Fill.Visible = Office.MsoTriState.msoTrue;
            glass.Fill.ForeColor.RGB = colorRgb; 
            glass.Fill.Transparency = transparency; 

            glass.Line.Visible = Office.MsoTriState.msoFalse;

            glass.Shadow.Visible = Office.MsoTriState.msoFalse;

            // Tag it for identification
            glass.Tags.Add("oPE_Type", "GlassHide");
            glass.Name = "GlassHide_" + Guid.NewGuid().ToString().Substring(0, 8);
        }
    }
}
