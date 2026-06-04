using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Services;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert or update a Thermometer infographic.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnThermometer",
        Name = "Thermometer",
        Tooltip = "Thermometer",
        IconData = "M15,13V5A3,3 0 0,0 9,5V13A5,5 0 1,0 12,22A5,5 0 0,0 15,13M12,20A3,3 0 0,1 9,17A3,3 0 0,1 12,14V5A1,1 0 1,1 14,5V14A3,3 0 0,1 12,20Z",
        Color = "#F43F5E",
        Description = "Inserts a thermometer visualization for project progress. Default is 50% fill.",
        DetailedHelpText = "### Thermometer\nInserts or updates a vertical thermometer-style progress bar.\n**Usage:**\n* Left-click to cycle the fill level up by 25%.\n**Right-Click Options:**\n* Set the fill level exactly (0%, 25%, 50%, 75%, 100%).",
        Keywords = "thermometer chart, vertical progress, status indicator, level fill",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class ThermometerFeature
    {
        public const string ThermometerTag = "oPen_Thermometer";

        /// <summary>
        /// Wrapper for auto-discovery - inserts thermometer with default 50% fill.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, 0.5f, colorRgb: 0x0000FF);
        }

        public static bool Execute(PowerPointManager manager, float percentage = 0.5f, int colorRgb = 0x0000FF)
        {
            if (manager == null) return false;

            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;

                // Check if a Thermometer is currently selected to update it
                var selection = manager.GetSelectedShapes();
                if (selection != null && selection.Count == 1)
                {
                    var selected = selection[1];
                    if (IsThermometer(selected))
                    {
                        UpdateThermometer(selected, percentage, colorRgb);
                        return true;
                    }
                }

                // Otherwise insert new one at center of slide
                var slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                float thermometerWidth = 18f;
                float thermometerHeight = 80f;
                float bulbSize = 30f;

                float totalHeight = thermometerHeight + (bulbSize / 2);
                float left = (app.ActivePresentation.PageSetup.SlideWidth / 2) - (bulbSize / 2);
                float top = (app.ActivePresentation.PageSetup.SlideHeight / 2) - (totalHeight / 2);

                // 1. Stem (Rounded Rectangle)
                var stem = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                    left + (bulbSize - thermometerWidth) / 2, top, thermometerWidth, thermometerHeight + 5);
                stem.Name = "Stem";
                stem.Fill.ForeColor.RGB = 0xEEEEEE;
                stem.Line.ForeColor.RGB = 0x808080;
                stem.Line.Weight = 1.0f;
                try { stem.Adjustments[1] = 0.5f; } catch (Exception ex) { ExceptionLogger.Log(ex, "ThermometerFeature.Execute.Stem"); }

                // 2. Mercury Fill (Rounded Rectangle)
                float fillHeight = percentage * thermometerHeight;
                var mercury = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                    left + (bulbSize - thermometerWidth) / 2 + 3,
                    top + thermometerHeight - fillHeight,
                    thermometerWidth - 6, fillHeight + 10);
                mercury.Name = "Mercury";
                mercury.Fill.ForeColor.RGB = colorRgb;
                mercury.Line.Visible = Office.MsoTriState.msoFalse;
                try { mercury.Adjustments[1] = 0.5f; } catch (Exception ex) { ExceptionLogger.Log(ex, "ThermometerFeature.Execute.Mercury"); }

                // 3. Bulb (Circle)
                var bulb = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval,
                    left, top + thermometerHeight - 10, bulbSize, bulbSize);
                bulb.Name = "Bulb";
                bulb.Fill.ForeColor.RGB = colorRgb;
                bulb.Line.ForeColor.RGB = 0x808080;
                bulb.Line.Weight = 1.0f;

                // 4. Group them
                var items = new string[] { stem.Name, mercury.Name, bulb.Name };
                var group = slide.Shapes.Range(items).Group();
                group.AlternativeText = $"{ThermometerTag}|{percentage}|{colorRgb}";
                group.Name = "Thermometer_" + Guid.NewGuid().ToString().Substring(0, 8);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThermometerFeature.Execute");
                return false;
            }
        }

        public static bool IsThermometer(Shape shape)
        {
            if (shape == null) return false;
            return shape.AlternativeText.StartsWith(ThermometerTag);
        }

        public static void UpdateThermometer(Shape group, float percentage, int? colorRgb)
        {
            if (group == null || (group.Type != Office.MsoShapeType.msoGroup && !IsThermometer(group))) return;

            // If it's a grouped shape, find sub-components
            Shape stem = null, bulb = null, mercury = null;
            if (group.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (Shape s in group.GroupItems)
                {
                    if (s.Name == "Stem") stem = s;
                    else if (s.Name == "Bulb") bulb = s;
                    else if (s.Name == "Mercury") mercury = s;
                }

                // Fallback by type/order if names lost
                if (stem == null) stem = group.GroupItems[1];
                if (mercury == null) mercury = group.GroupItems[2];
                if (bulb == null) bulb = group.GroupItems[3];
            }
            else
            {
                return;
            }

            // Clamp percentage
            percentage = Math.Max(0, Math.Min(1, percentage));

            // Update Color
            if (colorRgb.HasValue)
            {
                bulb.Fill.ForeColor.RGB = colorRgb.Value;
                mercury.Fill.ForeColor.RGB = colorRgb.Value;
            }

            // Update Mercury Fill
            float stemHeight = stem.Height - 5;
            float fillHeight = percentage * stemHeight;

            float mercuryBottom = stem.Top + stemHeight;
            mercury.Top = mercuryBottom - fillHeight;
            mercury.Height = fillHeight + 10;

            // Update Tag
            int finalColor = colorRgb ?? (bulb.Fill.ForeColor.RGB);
            group.AlternativeText = $"{ThermometerTag}|{percentage}|{finalColor}";
        }
    }
}
