using System;
using Microsoft.Office.Interop.PowerPoint;
using System.Globalization;
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
        SearchAliases = new[] { "Thermometer Chart" },
        Tooltip = "Insert a thermometer or increase the selected thermometer by 25%",
        IconData = "M15,13V5A3,3 0 0,0 9,5V13A5,5 0 1,0 12,22A5,5 0 0,0 15,13M12,20A3,3 0 0,1 9,17A3,3 0 0,1 12,14V5A1,1 0 1,1 14,5V14A3,3 0 0,1 12,20Z",
        Color = "#F43F5E",
        Description = "Inserts a thermometer visualization for project progress. The default fill level is 50%.",
        DetailedHelpText = "### Thermometer\nInserts or updates a vertical thermometer-style progress bar.\n\n**Usage:**\n* With no thermometer selected, left-click to insert one at 50%.\n* Select an existing thermometer and left-click to increase its fill by 25%; after 100% it wraps to 0%.\n\n**Right-click options:**\n* Set the fill level exactly (0%, 25%, 50%, 75%, or 100%). Existing thermometers keep their current color.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class ThermometerFeature
    {
        public const string ThermometerTag = oPenEfficiency.Core.Services.ThermometerStatePolicy.ThermometerTag;

        /// <summary>
        /// Wrapper for auto-discovery - inserts thermometer with default 50% fill.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;

            try
            {
                var selection = manager.GetSelectedShapes();
                if (selection != null && selection.Count == 1 &&
                    TryReadState(selection[1], out float currentPercentage, out int currentColor))
                {
                    float nextPercentage = oPenEfficiency.Core.Services.ThermometerStatePolicy.NextPercentage(currentPercentage);
                    return Execute(manager, nextPercentage, currentColor);
                }

                return Execute(manager, 0.5f, colorRgb: 0x0000FF);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThermometerFeature.Execute.Default");
                return false;
            }
        }

        public static bool Execute(PowerPointManager manager, float percentage = 0.5f, int colorRgb = -1)
        {
            if (manager == null) return false;

            Shape stem = null;
            Shape mercury = null;
            Shape bulb = null;
            Shape group = null;
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;

                percentage = oPenEfficiency.Core.Services.ThermometerStatePolicy.ClampPercentage(percentage);

                // Check if a Thermometer is currently selected to update it
                var selection = manager.GetSelectedShapes();
                if (selection != null && selection.Count == 1)
                {
                    var selected = selection[1];
                    if (IsThermometer(selected))
                    {
                        if (colorRgb < 0 && TryReadState(selected, out _, out int currentColor))
                            colorRgb = currentColor;
                        if (colorRgb < 0) colorRgb = 0x0000FF;
                        return UpdateThermometer(selected, percentage, colorRgb);
                    }
                }

                if (colorRgb < 0) colorRgb = 0x0000FF;

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
                string componentId = Guid.NewGuid().ToString("N").Substring(0, 8);
                stem = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                    left + (bulbSize - thermometerWidth) / 2, top, thermometerWidth, thermometerHeight + 5);
                stem.Name = $"oPE_Thermometer_Stem_{componentId}";
                stem.Fill.ForeColor.RGB = 0xEEEEEE;
                stem.Line.ForeColor.RGB = 0x808080;
                stem.Line.Weight = 1.0f;
                try { stem.Adjustments[1] = 0.5f; } catch (Exception ex) { ExceptionLogger.Log(ex, "ThermometerFeature.Execute.Stem"); }

                // 2. Mercury Fill (Rounded Rectangle)
                float fillHeight = percentage * thermometerHeight;
                mercury = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                    left + (bulbSize - thermometerWidth) / 2 + 3,
                    top + thermometerHeight - fillHeight,
                    thermometerWidth - 6, fillHeight + 10);
                mercury.Name = $"oPE_Thermometer_Mercury_{componentId}";
                mercury.Fill.ForeColor.RGB = colorRgb;
                mercury.Line.Visible = Office.MsoTriState.msoFalse;
                try { mercury.Adjustments[1] = 0.5f; } catch (Exception ex) { ExceptionLogger.Log(ex, "ThermometerFeature.Execute.Mercury"); }

                // 3. Bulb (Circle)
                bulb = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval,
                    left, top + thermometerHeight - 10, bulbSize, bulbSize);
                bulb.Name = $"oPE_Thermometer_Bulb_{componentId}";
                bulb.Fill.ForeColor.RGB = colorRgb;
                bulb.Line.ForeColor.RGB = 0x808080;
                bulb.Line.Weight = 1.0f;

                // 4. Group them
                var items = new string[] { stem.Name, mercury.Name, bulb.Name };
                group = slide.Shapes.Range(items).Group();
                group.AlternativeText = oPenEfficiency.Core.Services.ThermometerStatePolicy.FormatTag(percentage, colorRgb);
                group.Name = "Thermometer_" + Guid.NewGuid().ToString().Substring(0, 8);

                return true;
            }
            catch (Exception ex)
            {
                if (group != null) TryDelete(group);
                else
                {
                    TryDelete(bulb);
                    TryDelete(mercury);
                    TryDelete(stem);
                }
                ExceptionLogger.Log(ex, "ThermometerFeature.Execute");
                return false;
            }
        }

        public static bool IsThermometer(Shape shape)
        {
            if (shape == null) return false;
            try
            {
                string alternativeText = shape.AlternativeText;
                return string.Equals(alternativeText, ThermometerTag, StringComparison.Ordinal) ||
                    (!string.IsNullOrEmpty(alternativeText) &&
                     alternativeText.StartsWith(ThermometerTag + "|", StringComparison.Ordinal));
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThermometerFeature.IsThermometer");
                return false;
            }
        }

        public static bool UpdateThermometer(Shape group, float percentage, int? colorRgb)
        {
            if (group == null || group.Type != Office.MsoShapeType.msoGroup || !IsThermometer(group)) return false;

            try
            {
                // Find components by their stable names; retain the positional fallback for
                // thermometers created by earlier versions of the add-in.
                Shape stem = null, bulb = null, mercury = null;
                foreach (Shape s in group.GroupItems)
                {
                    if (s.Name == "Stem" || s.Name.StartsWith("oPE_Thermometer_Stem_", StringComparison.Ordinal)) stem = s;
                    else if (s.Name == "Bulb" || s.Name.StartsWith("oPE_Thermometer_Bulb_", StringComparison.Ordinal)) bulb = s;
                    else if (s.Name == "Mercury" || s.Name.StartsWith("oPE_Thermometer_Mercury_", StringComparison.Ordinal)) mercury = s;
                }

                // Fallback by type/order if names lost
                if ((stem == null || mercury == null || bulb == null) && group.GroupItems.Count < 3) return false;
                if (stem == null) stem = group.GroupItems[1];
                if (mercury == null) mercury = group.GroupItems[2];
                if (bulb == null) bulb = group.GroupItems[3];

                // Clamp percentage
                percentage = Math.Max(0, Math.Min(1, percentage));

                // Update Color
                if (colorRgb.HasValue)
                {
                    bulb.Fill.ForeColor.RGB = colorRgb.Value;
                    mercury.Fill.ForeColor.RGB = colorRgb.Value;
                }

                // Update Mercury Fill
                float stemHeight = Math.Max(0, stem.Height - 5);
                float fillHeight = percentage * stemHeight;

                float mercuryBottom = stem.Top + stemHeight;
                mercury.Top = mercuryBottom - fillHeight;
                mercury.Height = fillHeight + 10;

                // Update Tag
                int finalColor = colorRgb ?? bulb.Fill.ForeColor.RGB;
                group.AlternativeText = $"{ThermometerTag}|{percentage.ToString(CultureInfo.InvariantCulture)}|{finalColor}";
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThermometerFeature.UpdateThermometer");
                return false;
            }
        }

        private static bool TryReadState(Shape shape, out float percentage, out int colorRgb)
        {
            percentage = 0.5f;
            colorRgb = 0x0000FF;
            if (!IsThermometer(shape)) return false;

            try
            {
                oPenEfficiency.Core.Services.ThermometerStatePolicy.ParseState(shape.AlternativeText, out percentage, out colorRgb);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThermometerFeature.TryReadState");
                return false;
            }
        }

        private static void TryDelete(Shape shape)
        {
            if (shape == null) return;
            try { shape.Delete(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ThermometerFeature.TryDelete"); }
        }
    }
}
