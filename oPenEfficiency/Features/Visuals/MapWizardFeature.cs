using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.UI.Dialogs;

namespace oPenEfficiency.Features.Visuals
{
    [FeatureMetadata(
        Id = "BtnMapWizard",
        Name = "Map Wizard",
        Tooltip = "Map Wizard: Geographic Colorizer",
        IconData = "M15,19L9,16.89V5L15,7.11M20.5,3C20.44,3 20.39,3.03 20.34,3.05L15,5L9,3L3.36,4.9C3.15,4.97 3,5.15 3,5.38V20.5A0.5,0.5 0 0,0 3.5,21C3.55,21 3.61,20.97 3.66,20.95L9,19L15,21L20.64,19.1C20.84,19.03 21,18.85 21,18.62V3.5A0.5,0.5 0 0,0 20.5,3Z",
        Color = "#10B981", // Green mapping
        MinSelection = 1,
        RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes,
        Description = "Links geographic vector maps to data to achieve conditional color rendering of regions natively.",
        Keywords = "map, wizard, geography, colorizer, regions",
        DetailedHelpText = "### Map Wizard\nManages geographic vector map assets (world map, regions, countries). Allows binding Excel conditional formatting data to fill individual country shapes automatically based on data values.")
    ]
    public static class MapWizardFeature
    {
        private const string TagsKey = "oPE_MapWizardConfig";

        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var selection = manager.GetApplication().ActiveWindow.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count != 1)
                {
                    System.Windows.MessageBox.Show("Please select exactly one grouped Map shape.", "Map Wizard", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return false;
                }

                PowerPoint.Shape mapGroup = selection.ShapeRange[1];

                if (mapGroup.Type != MsoShapeType.msoGroup)
                {
                    System.Windows.MessageBox.Show("The map must be a Grouped shape where each child shape represents a region.", "Map Wizard", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return false;
                }

                // Try to load any existing config
                MapWizardConfig config = null;
                try
                {
                    string existingJson = mapGroup.Tags[TagsKey];
                    if (!string.IsNullOrEmpty(existingJson))
                    {
                        config = JsonService.Deserialize<MapWizardConfig>(existingJson);
                    }
                }
                catch { }

                var window = new MapWizardWindow(mapGroup, config);
                if (window.ShowDialog() == true)
                {
                    config = window.Config;
                    
                    // Save state back to shape
                    try {
                        mapGroup.Tags.Add(TagsKey, JsonService.Serialize(config));
                    } catch { }

                    // Apply Colors
                    ApplyColors(mapGroup, config);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex.Message, "MapWizardFeature");
                return false;
            }
        }

        private static void ApplyColors(PowerPoint.Shape mapGroup, MapWizardConfig config)
        {
            if (config.Data == null || config.Data.Count == 0) return;

            // Setup bounds for continuous mode
            double minVal = 0, maxVal = 0;
            if (config.Mode == MapColorMode.Continuous)
            {
                minVal = config.Data.Min(x => x.NumericValue);
                maxVal = config.Data.Max(x => x.NumericValue);
                if (maxVal == minVal) maxVal = minVal + 1; // Prevent div zero
            }

            // Assign standard fallback color
            int defaultColor = ColorConstants.FromHex(config.DefaultColorHex);

            // Palette generator for Categorical Mode
            string[] paletteHex = { "#4285F4", "#EA4335", "#FBBC05", "#34A853", "#8E24AA", "#FF6D00" };
            int paletteIndex = 0;

            foreach (PowerPoint.Shape regionShape in mapGroup.GroupItems)
            {
                string shapeName = regionShape.Name;
                string altText = regionShape.AlternativeText;

                // Find data row by Name or AltText
                var row = config.Data.FirstOrDefault(d => 
                    string.Equals(d.Region, shapeName, StringComparison.OrdinalIgnoreCase) || 
                    string.Equals(d.Region, altText, StringComparison.OrdinalIgnoreCase));

                if (row != null)
                {
                    if (config.Mode == MapColorMode.Continuous)
                    {
                        double ratio = (row.NumericValue - minVal) / (maxVal - minVal);
                        int targetRgb = InterpolateColor(config.LowColorHex, config.HighColorHex, ratio);
                        regionShape.Fill.ForeColor.RGB = targetRgb;
                    }
                    else // Categorical
                    {
                        if (string.IsNullOrEmpty(row.RawValue))
                        {
                            regionShape.Fill.ForeColor.RGB = defaultColor;
                            continue;
                        }

                        if (!config.CategoryColors.ContainsKey(row.RawValue))
                        {
                            // Assign sequential color from palette
                            config.CategoryColors[row.RawValue] = paletteHex[paletteIndex % paletteHex.Length];
                            paletteIndex++;
                        }
                        
                        regionShape.Fill.ForeColor.RGB = ColorConstants.FromHex(config.CategoryColors[row.RawValue]);
                    }
                }
                else
                {
                    // Reset to default if no data found
                    regionShape.Fill.ForeColor.RGB = defaultColor;
                }
            }
        }

        private static int InterpolateColor(string hex1, string hex2, double ratio)
        {
            if (ratio < 0) ratio = 0;
            if (ratio > 1) ratio = 1;

            hex1 = hex1.Replace("#", "").Trim();
            hex2 = hex2.Replace("#", "").Trim();

            if (hex1.Length != 6 || hex2.Length != 6) return oPenEfficiency.Utils.ColorConstants.FromHex(hex1);

            int r1 = Convert.ToInt32(hex1.Substring(0, 2), 16);
            int g1 = Convert.ToInt32(hex1.Substring(2, 2), 16);
            int b1 = Convert.ToInt32(hex1.Substring(4, 2), 16);

            int r2 = Convert.ToInt32(hex2.Substring(0, 2), 16);
            int g2 = Convert.ToInt32(hex2.Substring(2, 2), 16);
            int b2 = Convert.ToInt32(hex2.Substring(4, 2), 16);

            int r3 = (int)(r1 + (r2 - r1) * ratio);
            int g3 = (int)(g1 + (g2 - g1) * ratio);
            int b3 = (int)(b1 + (b2 - b1) * ratio);

            return oPenEfficiency.Utils.ColorConstants.FromRgb(r3, g3, b3);
        }
    }
}
