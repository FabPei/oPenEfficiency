using System;
using System.Collections.Generic;
using System.Linq;
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
        Name = "Map Builder",
        SearchAliases = new[] { "Map Wizard" },
        Tooltip = "Insert and color editable PowerPoint maps",
        IconData = "M15,19L9,16.89V5L15,7.11M20.5,3C20.44,3 20.39,3.03 20.34,3.05L15,5L9,3L3.36,4.9C3.15,4.97 3,5.15 3,5.38V20.5A0.5,0.5 0 0,0 3.5,21C3.55,21 3.61,20.97 3.66,20.95L9,19L15,21L20.64,19.1C20.84,19.03 21,18.85 21,18.62V3.5A0.5,0.5 0 0,0 20.5,3Z",
        Color = "#10B981", // Green mapping
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None,
        Description = "Inserts editable world, continent, country, and administrative-region maps, and colors their regions from pasted tabular data.",
        DetailedHelpText = "### Map Builder\nChoose **World**, **Continent**, or **Country or territory**, then choose the available detail level. Right-click starts at World, Continent, Country, the selected tagged map, or the most recent configuration. Administrative levels are enabled only where bundled data is available, and every inserted region is an editable native PowerPoint shape.\n\nChoose whether remote regions stay in place, move into compact insets, or are excluded. Use **Mercator** or **Natural proportions**, restore a recent selection, and color shapes by pasting region names/IDs and values from Excel.")
    ]
    public static class MapWizardFeature
    {
        public enum StartMode { Default, NewWorld, NewContinent, NewCountry, EditSelected, Recent }

        private const string TagsKey = "OPE_MAP_WIZARD_CONFIG";

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, StartMode.Default);
        }

        public static bool Execute(PowerPointManager manager, StartMode startMode)
        {
            try
            {
                var selection = manager.GetApplication().ActiveWindow.Selection;
                PowerPoint.Shape mapGroup = null;
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes &&
                    selection.ShapeRange.Count == 1)
                {
                    PowerPoint.Shape candidate = selection.ShapeRange[1];
                    if (candidate.Type == MsoShapeType.msoGroup ||
                        !string.IsNullOrWhiteSpace(ShapeMetadataService.GetTag(candidate, MapAssetService.MapIdTag)))
                    {
                        mapGroup = candidate;
                    }
                }

                if (startMode == StartMode.EditSelected && mapGroup == null)
                {
                    System.Windows.Forms.MessageBox.Show("Select one openEfficiency map to edit it.", "Map Builder");
                    return false;
                }
                if (startMode == StartMode.NewWorld || startMode == StartMode.NewContinent || startMode == StartMode.NewCountry || startMode == StartMode.Recent)
                    mapGroup = null;

                // Try to load any existing config
                MapWizardConfig config = null;
                try
                {
                    string existingJson = ShapeMetadataService.GetTag(mapGroup, TagsKey);
                    if (!string.IsNullOrEmpty(existingJson))
                    {
                        config = JsonService.Deserialize<MapWizardConfig>(existingJson);
                    }
                }
                catch { /* Unsaved presentations use the map wizard's default asset directory. */ }

                string initialScope = startMode == StartMode.NewWorld ? "World" :
                    startMode == StartMode.NewContinent ? "Continent" :
                    startMode == StartMode.NewCountry ? "Country" : null;
                var window = new MapWizardWindow(manager, mapGroup, config, initialScope, startMode == StartMode.Recent);
                if (window.ShowDialog() == true)
                {
                    mapGroup = window.MapGroup;
                    config = window.Config;
                    if (mapGroup == null)
                    {
                        return false;
                    }

                    // Save state back to shape
                    ShapeMetadataService.SetTag(mapGroup, TagsKey, JsonService.Serialize(config));

                    // Apply Colors
                    ApplyColors(mapGroup, config);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MapWizardFeature.Execute");
                throw;
            }
        }

        public static HashSet<string> GetAvailableRegionKeys(PowerPoint.Shape mapGroup)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PowerPoint.Shape regionShape in GetRegionRoots(mapGroup))
            {
                foreach (string key in GetRegionKeys(regionShape))
                {
                    keys.Add(key);
                }
            }
            return keys;
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

            foreach (PowerPoint.Shape regionShape in GetRegionRoots(mapGroup))
            {
                HashSet<string> regionKeys = GetRegionKeys(regionShape);

                // Stable ISO/admin IDs and aliases are preferred over mutable PowerPoint names.
                var row = config.Data.FirstOrDefault(d =>
                    !string.IsNullOrWhiteSpace(d.Region) && regionKeys.Contains(d.Region.Trim()));

                if (row != null)
                {
                    if (config.Mode == MapColorMode.Continuous)
                    {
                        double ratio = (row.NumericValue - minVal) / (maxVal - minVal);
                        int targetRgb = InterpolateColor(config.LowColorHex, config.HighColorHex, ratio);
                        ApplyFillRecursively(regionShape, targetRgb);
                    }
                    else // Categorical
                    {
                        if (string.IsNullOrEmpty(row.RawValue))
                        {
                            ApplyFillRecursively(regionShape, defaultColor);
                            continue;
                        }

                        if (!config.CategoryColors.ContainsKey(row.RawValue))
                        {
                            // Assign sequential color from palette
                            config.CategoryColors[row.RawValue] = paletteHex[paletteIndex % paletteHex.Length];
                            paletteIndex++;
                        }
                        
                        ApplyFillRecursively(regionShape, ColorConstants.FromHex(config.CategoryColors[row.RawValue]));
                    }
                }
                else
                {
                    // Reset to default if no data found
                    ApplyFillRecursively(regionShape, defaultColor);
                }
            }
        }

        private static IEnumerable<PowerPoint.Shape> GetRegionRoots(PowerPoint.Shape mapGroup)
        {
            if (mapGroup == null)
            {
                yield break;
            }

            if (mapGroup.Type != MsoShapeType.msoGroup)
            {
                yield return mapGroup;
                yield break;
            }

            for (int index = 1; index <= mapGroup.GroupItems.Count; index++)
            {
                PowerPoint.Shape child = mapGroup.GroupItems[index];
                // Bundled maps tag each direct child as a logical region. Untagged direct
                // children are retained for compatibility with user-supplied legacy maps.
                yield return child;
            }
        }

        private static HashSet<string> GetRegionKeys(PowerPoint.Shape regionShape)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            AddKey(keys, regionShape?.Name);
            AddKey(keys, regionShape?.AlternativeText);
            AddKey(keys, ShapeMetadataService.GetTag(regionShape, MapAssetService.RegionIdTag));

            string aliases = ShapeMetadataService.GetTag(regionShape, MapAssetService.RegionAliasesTag);
            if (!string.IsNullOrWhiteSpace(aliases))
            {
                foreach (string alias in aliases.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    AddKey(keys, alias);
                }
            }
            return keys;
        }

        private static void AddKey(HashSet<string> keys, string value)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                keys.Add(value.Trim());
            }
        }

        private static void ApplyFillRecursively(PowerPoint.Shape shape, int color)
        {
            if (shape == null)
            {
                return;
            }

            if (shape.Type == MsoShapeType.msoGroup)
            {
                for (int index = 1; index <= shape.GroupItems.Count; index++)
                {
                    ApplyFillRecursively(shape.GroupItems[index], color);
                }
                return;
            }

            shape.Fill.Visible = MsoTriState.msoTrue;
            shape.Fill.Solid();
            shape.Fill.ForeColor.RGB = color;
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
