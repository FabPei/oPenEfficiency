using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Services;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert or update a Harvey Ball.
    /// Values usually toggle: 0 -> 0.25 -> 0.5 -> 0.75 -> 1.0
    /// </summary>
    /// <remarks>
    /// <para>State Persistence: Harvey Ball state (percentage) is stored in the shape's
    /// <c>AlternativeText</c> property using the format "oPE_HarveyBall|{percentage}".</para>
    /// <para>Why AlternativeText? PowerPoint doesn't persist custom properties reliably through
    /// copy/paste operations. AlternativeText is the only shape property that survives
    /// copy/paste, making it ideal for persisting feature state with the shape.</para>
    /// </remarks>
    [FeatureMetadata(
        Id = "BtnHarveyBall",
        Name = "Harvey Ball",
        Tooltip = "Harvey ball",
        IconData = "M12,2A10,10 0 0,0 2,12A10,10 0 0,0 12,22A10,10 0 0,0 22,12A10,10 0 0,0 12,2M12,4A8,8 0 0,1 20,12A8,8 0 0,1 12,20A8,8 0 0,1 4,12A8,8 0 0,1 12,4Z",
        Color = "#F43F5E",
        Description = "Cycles through Harvey Ball states (0%, 25%, 50%, 75%, 100%) to represent progress or qualitative data.",
        DetailedHelpText = "### Harvey Ball\nInserts or updates a circular progress indicator (Harvey Ball).\n**Usage:**\n* Left-click to cycle through 0 -> 25 -> 50 -> 75 -> 100% fill.\n**Right-Click Options:**\n* Jump directly to a specific fill level (0%, 25%, 50%, 75%, or 100%).",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class HarveyBallFeature
    {
        private const string HarveyBallTag = "oPE_HarveyBall";

        /// <summary>
        /// Cycles through 0 → 25 → 50 → 75 → 100 → 0% when a Harvey Ball is selected;
        /// inserts a new 25% ball otherwise.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            var selection = manager?.GetSelectedShapes();
            float nextPercentage = 0.25f;
            if (selection != null && selection.Count == 1)
            {
                var selected = selection[1];
                if (IsHarveyBall(selected))
                {
                    float current = GetCurrentPercentage(selected);
                    nextPercentage = current + 0.25f;
                    if (nextPercentage > 1.01f) nextPercentage = 0f;
                }
            }
            return Execute(manager, nextPercentage);
        }

        public static bool Execute(PowerPointManager manager, float percentage = 0.25f)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;

                var selection = manager.GetSelectedShapes();
                if (selection != null && selection.Count == 1)
                {
                    var selected = selection[1];
                    if (IsHarveyBall(selected))
                    {
                        UpdateHarveyBall(selected, percentage);
                        return true;
                    }
                }

                var slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                float size = 30f;
                float left = (app.ActivePresentation.PageSetup.SlideWidth / 2) - (size / 2);
                float top = (app.ActivePresentation.PageSetup.SlideHeight / 2) - (size / 2);

                var bg = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left, top, size, size);
                bg.Fill.Visible = Office.MsoTriState.msoFalse;
                bg.Line.ForeColor.RGB = 0x000000;
                bg.Line.Weight = 1f;

                var pie = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapePie, left, top, size, size);
                pie.Line.Visible = Office.MsoTriState.msoFalse;
                pie.Fill.ForeColor.RGB = 0x000000;
                
                pie.Adjustments[1] = -90f; 
                pie.Adjustments[2] = -90f + (percentage * 360f);

                if (percentage <= 0) pie.Visible = Office.MsoTriState.msoFalse;

                var items = new string[] { bg.Name, pie.Name };
                var group = slide.Shapes.Range(items).Group();
                group.AlternativeText = $"{HarveyBallTag}|{percentage}";
                group.Name = "HarveyBall_" + Guid.NewGuid().ToString().Substring(0, 8);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "HarveyBallFeature.Execute");
                return false;
            }
        }

        public static bool IsHarveyBall(Shape shape)
        {
            if (shape == null) return false;

            if (shape.AlternativeText != null && shape.AlternativeText.StartsWith(HarveyBallTag)) return true;

            bool isEE4P = ShapeMetadataService.GetTag(shape, "EE4P_SMART_ELEMENT") == "HarveyBall";
            bool isThinkcell = !string.IsNullOrEmpty(ShapeMetadataService.GetTag(shape, "THINKCELLSHAPEDONOTDELETE"));
            if (isEE4P) return true;

            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                if (isThinkcell)
                {
                    foreach (Shape item in shape.GroupItems)
                        if (item.AutoShapeType == Office.MsoAutoShapeType.msoShapePie || item.Name.ToLower().Contains("arc") || item.Name.ToLower().Contains("pie")) return true;
                }

                if (shape.GroupItems.Count >= 2 && shape.GroupItems.Count <= 4)
                {
                    bool hasArc = false;
                    bool hasBg = false;

                    foreach (Shape item in shape.GroupItems)
                    {
                        string name = item.Name.ToLower();
                        if (item.AutoShapeType == Office.MsoAutoShapeType.msoShapePie || name == "arc" || name.Contains("pie")) hasArc = true;
                        if (item.AutoShapeType == Office.MsoAutoShapeType.msoShapeOval || name == "circle" || name == "background" || name.Contains("oval")) hasBg = true;
                    }

                    return hasArc && hasBg;
                }
            }

            return false;
        }

        public static float GetCurrentPercentage(Shape shape)
        {
            if (shape == null) return 0f;
            if (shape.AlternativeText != null && shape.AlternativeText.StartsWith(HarveyBallTag))
            {
                var parts = shape.AlternativeText.Split('|');
                if (parts.Length == 2 && float.TryParse(parts[1], out float current)) return current;
            }

            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (Shape s in shape.GroupItems)
                {
                    if (s.AutoShapeType == Office.MsoAutoShapeType.msoShapePie || s.Name.ToLower().Contains("arc") || s.Name.ToLower().Contains("pie"))
                    {
                        if (s.Visible == Office.MsoTriState.msoFalse) return 0f;
                        try
                        {
                            float pct = (s.Adjustments[2] + 90f) / 360f;
                            if (pct < 0) pct += 1.0f;
                            return (float)Math.Round(pct * 4f) / 4f;
                        }
                        catch { }
                    }
                }
            }
            return 0f;
        }

        private static void UpdateHarveyBall(Shape group, float percentage)
        {
            if (group.Type != Office.MsoShapeType.msoGroup) return;

            foreach (Shape s in group.GroupItems)
            {
                if (s.AutoShapeType == Office.MsoAutoShapeType.msoShapePie || s.Name.ToLower().Contains("arc") || s.Name.ToLower().Contains("pie"))
                {
                    if (percentage <= 0)
                    {
                        s.Visible = Office.MsoTriState.msoFalse;
                    }
                    else
                    {
                        s.Visible = Office.MsoTriState.msoTrue;
                        try
                        {
                            s.Adjustments[1] = -90f;
                            s.Adjustments[2] = -90f + (percentage * 360f);
                        }
                        catch (Exception ex)
                        {
                            ExceptionLogger.Log(ex, "HarveyBallFeature.UpdateHarveyBall.Adjustments");
                        }
                    }
                }
            }
            ShapeMetadataService.SetTag(group, "oPE_State", $"{HarveyBallTag}|{percentage}");
        }
    }
}
