using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Services;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert or update a Star Rating infographic.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStarRating",
        Name = "Star Rating",
        Tooltip = "Insert a 1-5 star rating",
        IconData = "M12,17.27L18.18,21L16.54,13.97L22,9.24L14.81,8.63L12,2L9.19,8.63L2,9.24L7.46,13.97L5.82,21L12,17.27Z",
        Color = "#F43F5E",
        Description = "Inserts a sequence of stars representing a rating from 1 to 5. Default is 5 stars, 3 rating.",
        DetailedHelpText = "### Star Rating\nInserts a customizable star rating indicator (1 to 5 stars, with half-star support filled from left to right) as a native PowerPoint vector group. Right-click for direct 1-of-5 through 5-of-5 presets; a selected rating is updated in place.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class StarRatingFeature
    {
        public const string StarRatingTag = oPenEfficiency.Core.Services.StarRatingStatePolicy.StarRatingTag;
        private static readonly int SR_Dim = 0xD0D0D0; // Light grey for empty stars

        /// <summary>
        /// Wrapper for auto-discovery - inserts 5-star rating with default 3.0 rating.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, starCount: 5, rating: 3.0, iconMode: 0, hasFrame: false, fillColorBgr: 0x0000FF);
        }

        public static bool IsStarRating(Shape shape)
        {
            if (shape == null) return false;
            if (shape.AlternativeText != null && shape.AlternativeText.StartsWith(StarRatingTag)) return true;
            
            string ee4pTag = ShapeMetadataService.GetTag(shape, "EE4P_SMART_ELEMENT");
            if (ee4pTag == "Indicator") return true;
            
            if (!string.IsNullOrEmpty(ShapeMetadataService.GetTag(shape, "THINKCELLSHAPEDONOTDELETE")))
            {
                // Thinkcell usually uses groups for everything, check if group items look like stars
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    foreach (Shape item in shape.GroupItems)
                    {
                        if (item.AutoShapeType == Office.MsoAutoShapeType.msoShape5pointStar) return true;
                    }
                }
            }

            return false;
        }

        public static bool Execute(PowerPointManager manager, int starCount, double rating, int iconMode, bool hasFrame, int fillColorBgr, Shape targetShape = null)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;

                Shape existingGroup = targetShape;

                if (existingGroup == null)
                {
                    // Check if a StarRating is selected to update
                    var selection = manager.GetSelectedShapes();
                    if (selection != null && selection.Count == 1)
                    {
                        var selected = selection[1];
                        if (IsStarRating(selected))
                        {
                            existingGroup = selected;
                        }
                        else if (selected.Child == Office.MsoTriState.msoTrue)
                        {
                            try
                            {
                                var parent = selected.ParentGroup;
                                if (IsStarRating(parent))
                                {
                                    existingGroup = parent;
                                }
                            }
                            catch { /* Optional parent-group detection falls back to the selected star. */ }
                        }
                    }
                }

                if (existingGroup != null)
                {
                    UpdateStarRating(existingGroup, starCount, rating, iconMode, hasFrame, fillColorBgr);
                    return true;
                }

                // Insert new
                var slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                float shapeSize = 24f;
                float gap = 4f;
                float totalWidth = starCount * shapeSize + (starCount - 1) * gap;
                
                // Calculate position: center of selection if any, else center of slide
                float startLeft, top;
                var selShapes = manager.GetSelectedShapes();
                if (selShapes != null && selShapes.Count > 0)
                {
                    float selCenterW = selShapes.Left + (selShapes.Width / 2);
                    float selCenterH = selShapes.Top + (selShapes.Height / 2);
                    startLeft = selCenterW - (totalWidth / 2);
                    top = selCenterH - (shapeSize / 2);
                }
                else
                {
                    startLeft = (app.ActivePresentation.PageSetup.SlideWidth / 2) - (totalWidth / 2);
                    top = (app.ActivePresentation.PageSetup.SlideHeight / 2) - (shapeSize / 2);
                }

                var shapeNames = new string[starCount];

                for (int i = 0; i < starCount; i++)
                {
                    float left = startLeft + i * (shapeSize + gap);
                    Shape star = CreateRatingShape(slide, iconMode, left, top, shapeSize);
                    shapeNames[i] = star.Name;

                    ApplyStarFill(star, i, rating, fillColorBgr);

                    if (hasFrame)
                    {
                        star.Line.Visible = Office.MsoTriState.msoTrue;
                        star.Line.ForeColor.RGB = 0x808080;
                        star.Line.Weight = 0.75f;
                    }
                    else
                    {
                        star.Line.Visible = Office.MsoTriState.msoFalse;
                    }
                }

                var group = slide.Shapes.Range(shapeNames).Group();
                string tag = FormatStarRatingTag(starCount, rating, iconMode, hasFrame, fillColorBgr);
                ShapeMetadataService.SetTag(group, "oPE_State", tag);
                group.AlternativeText = tag;
                group.Name = "StarRating_" + Guid.NewGuid().ToString().Substring(0, 8);
                group.Select();

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StarRatingFeature.Execute");
                return false;
            }
        }

        private static void UpdateStarRating(Shape group, int starCount, double rating, int iconMode, bool hasFrame, int fillColorBgr)
        {
            try
            {
                if (group.Type != Office.MsoShapeType.msoGroup) return;
                var items = group.GroupItems;
                int existingCount = items.Count;

                string oldTag = group.AlternativeText ?? "";
                int oldIconMode = 0;
                bool oldHasFrame = false;
                if (oPenEfficiency.Core.Services.StarRatingStatePolicy.TryParseTag(oldTag, out _, out _, out oldIconMode, out oldHasFrame, out _))
                {
                    // parsed through the production codec
                }

                if (existingCount != starCount || oldIconMode != iconMode || oldHasFrame != hasFrame)
                {
                    float groupLeft = group.Left;
                    float groupTop = group.Top;
                    var slide = group.Parent as Slide;
                    if (slide == null) return;

                    group.Delete();

                    float shapeSize = 24f;
                    float gap = 4f;
                    float top = groupTop;

                    var shapeNames = new string[starCount];
                    for (int i = 0; i < starCount; i++)
                    {
                        float left = groupLeft + i * (shapeSize + gap);
                        Shape star = CreateRatingShape(slide, iconMode, left, top, shapeSize);
                        shapeNames[i] = star.Name;
                        ApplyStarFill(star, i, rating, fillColorBgr);
                        star.Line.Visible = hasFrame ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                        if (hasFrame) { star.Line.ForeColor.RGB = 0x808080; star.Line.Weight = 0.75f; }
                    }

                    var newGroup = slide.Shapes.Range(shapeNames).Group();
                    string newTag = FormatStarRatingTag(starCount, rating, iconMode, hasFrame, fillColorBgr);
                    ShapeMetadataService.SetTag(newGroup, "oPE_State", newTag);
                    newGroup.AlternativeText = newTag;
                    newGroup.Name = "StarRating_" + Guid.NewGuid().ToString().Substring(0, 8);
                    newGroup.Select();
                    return;
                }

                for (int i = 0; i < existingCount; i++)
                {
                    var star = items[i + 1]; 
                    ApplyStarFill(star, i, rating, fillColorBgr);
                    star.Line.Visible = hasFrame ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                    if (hasFrame) { star.Line.ForeColor.RGB = 0x808080; star.Line.Weight = 0.75f; }
                }

                group.AlternativeText = FormatStarRatingTag(starCount, rating, iconMode, hasFrame, fillColorBgr);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StarRatingFeature.UpdateStarRating");
            }
        }

        private static Shape CreateRatingShape(Slide slide, int iconMode, float left, float top, float size)
        {
            Office.MsoAutoShapeType shapeType;
            switch (iconMode)
            {
                case 1: shapeType = Office.MsoAutoShapeType.msoShapeOval; break;
                case 2: shapeType = Office.MsoAutoShapeType.msoShapeDiamond; break;
                default: shapeType = Office.MsoAutoShapeType.msoShape5pointStar; break;
            }
            return slide.Shapes.AddShape(shapeType, left, top, size, size);
        }

        private static void ApplyStarFill(Shape star, int index, double rating, int fillColorBgr)
        {
            oPenEfficiency.Core.Services.StarRatingStatePolicy.ResolveFill(index, rating, out var fillKind, out double fraction);

            if (fillKind == oPenEfficiency.Core.Services.StarRatingStatePolicy.FillKind.Full)
            {
                star.Fill.Solid();
                star.Fill.ForeColor.RGB = fillColorBgr;
            }
            else if (fillKind == oPenEfficiency.Core.Services.StarRatingStatePolicy.FillKind.Dim)
            {
                star.Fill.Solid();
                star.Fill.ForeColor.RGB = SR_Dim;
            }
            else
            {
                float stopPercent = (float)(fraction * 100);

                try
                {
                    // PowerPoint's "vertical" gradient style moves colors horizontally.
                    star.Fill.TwoColorGradient(Office.MsoGradientStyle.msoGradientVertical, 1);
                    star.Fill.GradientStops[1].Color.RGB = fillColorBgr;
                    star.Fill.GradientStops[1].Position = 0f;
                    star.Fill.GradientStops.Insert(fillColorBgr, stopPercent / 100f, 1f, 2);
                    star.Fill.GradientStops.Insert(SR_Dim, stopPercent / 100f, 1f, 3);
                    star.Fill.GradientStops[star.Fill.GradientStops.Count].Color.RGB = SR_Dim;
                    star.Fill.GradientStops[star.Fill.GradientStops.Count].Position = 1f;
                }
                catch
                {
                    star.Fill.Solid();
                    star.Fill.ForeColor.RGB = oPenEfficiency.Core.Services.StarRatingStatePolicy.UseSolidFallback(fraction) ? fillColorBgr : SR_Dim;
                }
            }
        }

        private static string FormatStarRatingTag(int count, double rating, int iconMode, bool hasFrame, int colorBgr)
        {
            return oPenEfficiency.Core.Services.StarRatingStatePolicy.FormatTag(count, rating, iconMode, hasFrame, colorBgr);
        }
    }
}
