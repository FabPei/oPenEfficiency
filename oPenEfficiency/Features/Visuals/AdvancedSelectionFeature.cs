using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Selection engine used by SelectSameTypeFeature. Not registered as a sidebar button directly.
    /// </summary>
    public static class AdvancedSelectionFeature
    {
        public enum SelectSameCriteria
        {
            Type,
            FillColor,
            LineColor,
            FillAndLineColor,
            Size,
            Width,
            Height,
            TypeAndFormatting,
        }

        private static bool AreGradientsEqual(Shape shape1, Shape shape2)
        {
            try
            {
                // Check gradient style and direction
                if (shape1.Fill.GradientStyle != shape2.Fill.GradientStyle ||
                    Math.Abs(shape1.Fill.GradientDegree - shape2.Fill.GradientDegree) > 0.1f)
                {
                    return false;
                }

                // Check gradient stops count
                var stops1 = shape1.Fill.GradientStops;
                var stops2 = shape2.Fill.GradientStops;
                if (stops1.Count != stops2.Count)
                {
                    return false;
                }

                // Check each gradient stop
                for (int i = 1; i <= stops1.Count; i++)
                {
                    if (stops1[i].Color.RGB != stops2[i].Color.RGB ||
                        Math.Abs(stops1[i].Position - stops2[i].Position) > 0.01f)
                    {
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error comparing gradients: {ex.Message}");
                return false;
            }
        }

        private static bool ArePatternsEqual(Shape shape1, Shape shape2)
        {
            try
            {
                // Check pattern type
                if (shape1.Fill.Pattern != shape2.Fill.Pattern)
                {
                    return false;
                }

                // Check foreground and background colors
                if (shape1.Fill.ForeColor.RGB != shape2.Fill.ForeColor.RGB ||
                    shape1.Fill.BackColor.RGB != shape2.Fill.BackColor.RGB)
                {
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error comparing patterns: {ex.Message}");
                return false;
            }
        }

        private static bool AreTexturesEqual(Shape shape1, Shape shape2)
        {
            try
            {
                // Check texture type
                if (shape1.Fill.TextureType != shape2.Fill.TextureType)
                {
                    return false;
                }

                // Check texture name (if available)
                try
                {
                    var textureName1 = shape1.Fill.TextureName;
                    var textureName2 = shape2.Fill.TextureName;
                    if (!string.IsNullOrEmpty(textureName1) && !string.IsNullOrEmpty(textureName2))
                    {
                        return textureName1 == textureName2;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error getting texture name: {ex.Message}");
                    // If texture names aren't available, fall back to other properties
                }

                // Check if both have pictures
                if (shape1.Fill.Type == Office.MsoFillType.msoFillPicture && shape2.Fill.Type == Office.MsoFillType.msoFillPicture)
                {
                    try
                    {
                        // For picture fills, we can't easily compare the actual images,
                        // but we can check if they have the same dimensions and other properties
                        return Math.Abs(shape1.Fill.PictureEffects.Count - shape2.Fill.PictureEffects.Count) < 1;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error comparing picture fills: {ex.Message}");
                        // If picture comparison fails, consider them different
                        return false;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error comparing textures: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Wrapper for auto-discovery - selects shapes of the same type as the first selected shape.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, SelectSameCriteria.Type);
        }

        public static bool Execute(PowerPointManager manager, SelectSameCriteria criteria)
        {
            if (manager == null) return false;

            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count == 0) return false;

            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;
                var slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                var refShape = manager.GetReferenceShape();
                
                // Reference values
                var refType = refShape.Type;
                Office.MsoAutoShapeType? refAutoShapeType = null;
                if (refType == Office.MsoShapeType.msoAutoShape)
                {
                    try { refAutoShapeType = refShape.AutoShapeType; }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error getting auto shape type: {ex.Message}");
                    }
                }

                int? refFillColor = null;
                Office.MsoFillType? refFillType = null;
                Office.MsoGradientStyle? refGradientStyle = null;
                float? refGradientDegree = null;
                List<(int Color, float Position)> refGradientStops = new List<(int, float)>();
                Office.MsoPatternType? refPatternType = null;
                int? refPatternForeColor = null;
                int? refPatternBackColor = null;
                Office.MsoTextureType? refTextureType = null;
                string refTextureName = null;
                
                try 
                { 
                    if (refShape.Fill.Visible == Office.MsoTriState.msoTrue) 
                    {
                        refFillType = refShape.Fill.Type;
                        
                        // For solid fills, get the RGB value
                        if (refShape.Fill.Type == Office.MsoFillType.msoFillSolid)
                        {
                            refFillColor = refShape.Fill.ForeColor.RGB;
                        }
                        // For gradient fills, capture full gradient pattern
                        else if (refShape.Fill.Type == Office.MsoFillType.msoFillGradient)
                        {
                            try
                            {
                                refGradientStyle = refShape.Fill.GradientStyle;
                                refGradientDegree = refShape.Fill.GradientDegree;
                                
                                var gradientStops = refShape.Fill.GradientStops;
                                for (int i = 1; i <= gradientStops.Count; i++)
                                {
                                    refGradientStops.Add((gradientStops[i].Color.RGB, gradientStops[i].Position));
                                }
                                
                                // Use first gradient stop as primary color for fallback
                                if (refGradientStops.Count > 0)
                                {
                                    refFillColor = refGradientStops[0].Color;
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error extracting gradient pattern: {ex.Message}");
                                // If gradient pattern extraction fails, fall back to ForeColor
                                refFillColor = refShape.Fill.ForeColor.RGB;
                            }
                        }
                        // For pattern fills, capture pattern information
                        else if (refShape.Fill.Type == Office.MsoFillType.msoFillPatterned)
                        {
                            try
                            {
                                refPatternType = refShape.Fill.Pattern;
                                refPatternForeColor = refShape.Fill.ForeColor.RGB;
                                refPatternBackColor = refShape.Fill.BackColor.RGB;
                                refFillColor = refPatternForeColor; // Use foreground color as primary
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error extracting pattern: {ex.Message}");
                                refFillColor = refShape.Fill.ForeColor.RGB;
                            }
                        }
                        // For texture/picture fills, capture texture information
                        else if (refShape.Fill.Type == Office.MsoFillType.msoFillTextured || refShape.Fill.Type == Office.MsoFillType.msoFillPicture)
                        {
                            try
                            {
                                refTextureType = refShape.Fill.TextureType;
                                refTextureName = refShape.Fill.TextureName;
                                refFillColor = refShape.Fill.ForeColor.RGB; // Use ForeColor as primary
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error extracting texture: {ex.Message}");
                                refFillColor = refShape.Fill.ForeColor.RGB;
                            }
                        }
                        else
                        {
                            // For other fill types, use ForeColor as fallback
                            refFillColor = refShape.Fill.ForeColor.RGB;
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error getting fill properties: {ex.Message}");
                }

                int? refLineColor = null;
                try
                {
                    if (refShape.Line.Visible == Office.MsoTriState.msoTrue)
                        refLineColor = refShape.Line.ForeColor.RGB;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error getting line color: {ex.Message}");
                }

                float refWidth = refShape.Width;
                float refHeight = refShape.Height;

                List<string> matchingShapeNames = new List<string>();

                foreach (Shape s in slide.Shapes)
                {
                    bool match = false;
                    try
                    {
                        switch (criteria)
                        {
                            case SelectSameCriteria.Type:
                                if (s.Type == refType)
                                {
                                    if (refType == Office.MsoShapeType.msoAutoShape)
                                    {
                                        if (refAutoShapeType.HasValue && s.AutoShapeType == refAutoShapeType.Value) match = true;
                                    }
                                    else match = true;
                                }
                                break;

                            case SelectSameCriteria.FillColor:
                                try
                                {
                                    if (refFillColor.HasValue && s.Fill.Visible == Office.MsoTriState.msoTrue) 
                                    {
                                        // For solid fills, compare directly
                                        if (s.Fill.Type == Office.MsoFillType.msoFillSolid && refFillType == Office.MsoFillType.msoFillSolid)
                                        {
                                            if (s.Fill.ForeColor.RGB == refFillColor.Value) match = true;
                                        }
                                        // For gradient fills, compare full gradient pattern
                                        else if (s.Fill.Type == Office.MsoFillType.msoFillGradient && refFillType == Office.MsoFillType.msoFillGradient)
                                        {
                                            if (AreGradientsEqual(refShape, s))
                                            {
                                                match = true;
                                            }
                                            else
                                            {
                                                // Fallback to first gradient stop comparison
                                                try
                                                {
                                                    var gradientStops = s.Fill.GradientStops;
                                                    if (gradientStops.Count > 0 && gradientStops[1].Color.RGB == refFillColor.Value)
                                                    {
                                                        match = true;
                                                    }
                                                }
                                                catch
                                                {
                                                    // Final fallback to ForeColor comparison
                                                    if (s.Fill.ForeColor.RGB == refFillColor.Value) match = true;
                                                }
                                            }
                                        }
                                        // For pattern fills, compare full pattern
                                        else if (s.Fill.Type == Office.MsoFillType.msoFillPatterned && refFillType == Office.MsoFillType.msoFillPatterned)
                                        {
                                            if (ArePatternsEqual(refShape, s))
                                            {
                                                match = true;
                                            }
                                            else
                                            {
                                                // Fallback to pattern foreground color comparison
                                                if (s.Fill.ForeColor.RGB == refPatternForeColor.Value) match = true;
                                            }
                                        }
                                        // For texture/picture fills, compare texture
                                        else if ((s.Fill.Type == Office.MsoFillType.msoFillTextured || s.Fill.Type == Office.MsoFillType.msoFillPicture) && 
                                                (refFillType == Office.MsoFillType.msoFillTextured || refFillType == Office.MsoFillType.msoFillPicture))
                                        {
                                            if (AreTexturesEqual(refShape, s))
                                            {
                                                match = true;
                                            }
                                            else
                                            {
                                                // Fallback to ForeColor comparison
                                                if (s.Fill.ForeColor.RGB == refFillColor.Value) match = true;
                                            }
                                        }
                                        // For different fill types, compare ForeColor as fallback
                                        else if (s.Fill.ForeColor.RGB == refFillColor.Value)
                                        {
                                            match = true;
                                        }
                                    }
                                    else if (!refFillColor.HasValue && s.Fill.Visible == Office.MsoTriState.msoFalse) 
                                    {
                                        match = true;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error comparing fill colors: {ex.Message}");
                                }
                                break;

                            case SelectSameCriteria.LineColor:
                                try
                                {
                                    if (refLineColor.HasValue && s.Line.Visible == Office.MsoTriState.msoTrue && s.Line.ForeColor.RGB == refLineColor.Value) match = true;
                                    else if (!refLineColor.HasValue && s.Line.Visible == Office.MsoTriState.msoFalse) match = true;
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error comparing line colors: {ex.Message}");
                                }
                                break;

                            case SelectSameCriteria.FillAndLineColor:
                                try
                                {
                                    bool fillMatch = false;
                                    if (refFillColor.HasValue && s.Fill.Visible == Office.MsoTriState.msoTrue) 
                                    {
                                        // For solid fills, compare directly
                                        if (s.Fill.Type == Office.MsoFillType.msoFillSolid && refFillType == Office.MsoFillType.msoFillSolid)
                                        {
                                            if (s.Fill.ForeColor.RGB == refFillColor.Value) fillMatch = true;
                                        }
                                        // For gradient fills, compare full gradient pattern
                                        else if (s.Fill.Type == Office.MsoFillType.msoFillGradient && refFillType == Office.MsoFillType.msoFillGradient)
                                        {
                                            if (AreGradientsEqual(refShape, s))
                                            {
                                                fillMatch = true;
                                            }
                                            else
                                            {
                                                // Fallback to first gradient stop comparison
                                                try
                                                {
                                                    var gradientStops = s.Fill.GradientStops;
                                                    if (gradientStops.Count > 0 && gradientStops[1].Color.RGB == refFillColor.Value)
                                                    {
                                                        fillMatch = true;
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error comparing gradient stops: {ex.Message}");
                                                    // Final fallback to ForeColor comparison
                                                    if (s.Fill.ForeColor.RGB == refFillColor.Value) fillMatch = true;
                                                }
                                            }
                                        }
                                        // For pattern fills, compare full pattern
                                        else if (s.Fill.Type == Office.MsoFillType.msoFillPatterned && refFillType == Office.MsoFillType.msoFillPatterned)
                                        {
                                            if (ArePatternsEqual(refShape, s))
                                            {
                                                fillMatch = true;
                                            }
                                            else
                                            {
                                                // Fallback to pattern foreground color comparison
                                                if (s.Fill.ForeColor.RGB == refPatternForeColor.Value) fillMatch = true;
                                            }
                                        }
                                        // For texture/picture fills, compare texture
                                        else if ((s.Fill.Type == Office.MsoFillType.msoFillTextured || s.Fill.Type == Office.MsoFillType.msoFillPicture) && 
                                                (refFillType == Office.MsoFillType.msoFillTextured || refFillType == Office.MsoFillType.msoFillPicture))
                                        {
                                            if (AreTexturesEqual(refShape, s))
                                            {
                                                fillMatch = true;
                                            }
                                            else
                                            {
                                                // Fallback to ForeColor comparison
                                                if (s.Fill.ForeColor.RGB == refFillColor.Value) fillMatch = true;
                                            }
                                        }
                                        // For different fill types, compare ForeColor as fallback
                                        else if (s.Fill.ForeColor.RGB == refFillColor.Value)
                                        {
                                            fillMatch = true;
                                        }
                                    }
                                    else if (!refFillColor.HasValue && s.Fill.Visible == Office.MsoTriState.msoFalse) 
                                    {
                                        fillMatch = true;
                                    }

                                    bool lineMatch = false;
                                    if (refLineColor.HasValue && s.Line.Visible == Office.MsoTriState.msoTrue && s.Line.ForeColor.RGB == refLineColor.Value) 
                                        lineMatch = true;
                                    else if (!refLineColor.HasValue && s.Line.Visible == Office.MsoTriState.msoFalse) 
                                        lineMatch = true;

                                    if (fillMatch && lineMatch) match = true;
                                } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"AdvancedSelectionFeature.FillLineMatch error: {ex.Message}"); }
                                break;

                            case SelectSameCriteria.Size:
                                if (Math.Abs(s.Width - refWidth) < 0.1f && Math.Abs(s.Height - refHeight) < 0.1f) match = true;
                                break;

                            case SelectSameCriteria.Width:
                                if (Math.Abs(s.Width - refWidth) < 0.1f) match = true;
                                break;

                            case SelectSameCriteria.Height:
                                if (Math.Abs(s.Height - refHeight) < 0.1f) match = true;
                                break;

                            case SelectSameCriteria.TypeAndFormatting:
                                if (s.Type == refType)
                                {
                                    bool typeMatch = true;
                                    if (refType == Office.MsoShapeType.msoAutoShape)
                                    {
                                        if (refAutoShapeType.HasValue && s.AutoShapeType != refAutoShapeType.Value) typeMatch = false;
                                    }
                                    
                                    if (typeMatch)
                                    {
                                        bool formattingMatch = true;
                                        try
                                        {
                                            if (s.Fill.Visible != refShape.Fill.Visible) formattingMatch = false;
                                            else if (s.Fill.Visible == Office.MsoTriState.msoTrue)
                                            {
                                                if (s.Fill.ForeColor.RGB != refShape.Fill.ForeColor.RGB) formattingMatch = false;
                                                else if (Math.Abs(s.Fill.Transparency - refShape.Fill.Transparency) > 0.01f) formattingMatch = false;
                                            }

                                            if (formattingMatch)
                                            {
                                                if (s.Line.Visible != refShape.Line.Visible) formattingMatch = false;
                                                else if (s.Line.Visible == Office.MsoTriState.msoTrue)
                                                {
                                                    if (s.Line.ForeColor.RGB != refShape.Line.ForeColor.RGB) formattingMatch = false;
                                                    else if (Math.Abs(s.Line.Weight - refShape.Line.Weight) > 0.01f) formattingMatch = false;
                                                    else if (s.Line.DashStyle != refShape.Line.DashStyle) formattingMatch = false;
                                                    else if (Math.Abs(s.Line.Transparency - refShape.Line.Transparency) > 0.01f) formattingMatch = false;
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error comparing formatting: {ex.Message}");
                                            formattingMatch = false;
                                        }

                                        if (formattingMatch) match = true;
                                    }
                                }
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"AdvancedSelection: Error in shape matching: {ex.Message}");
                    }

                    if (match) matchingShapeNames.Add(s.Name);
                }

                if (matchingShapeNames.Count > 0)
                {
                    slide.Shapes.Range(matchingShapeNames.ToArray()).Select();
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AdvancedSelection Error: {ex}");
                return false;
            }
        }
    }
}
