using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Visuals
{
    [FeatureMetadata(
        Id = "BtnTransparency", 
        Name = "Transparency", 
        Tooltip = "Adjust Transparency", 
        IconData = "M12 2A10 10 0 0 0 2 12A10 10 0 0 0 12 22A10 10 0 0 0 22 12A10 10 0 0 0 12 2M12 4A8 8 0 0 1 20 12A8 8 0 0 1 12 20V4Z", 
        Color = "#06B6D4", 
        MinSelection = 1, 
        RequiredType = PpSelectionType.ppSelectionShapes,
        IsToggle = true)]
    public static class TransparencyFeature
    {
        private static oPenEfficiency.UI.Toolbars.TransparencyToolbar _currentWindow;

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, null);
        }

        public static bool Execute(PowerPointManager manager, System.Windows.Controls.Primitives.ToggleButton toggleBtn)
        {
            if (_currentWindow != null && _currentWindow.IsLoaded)
            {
                _currentWindow.Close();
                // Close triggers the event to uncheck the button
                return true;
            }

            _currentWindow = new oPenEfficiency.UI.Toolbars.TransparencyToolbar(manager);
            
            if (toggleBtn != null)
            {
                toggleBtn.IsChecked = true;
                _currentWindow.Closed += (s, e) => {
                    toggleBtn.IsChecked = false;
                };
            }

            _currentWindow.Show();
            return true;
        }

        public static bool SetTransparency(PowerPointManager manager, float transparencyPercentage)
        {
            if (manager == null) return false;
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count == 0) return false;

            var list = new System.Collections.Generic.List<Shape>();
            for (int i = 1; i <= shapeRange.Count; i++) list.Add(shapeRange[i]);

            return SetTransparency(manager, list, transparencyPercentage);
        }

        public static bool SetTransparency(PowerPointManager manager, System.Collections.Generic.List<Shape> shapes, float transparencyPercentage)
        {
            if (manager == null || shapes == null || shapes.Count == 0) return false;

            // Ensure transparency is between 0.0 and 1.0
            float t = Math.Max(0f, Math.Min(1f, transparencyPercentage));

            try
            {
                foreach (var shape in shapes)
                {

                    // Determine if it's a line/connector
                    if (shape.Type == Office.MsoShapeType.msoLine || shape.Connector == Office.MsoTriState.msoTrue)
                    {
                        if (shape.Line != null)
                        {
                            shape.Line.Transparency = t;
                        }
                    }
                    // Handle pictures
                    else if (shape.Type == Office.MsoShapeType.msoPicture || shape.Type == Office.MsoShapeType.msoLinkedPicture)
                    {
                        try 
                        {
                            if (shape.PictureFormat != null)
                            {
                                ((dynamic)shape.PictureFormat).Transparency = t;
                            }
                        }
                        catch 
                        {
                            // Fallback if PictureFormat.Transparency is not supported (older PPT versions)
                            if (shape.Fill != null)
                            {
                                shape.Fill.Transparency = t;
                            }
                        }
                    }
                    // Standard shapes (rectangles, circles, etc.)
                    else
                    {
                        if (shape.Fill != null)
                        {
                            shape.Fill.Transparency = t;
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TransparencyFeature.SetTransparency");
                return false;
            }
        }
    }
}
