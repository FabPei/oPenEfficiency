using System;
using System.Collections.Generic;
using System.Drawing;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Utilities
{
    /// <summary>
    /// Feature: Flight Mode - Temporarily obscure sensitive content (logos, branding)
    /// when presenting in public spaces.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnFlightMode",
        Name = "Flight Mode",
        Tooltip = "Flight Mode (Privacy)",
        IconData = "M21,16V4H13L11,2H3V16H11L13,18H21M19,14H13.8L12.8,12H5V4H9.2L10.2,6H19V14M17,18V20H7V18H17Z",
        Color = "#F43F5E",
        Description = "Toggles flight mode to obscure sensitive content when presenting in public.",
        DetailedHelpText = "### Flight Mode\nToggles privacy overlays to hide sensitive company branding or logos when presenting in public spaces.\n\n**Right-Click Options:**\n* **Cover Logos Only:** Identifies picture shapes in the slide master and covers them with white rectangles.\n* **Light Overlay:** Adds a 5% semi-transparent white overlay over the entire slide.\n* **Heavy Overlay:** Adds a 15% semi-transparent white overlay over the entire slide.\n* **Disable:** Removes all privacy overlays.",
        MinSelection = 0,
        IsToggle = true,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class FlightModeFeature
    {
        private const string FlightModeTag = "oPE_FlightMode";
        private const string ImageCoverTag = "ImageCover";
        private const string OverlayTag = "Overlay";

        public enum FlightModeType
        {
            None,
            CoverLogos,
            LightOverlay,
            HeavyOverlay,
            HideBackground
        }

        /// <summary>
        /// Toggles flight mode (defaults to cover logos).
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Toggle(manager);
        }

        /// <summary>
        /// Gets the current Flight Mode state of the presentation.
        /// </summary>
        public static FlightModeType GetMode(PowerPointManager manager)
        {
            try
            {
                var pres = manager.GetApplication().ActivePresentation;
                var master = pres.SlideMaster;

                // Check for overlay shapes in the master
                for (int i = 1; i <= master.Shapes.Count; i++)
                {
                    Shape shape = master.Shapes[i];
                    string tagValue = GetTagValue(shape, FlightModeTag);
                    if (tagValue == ImageCoverTag)
                        return FlightModeType.CoverLogos;
                    if (tagValue == OverlayTag)
                    {
                        // Check transparency to determine light vs heavy
                        if (shape.Fill.Transparency < 0.1f)
                            return FlightModeType.HeavyOverlay;
                        return FlightModeType.LightOverlay;
                    }
                }

                return FlightModeType.None;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FlightModeFeature.GetMode");
                return FlightModeType.None;
            }
        }

        /// <summary>
        /// Toggles Flight Mode (defaults to Cover Logos Only).
        /// </summary>
        public static bool Toggle(PowerPointManager manager)
        {
            try
            {
                var currentMode = GetMode(manager);

                if (currentMode != FlightModeType.None)
                {
                    return Disable(manager);
                }
                else
                {
                    return EnableCoverLogos(manager);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FlightModeFeature.Toggle");
                return false;
            }
        }

        /// <summary>
        /// Enables Flight Mode - Cover Logos Only method.
        /// Covers all image shapes in master slides with white rectangles.
        /// </summary>
        public static bool EnableCoverLogos(PowerPointManager manager)
        {
            try
            {
                var pres = manager.GetApplication().ActivePresentation;
                int coversAdded = 0;

                // First, disable any other flight modes
                DisableOverlays(pres);

                var master = pres.SlideMaster;
                // Find all picture/image shapes in the master
                var imagesToCover = new List<Shape>();
                FindImageShapes(master.Shapes, imagesToCover);

                foreach (Shape image in imagesToCover)
                {
                    try
                    {
                        // Create a white cover rectangle over the image
                        var cover = master.Shapes.AddShape(
                            Office.MsoAutoShapeType.msoShapeRectangle,
                            image.Left,
                            image.Top,
                            image.Width,
                            image.Height);

                        // Style the cover
                        cover.Name = $"FlightMode_Cover_{image.Name}";
                        cover.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
                        cover.Fill.Transparency = 0; // Fully opaque
                        cover.Line.Visible = Office.MsoTriState.msoFalse;

                        // Send to back so it doesn't cover content placeholders
                        cover.ZOrder(Office.MsoZOrderCmd.msoSendToBack);

                        // Tag for identification
                        SetTagValue(cover, FlightModeTag, ImageCoverTag);

                        coversAdded++;
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "FlightModeFeature.EnableCoverLogos.AddCover");
                    }
                }

                return coversAdded > 0;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FlightModeFeature.EnableCoverLogos");
                return false;
            }
        }

        /// <summary>
        /// Enables Flight Mode - Light Overlay method (5% white overlay).
        /// </summary>
        public static bool EnableLightOverlay(PowerPointManager manager)
        {
            return EnableOverlay(manager, 0.05f); // 5% transparency = 95% visible
        }

        /// <summary>
        /// Enables Flight Mode - Heavy Overlay method (15% white overlay).
        /// </summary>
        public static bool EnableHeavyOverlay(PowerPointManager manager)
        {
            return EnableOverlay(manager, 0.15f); // 15% transparency = 85% visible
        }

        /// <summary>
        /// Enables Flight Mode - Hide Master Background.
        /// Iterates all slides and sets FollowMasterBackground to false.
        /// </summary>
        public static bool EnableHideBackground(PowerPointManager manager)
        {
            try
            {
                var pres = manager.GetApplication().ActivePresentation;
                
                // Disable other modes first
                Disable(pres);

                foreach (Slide slide in pres.Slides)
                {
                    try { slide.FollowMasterBackground = Office.MsoTriState.msoFalse; } catch { }
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FlightModeFeature.EnableHideBackground");
                return false;
            }
        }

        /// <summary>
        /// Disables all Flight Mode methods.
        /// </summary>
        public static bool Disable(PowerPointManager manager)
        {
            try
            {
                var pres = manager.GetApplication().ActivePresentation;
                Disable(pres);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FlightModeFeature.Disable");
                return false;
            }
        }

        #region Private Helper Methods

        private static bool EnableOverlay(PowerPointManager manager, float transparency)
        {
            try
            {
                var pres = manager.GetApplication().ActivePresentation;

                // Disable other modes first
                Disable(pres);

                int overlaysAdded = 0;
                var master = pres.SlideMaster;

                try
                {
                    // Create full-slide overlay
                    var overlay = master.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeRectangle,
                        0,
                        0,
                        pres.PageSetup.SlideWidth,
                        pres.PageSetup.SlideHeight);

                    // Style the overlay
                    overlay.Name = "FlightMode_Overlay";
                    overlay.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
                    overlay.Fill.Transparency = transparency;
                    overlay.Line.Visible = Office.MsoTriState.msoFalse;

                    // Send to back
                    overlay.ZOrder(Office.MsoZOrderCmd.msoSendToBack);

                    // Tag for identification
                    SetTagValue(overlay, FlightModeTag, OverlayTag);

                    overlaysAdded++;
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "FlightModeFeature.EnableOverlay.AddOverlay");
                }

                return overlaysAdded > 0;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FlightModeFeature.EnableOverlay");
                return false;
            }
        }

        private static void Disable(Presentation pres)
        {
            DisableOverlays(pres);
            
            // Re-enable background graphics
            foreach (Slide slide in pres.Slides)
            {
                try { slide.FollowMasterBackground = Office.MsoTriState.msoTrue; } catch { }
            }
        }

        private static void DisableOverlays(Presentation pres)
        {
            try
            {
                var master = pres.SlideMaster;
                // Remove overlay shapes (reverse iteration for safe deletion)
                for (int i = master.Shapes.Count; i >= 1; i--)
                {
                    try
                    {
                        var shape = master.Shapes[i];
                        string tagValue = GetTagValue(shape, FlightModeTag);

                        if (tagValue == ImageCoverTag || tagValue == OverlayTag)
                        {
                            shape.Delete();
                        }
                    }
                    catch
                    {
                        // Skip shapes that can't be accessed
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FlightModeFeature.DisableOverlays");
            }
        }

        private static void FindImageShapes(Microsoft.Office.Interop.PowerPoint.Shapes shapes, List<Shape> results)
        {
            try
            {
                var pres = shapes.Parent as Presentation;
                // If parent isn't presentation (e.g. Layout/Master), get it via application
                if (pres == null)
                {
                    try 
                    { 
                        var app = shapes.Application as Microsoft.Office.Interop.PowerPoint.Application;
                        if (app != null) pres = app.ActivePresentation; 
                    } 
                    catch { }
                }

                float slideWidth = pres?.PageSetup.SlideWidth ?? 0;
                float slideHeight = pres?.PageSetup.SlideHeight ?? 0;

                for (int i = 1; i <= shapes.Count; i++)
                {
                    Shape shape = shapes[i];
                    ProcessShape(shape, results, slideWidth, slideHeight);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FlightModeFeature.FindImageShapes");
            }
        }

        private static void ProcessShape(Shape shape, List<Shape> results, float slideWidth, float slideHeight)
        {
            try
            {
                if (shape.Type == Office.MsoShapeType.msoPicture ||
                    shape.Type == Office.MsoShapeType.msoLinkedPicture)
                {
                    // Ignore full-slide background images (anything > 95% of slide dimensions)
                    bool isFullWidth = slideWidth > 0 && shape.Width >= (slideWidth * 0.95f);
                    bool isFullHeight = slideHeight > 0 && shape.Height >= (slideHeight * 0.95f);

                    if (!isFullWidth && !isFullHeight)
                    {
                        results.Add(shape);
                    }
                }
                else if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    foreach (Shape groupShape in shape.GroupItems)
                    {
                        ProcessShape(groupShape, results, slideWidth, slideHeight);
                    }
                }
            }
            catch { }
        }

        private static string GetTagValue(Shape shape, string tagName)
        {
            try
            {
                if (shape.Tags.Count > 0)
                {
                    for (int i = 1; i <= shape.Tags.Count; i++)
                    {
                        if (shape.Tags.Name(i) == tagName)
                        {
                            return shape.Tags.Value(i);
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        private static void SetTagValue(Shape shape, string tagName, string value)
        {
            try
            {
                // Check if tag already exists and remove it
                if (shape.Tags.Count > 0)
                {
                    for (int i = shape.Tags.Count; i >= 1; i--)
                    {
                        if (shape.Tags.Name(i) == tagName)
                        {
                            shape.Tags.Delete(tagName);
                            break;
                        }
                    }
                }

                // Add the tag
                shape.Tags.Add(tagName, value);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FlightModeFeature.SetTagValue");
            }
        }

        #endregion
    }
}
