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
    /// Feature: Master Covers & Overlays - Temporarily obscure sensitive content (logos, branding)
    /// when presenting in public spaces.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnFlightMode",
        Name = "Master Covers & Overlays",
        SearchAliases = new[] { "Flight Mode", "Disable Flight Mode", "Toggle Flight Mode" },
        Tooltip = "Toggle removable image covers and overlays on slide masters",
        IconData = "M21,16V4H13L11,2H3V16H11L13,18H21M19,14H13.8L12.8,12H5V4H9.2L10.2,6H19V14M17,18V20H7V18H17Z",
        Color = "#F43F5E",
        Description = "Temporarily covers master-slide images or adds a light or heavy white overlay across every slide master.",
        DetailedHelpText = "### Master Covers & Overlays\nTemporarily adds tagged, removable shapes to every slide master. Toggle defaults to Cover Logos, which places white rectangles over master pictures. The context menu also offers Light Overlay (95% of the slide remains visible), Heavy Overlay (85% remains visible), and Disable Master Covers & Overlays. Switching modes first removes the previous Master Covers & Overlays shapes; disabling restores the original masters.",
        MinSelection = 0,
        IsToggle = true,
        RequiredType = OfficeSelectionType.None)]
    public static class FlightModeFeature
    {
        private const string FlightModeTag = oPenEfficiency.Core.Services.FlightModeStatePolicy.FlightModeTag;

        public enum FlightModeType
        {
            None,
            CoverLogos,
            LightOverlay,
            HeavyOverlay,
            HideBackground
        }

        /// <summary>
        /// Wrapper for auto-discovery - toggles flight mode (cover logos).
        /// Note: This feature requires toggle state management, handled by the command adapter.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            // Toggle feature - requires ToggleButton parameter for state sync
            // Handled by the explicit toggle registration in MainSidebar.Commands.cs.
            return false;
        }

        /// <summary>
        /// Gets the current Master Covers & Overlays state of the presentation.
        /// </summary>
        public static FlightModeType GetMode(PowerPointManager manager)
        {
            try
            {
                var pres = manager.GetApplication().ActivePresentation;
                foreach (Master master in GetSlideMasters(pres))
                {
                    for (int i = 1; i <= master.Shapes.Count; i++)
                    {
                        Shape shape = master.Shapes[i];
                        string tagValue = GetTagValue(shape, FlightModeTag);
                        if (tagValue == oPenEfficiency.Core.Services.FlightModeStatePolicy.ImageCoverTag) return FlightModeType.CoverLogos;
                        if (tagValue == oPenEfficiency.Core.Services.FlightModeStatePolicy.OverlayTag)
                            return oPenEfficiency.Core.Services.FlightModeStatePolicy.IsLightOverlay(shape.Fill.Transparency) ? FlightModeType.LightOverlay : FlightModeType.HeavyOverlay;
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
        /// Toggles Master Covers & Overlays (defaults to Cover Logos Only).
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
        /// Enables Master Covers & Overlays - Cover Logos Only method.
        /// Covers all image shapes in master slides with white rectangles.
        /// </summary>
        public static bool EnableCoverLogos(PowerPointManager manager)
        {
            try
            {
                var pres = manager.GetApplication().ActivePresentation;
                int coversAdded = 0;

                TryStartUndoEntry(manager, "FlightModeFeature.EnableCoverLogos.StartNewUndoEntry");

                // First, disable any other flight modes
                DisableOverlays(pres);

                foreach (Master master in GetSlideMasters(pres))
                {
                    var imagesToCover = new List<Shape>();
                    FindImageShapes(master.Shapes, imagesToCover);

                    foreach (Shape image in imagesToCover)
                    {
                        try
                        {
                            var cover = master.Shapes.AddShape(
                                Office.MsoAutoShapeType.msoShapeRectangle,
                                image.Left,
                                image.Top,
                                image.Width,
                                image.Height);
                            cover.Name = $"FlightMode_Cover_{image.Name}";
                            cover.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
                            cover.Fill.Transparency = 0;
                            cover.Line.Visible = Office.MsoTriState.msoFalse;
                            cover.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                            SetTagValue(cover, FlightModeTag, oPenEfficiency.Core.Services.FlightModeStatePolicy.ImageCoverTag);
                            coversAdded++;
                        }
                        catch (Exception ex)
                        {
                            ExceptionLogger.Log(ex, "FlightModeFeature.EnableCoverLogos.AddCover");
                        }
                    }
                }

                if (coversAdded == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "No picture shapes were found on the presentation's slide masters.",
                        "Master Covers & Overlays",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
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
        /// Enables Master Covers & Overlays - Light Overlay method (5% white tint, 95% transparency).
        /// </summary>
        public static bool EnableLightOverlay(PowerPointManager manager)
        {
            return EnableOverlay(manager, oPenEfficiency.Core.Services.FlightModeStatePolicy.LightOverlayTransparency);
        }

        /// <summary>
        /// Enables Master Covers & Overlays - Heavy Overlay method (15% white tint, 85% transparency).
        /// </summary>
        public static bool EnableHeavyOverlay(PowerPointManager manager)
        {
            return EnableOverlay(manager, oPenEfficiency.Core.Services.FlightModeStatePolicy.HeavyOverlayTransparency);
        }

        /// <summary>
        /// Enables Master Covers & Overlays - Hide Master Background (native method).
        /// Currently not used - overlay methods are preferred.
        /// </summary>
        public static bool EnableHideBackground(PowerPointManager manager)
        {
            // Placeholder - PowerPoint doesn't have a simple ShowBackground property
            // Use overlay methods instead
            System.Windows.Forms.MessageBox.Show(
                "Hide Master Background is not available. Please use Cover Logos or Overlay methods.",
                "Master Covers & Overlays",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
            return false;
        }

        /// <summary>
        /// Disables all Master Covers & Overlays methods.
        /// </summary>
        public static bool Disable(PowerPointManager manager)
        {
            try
            {
                var pres = manager.GetApplication().ActivePresentation;
                TryStartUndoEntry(manager, "FlightModeFeature.Disable.StartNewUndoEntry");
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
                TryStartUndoEntry(manager, "FlightModeFeature.EnableOverlay.StartNewUndoEntry");

                // Disable other modes first
                Disable(pres);

                int overlaysAdded = 0;
                foreach (Master master in GetSlideMasters(pres))
                {
                    try
                    {
                        var overlay = master.Shapes.AddShape(
                            Office.MsoAutoShapeType.msoShapeRectangle,
                            0,
                            0,
                            pres.PageSetup.SlideWidth,
                            pres.PageSetup.SlideHeight);
                        overlay.Name = "FlightMode_Overlay";
                        overlay.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
                        overlay.Fill.Transparency = transparency;
                        overlay.Line.Visible = Office.MsoTriState.msoFalse;
                        overlay.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                        SetTagValue(overlay, FlightModeTag, oPenEfficiency.Core.Services.FlightModeStatePolicy.OverlayTag);
                        overlaysAdded++;
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "FlightModeFeature.EnableOverlay.AddOverlay");
                    }
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
        }

        private static void DisableOverlays(Presentation pres)
        {
            try
            {
                foreach (Master master in GetSlideMasters(pres))
                {
                    for (int i = master.Shapes.Count; i >= 1; i--)
                    {
                        try
                        {
                            var shape = master.Shapes[i];
                            string tagValue = GetTagValue(shape, FlightModeTag);
                            if (tagValue == oPenEfficiency.Core.Services.FlightModeStatePolicy.ImageCoverTag || tagValue == oPenEfficiency.Core.Services.FlightModeStatePolicy.OverlayTag) shape.Delete();
                        }
                        catch (Exception ex) { ExceptionLogger.Log(ex, "FlightModeFeature.DisableOverlays.Shape"); }
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
                for (int i = 1; i <= shapes.Count; i++)
                {
                    Shape shape = shapes[i];
                    if (shape.Type == Office.MsoShapeType.msoPicture ||
                        shape.Type == Office.MsoShapeType.msoLinkedPicture)
                    {
                        results.Add(shape);
                    }
                    else if (shape.Type == Office.MsoShapeType.msoGroup)
                    {
                        // Recurse into groups - GroupItems returns GroupShapes which can be iterated
                        try
                        {
                            foreach (Shape groupShape in shape.GroupItems)
                            {
                                if (groupShape.Type == Office.MsoShapeType.msoPicture ||
                                    groupShape.Type == Office.MsoShapeType.msoLinkedPicture)
                                {
                                    results.Add(groupShape);
                                }
                            }
                        }
                        catch
                        {
                            // Skip groups that can't be accessed
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FlightModeFeature.FindImageShapes");
            }
        }

        private static IEnumerable<Master> GetSlideMasters(Presentation presentation)
        {
            for (int i = 1; i <= presentation.Designs.Count; i++)
                yield return presentation.Designs[i].SlideMaster;
        }

        private static void TryStartUndoEntry(PowerPointManager manager, string context)
        {
            try { manager.GetApplication().StartNewUndoEntry(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, context); }
        }

        // Measured 2026-10-10 on the real host (QA-S9 batch 52 office leg): PowerPoint
        // STORES shape tag names UPPERCASE (Tags.Add("oPE_FlightMode") reads back as
        // "OPE_FLIGHTMODE" through the same C# interop), so a case-sensitive name
        // comparison never matches its own written tag - GetMode stayed None after a
        // successful Enable and Disable could not find its own overlay to remove.
        // Both tag comparisons therefore match OrdinalIgnoreCase.
        private static string GetTagValue(Shape shape, string tagName)
        {
            try
            {
                if (shape.Tags.Count > 0)
                {
                    for (int i = 1; i <= shape.Tags.Count; i++)
                    {
                        if (string.Equals(shape.Tags.Name(i), tagName,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            return shape.Tags.Value(i);
                        }
                    }
                }
            }
            catch { /* Missing or stale tags mean no persisted flight-mode state is available. */ }
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
                        if (string.Equals(shape.Tags.Name(i), tagName,
                            StringComparison.OrdinalIgnoreCase))
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
