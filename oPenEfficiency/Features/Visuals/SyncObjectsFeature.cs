using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.UI;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Distribute shapes to multiple slides and keep them in sync using a unique SyncID.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnSyncObjects",
        Name = "Sync Shapes Across Slides",
        Tooltip = "Sync shapes across slides",
        SearchAliases = new[] { "Sync Across Slides", "Shape Synchronizer", "Sync Objects" },
        IconData = "M12,4L14.07,6.07L16.21,3.93L18.34,6.07L20.48,3.93L22.62,6.07L20.48,8.21L22.62,10.34L20.48,12.48L18.34,10.34L16.21,12.48L14.07,10.34L12,12.48V4M4,12V20H12V12H4M4,4V10H10V4H4M14,14V20H20V14H14Z",
        Color = "#F43F5E",
        Description = "Opens dialog to distribute and sync shapes across multiple slides.",
        DetailedHelpText = "### Sync Shapes Across Slides\nDistributes selected shapes to chosen slides, then synchronizes tagged copies. Right-click can synchronize format, position, content, or all properties directly from the one selected tagged source shape; target-slide distribution remains in the full dialog.",
        MinSelection = 1)]
    public static class SyncObjectsFeature
    {
        public const string SyncIdTag = "oPE_SyncID";

        private static SyncObjectsWindow _window;

        /// <summary>
        /// Wrapper for auto-discovery - returns false as Sync Objects requires dialog interaction.
        /// Manual switch in MainSidebar.xaml.cs handles opening the SyncObjectsWindow dialog.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            ExecuteDialog(manager);
            return true;
        }

        public static void ExecuteDialog(PowerPointManager manager)
        {
            if (manager == null) return;

            if (_window != null && !_window.IsLoaded)
                _window = null;

            if (_window != null)
            {
                var existingApp = manager.GetApplication();
                if (existingApp != null) new System.Windows.Interop.WindowInteropHelper(_window).Owner = (IntPtr)existingApp.HWND;
                if (!_window.IsVisible) _window.Show();
                if (_window.WindowState == System.Windows.WindowState.Minimized)
                    _window.WindowState = System.Windows.WindowState.Normal;
                _window.Activate();
                _window.Focus();
                return;
            }

            try
            {
                var window = new SyncObjectsWindow(manager);
                _window = window;
                var app = manager.GetApplication();
                if (app != null) new System.Windows.Interop.WindowInteropHelper(window).Owner = (IntPtr)app.HWND;
                window.Closed += (sender, args) =>
                {
                    if (ReferenceEquals(_window, window)) _window = null;
                };
                window.Show();
                if (!_window.IsLoaded || !_window.IsVisible)
                    throw new InvalidOperationException(
                        "Sync Shapes window closed before it became visible.");
            }
            catch
            {
                try { _window?.Close(); }
                catch (Exception closeEx)
                {
                    Utils.ExceptionLogger.Log(closeEx, "SyncObjectsFeature.ExecuteDialog.FailureCleanup");
                }
                _window = null;
                throw;
            }
        }

        public static bool DistributeShapesToSlides(PowerPointManager manager, List<int> targetSlideIndices)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                var selection = manager.GetSelectedShapes();
                if (selection == null || selection.Count == 0) return false;

                foreach (Shape shape in selection)
                {
                    TryAssignSyncId(shape);
                }

                var pres = app.ActivePresentation;
                app.ActiveWindow.Selection.Copy();

                foreach (int slideIndex in targetSlideIndices)
                {
                    try
                    {
                        if (slideIndex >= 1 && slideIndex <= pres.Slides.Count)
                        {
                            Slide targetSlide = pres.Slides[slideIndex];
                            targetSlide.Shapes.Paste();
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to paste on slide {slideIndex}: {ex.Message}");
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Distribute failed: {ex.Message}");
                return false;
            }
        }

        public static bool SyncShapeProperties(PowerPointManager manager, bool syncFormat, bool syncPosition, bool syncContent)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                var selection = manager.GetSelectedShapes();
                if (selection == null || selection.Count != 1) return false;

                Shape sourceShape = selection[1];
                string syncId = GetSyncId(sourceShape);
                if (string.IsNullOrEmpty(syncId)) return false;

                var pres = app.ActivePresentation;
                
                float left = sourceShape.Left;
                float top = sourceShape.Top;
                float width = sourceShape.Width;
                float height = sourceShape.Height;
                string text = "";
                bool hasText = sourceShape.HasTextFrame == Office.MsoTriState.msoTrue && sourceShape.TextFrame.HasText == Office.MsoTriState.msoTrue;
                if (hasText) text = sourceShape.TextFrame.TextRange.Text;

                sourceShape.PickUp();

                int currentSlideIndex = -1;
                try
                {
                    if (app.ActiveWindow?.View?.Slide != null)
                    {
                        var activeSlide = app.ActiveWindow.View.Slide as Slide;
                        if (activeSlide != null) currentSlideIndex = activeSlide.SlideIndex;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"SyncObjectsFeature.currentSlideIndex error: {ex.Message}");
                }

                foreach (Slide slide in pres.Slides)
                {
                    foreach (Shape targetShape in slide.Shapes)
                    {
                        if (currentSlideIndex > 0 && slide.SlideIndex == currentSlideIndex && targetShape.Id == sourceShape.Id)
                            continue;

                        if (GetSyncId(targetShape) == syncId)
                        {
                            if (syncFormat) try { targetShape.Apply(); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"SyncObjectsFeature.Apply error: {ex.Message}"); }
                            if (syncPosition)
                            {
                                try
                                {
                                    targetShape.Left = left;
                                    targetShape.Top = top;
                                    targetShape.Width = width;
                                    targetShape.Height = height;
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"SyncObjectsFeature.position error: {ex.Message}");
                                }
                            }
                            if (syncContent && hasText)
                            {
                                try
                                {
                                    if (targetShape.HasTextFrame == Office.MsoTriState.msoTrue)
                                    {
                                        targetShape.TextFrame.TextRange.Text = text;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"SyncObjectsFeature.content error: {ex.Message}");
                                }
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Sync failed: {ex.Message}");
                return false;
            }
        }

        public static void TryAssignSyncId(Shape shape)
        {
            try
            {
                string altText = shape.AlternativeText ?? "";
                if (!altText.Contains(SyncIdTag))
                {
                    string id = Guid.NewGuid().ToString();
                    shape.AlternativeText = $"{SyncIdTag}:{id}|{altText}";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SyncObjectsFeature.TryAssignSyncId error: {ex.Message}");
            }
        }

        public static string GetSyncId(Shape shape)
        {
            try
            {
                string alt = shape.AlternativeText;
                if (!string.IsNullOrEmpty(alt) && alt.Contains(SyncIdTag))
                {
                    int start = alt.IndexOf(SyncIdTag + ":") + SyncIdTag.Length + 1;
                    int end = alt.IndexOf("|", start);
                    if (end > start) return alt.Substring(start, end - start);
                    return alt.Substring(start);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SyncObjectsFeature.GetSyncId error: {ex.Message}");
            }
            return null;
        }
    }
}
