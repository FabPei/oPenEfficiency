using System;
using System.Windows;
using System.Windows.Controls.Primitives;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnSlideNotes",
        Name = "Private Review Notes",
        Tooltip = "Private review-note scratchpad",
        SearchAliases = new[] { "Private oPE Comments", "Slide Notes" },
        IconData = "M5,2H19A3,3 0 0 1 22,5V11H20V5A1,1 0 0 0 19,4H5A1,1 0 0 0 4,5V18.6L6.6,16H11V18H7.4L2,23.4V5A3,3 0 0 1 5,2Z M7,7H17V9H7Z M7,11H13V13H7Z M12.3,18.7L13.7,17.3L16,19.6L21.3,14.3L22.7,15.7L16,22.4Z",
        Color = "#8B5CF6",
        Description = "Stores multiple private review notes per slide, with Open or Done status and optional shape anchoring.",
        DetailedHelpText = "### Private Review Notes\nCreate multiple private review notes on the current slide, optionally attach them to a selected shape, and manage them as **Open** or **Done**. Right-click the feature to open the presentation-wide overview or delete notes with confirmation. Notes are stored as presentation metadata and stay out of PowerPoint's review UI and slideshow, but they are not encrypted and are not a security boundary.",
        IsToggle = true,
        RequiredType = OfficeSelectionType.None)]
    public static class SlideNotesFeature
    {
        private static UI.Toolbars.SlideNotesWindow _currentWindow;
        private static UI.Toolbars.PrivateCommentOverlayWindow _currentOverlay;
        private static UI.Toolbars.PrivateCommentsOverviewWindow _overviewWindow;
        
        public static bool Execute(PowerPointManager manager)
        {
            return false; // trigger manual handling
        }

        public static void Execute(PowerPointManager manager, ToggleButton toggleBtn)
        {
            try
            {
                if (_currentWindow != null)
                {
                    _currentWindow.Close();
                    _currentWindow = null;
                    if (toggleBtn != null) toggleBtn.IsChecked = false;
                }
                else
                {
                    _currentWindow = new UI.Toolbars.SlideNotesWindow(manager, toggleBtn);
                    var helper = new System.Windows.Interop.WindowInteropHelper(_currentWindow);
                    helper.Owner = (IntPtr)manager.GetApplication().HWND;
                    _currentWindow.Show();
                    if (toggleBtn != null) toggleBtn.IsChecked = true;
                }
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "SlideNotesFeature.Execute");
            }
        }
        
        public static void SetWindowNull()
        {
            _currentWindow = null;
        }

        public static void ToggleOverlay(PowerPointManager manager)
        {
            try
            {
                if (_currentOverlay != null)
                {
                    _currentOverlay.Close();
                    _currentOverlay = null;
                    return;
                }

                _currentOverlay = new UI.Toolbars.PrivateCommentOverlayWindow(manager);
                var app = manager.GetApplication();
                var helper = new System.Windows.Interop.WindowInteropHelper(_currentOverlay);
                helper.Owner = (IntPtr)app.HWND;
                _currentOverlay.Show();
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "SlideNotesFeature.ToggleOverlay");
            }
        }

        public static void SetOverlayNull()
        {
            _currentOverlay = null;
        }

        public static void OpenOverview(PowerPointManager manager)
        {
            try
            {
                if (_overviewWindow != null)
                {
                    _overviewWindow.Activate();
                    return;
                }
                _overviewWindow = new UI.Toolbars.PrivateCommentsOverviewWindow(manager);
                _overviewWindow.Closed += (sender, args) => _overviewWindow = null;
                var helper = new System.Windows.Interop.WindowInteropHelper(_overviewWindow);
                helper.Owner = (IntPtr)manager.GetApplication().HWND;
                _overviewWindow.Show();
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "SlideNotesFeature.OpenOverview");
            }
        }

        public static void DeleteNotesOnCurrentSlide(PowerPointManager manager)
        {
            try
            {
                var app = manager?.GetApplication();
                if (app?.ActiveWindow?.View?.Slide != null)
                {
                    var slide = app.ActiveWindow.View.Slide as Slide;
                    int count = PrivateCommentsService.GetCommentsForSlide(app.ActivePresentation, slide).Count;
                    if (count == 0) return;
                    if (MessageBox.Show($"Delete {count} private review note(s) from the current slide? This cannot be undone.",
                        "Delete Private Comments", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return;
                    PrivateCommentsService.DeleteForSlides(app.ActivePresentation, new[] { slide });
                    
                    // If the window is open and showing this slide, clear its text
                    if (_currentWindow != null)
                    {
                        _currentWindow.ClearNotesIfSlideMatches(slide.SlideID);
                    }
                    MessageBox.Show("Private review notes on the current slide were deleted.", "Deletion Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "SlideNotesFeature.DeleteNotesOnCurrentSlide");
            }
        }

        public static void DeleteAllNotes(PowerPointManager manager)
        {
            try
            {
                var app = manager?.GetApplication();
                if (app?.ActivePresentation != null)
                {
                    var result = MessageBox.Show(
                        "Are you sure you want to delete all private review notes across the entire presentation? This cannot be undone.",
                        "Confirm Deletion",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Warning);

                    if (result == MessageBoxResult.Yes)
                    {
                        PrivateCommentsService.DeleteAll(app.ActivePresentation);
                        
                        if (_currentWindow != null)
                        {
                            // Force refresh of the current window to clear it if it's open
                            _currentWindow.ForceRefresh();
                        }
                        
                        MessageBox.Show("Successfully deleted all private review notes.", "Deletion Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "SlideNotesFeature.DeleteAllNotes");
            }
        }
    }
}
