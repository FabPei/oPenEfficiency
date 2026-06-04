using System;
using System.Windows;
using System.Windows.Controls.Primitives;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(Id = "BtnSlideNotes", Name = "Developer Notes", Tooltip = "Slide Notes (Developer Scratchpad)", IconData = "M19 3h-4.18C14.4.84 13.3 0 12 0c-1.3 0-2.4.84-2.82 2H5c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-7 0c.55 0 1 .45 1 1s-.45 1-1 1-1-.45-1-1 .45-1 1-1zm2 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z", Color = "#8B5CF6", Description = "A developer scratchpad to save private notes, to-do lists, or snippets directly on the slide (hidden from the presentation).", Keywords = "notes, scratchpad, comments, todo, developer, hidden text", IsToggle = true)]
    public static class SlideNotesFeature
    {
        private static UI.Toolbars.SlideNotesWindow _currentWindow;
        
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

        public static void DeleteNotesOnCurrentSlide(PowerPointManager manager)
        {
            try
            {
                var app = manager?.GetApplication();
                if (app?.ActiveWindow?.View?.Slide != null)
                {
                    var slide = app.ActiveWindow.View.Slide as Slide;
                    try { slide.Tags.Delete("oPE_SlideNotes"); } catch { }
                    
                    // If the window is open and showing this slide, clear its text
                    if (_currentWindow != null)
                    {
                        _currentWindow.ClearNotesIfSlideMatches(slide.SlideID);
                    }
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
                        "Are you sure you want to delete all Developer Notes across the entire presentation? This cannot be undone.",
                        "Confirm Deletion",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Warning);

                    if (result == MessageBoxResult.Yes)
                    {
                        int deletedCount = 0;
                        foreach (Slide slide in app.ActivePresentation.Slides)
                        {
                            try 
                            { 
                                // Checking if the tag exists before deleting isn't strictly necessary but good for counting
                                bool hasTag = false;
                                foreach (string tag in slide.Tags)
                                {
                                    if (tag == "oPE_SlideNotes")
                                    {
                                        hasTag = true;
                                        break;
                                    }
                                }
                                if (hasTag)
                                {
                                    slide.Tags.Delete("oPE_SlideNotes"); 
                                    deletedCount++;
                                }
                            } 
                            catch { }
                        }
                        
                        if (_currentWindow != null)
                        {
                            // Force refresh of the current window to clear it if it's open
                            _currentWindow.ForceRefresh();
                        }
                        
                        MessageBox.Show($"Successfully deleted notes from {deletedCount} slide(s).", "Deletion Complete", MessageBoxButton.OK, MessageBoxImage.Information);
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
