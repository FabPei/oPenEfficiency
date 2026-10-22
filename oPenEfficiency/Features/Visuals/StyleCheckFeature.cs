using System;
using System.Windows.Interop;
using oPenEfficiency.UI;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    public enum StyleCheckScope
    {
        CurrentSlide,
        SelectedSlides,
        EntirePresentation
    }

    /// <summary>
    /// Feature: Style Check - analyzes and displays presentation style information.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStyleCheck",
        Name = "Style Check",
        Tooltip = "Style check",
        SearchAliases = new[] { "Presentation Style Check", "Style & Consistency Check" },
        IconData = "M9,20.16L9,18.13C9,18.13 8.5,18.15 8.21,17.92C7.79,17.5 7.85,16.84 7.85,16.84L7.92,14.39L9.95,14.45L9.88,16.89L10.2,17.21C10.49,17.44 11,17.4 11,17.4L11,15.37L13,15.43L12.93,17.88C12.9,18.81 12.37,19.75 11.5,20.1C10.5,20.5 9,20.16 9,20.16M12,2A10,10 0 0,0 2,12A10,10 0 0,0 12,22A10,10 0 0,0 22,12A10,10 0 0,0 12,2M12,4A8,8 0 0,1 20,12A8,8 0 0,1 12,20A8,8 0 0,1 4,12A8,8 0 0,1 12,4M7,7H9V11H7V7M11,7H17V9H11V7M11,11H15V13H11V11Z",
        Color = "#F43F5E",
        Description = "Checks the current slide, selected slides, or the entire presentation for style and consulting-deck consistency issues.",
        DetailedHelpText = "### Style Check\nChoose **Current Slide**, **Selected Slides**, or **Entire Presentation** before scanning. Style Check identifies off-brand fonts and colors, spacing and text-overflow issues, distorted images, overlapping text, missing or overly long titles, missing source lines on chart or table slides, footer and pagination inconsistencies, and common wording defects. Use **Jump** to select the affected object, **Ignore Rule** to exempt only that rule for the object or slide, and **Fix Safe Issues...** to preview and apply low-risk color and wording corrections. Right-click the Sidebar button to open a specific scope directly.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class StyleCheckFeature
    {
        private static StyleCheckWindow _window;

        /// <summary>
        /// Auto-discovery entry - runs the style check window and returns true
        /// Manual switch in MainSidebar.xaml.cs handles opening the StyleCheckWindow dialog.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            ExecuteDialog(manager);
            return true;
        }

        public static void ExecuteDialog(
            PowerPointManager powerPointManager,
            System.Windows.Controls.Primitives.ToggleButton toggleBtn = null,
            StyleCheckScope scope = StyleCheckScope.SelectedSlides)
        {
            try
            {
                if (toggleBtn != null && toggleBtn.IsChecked == false)
                {
                    if (_window != null && _window.IsLoaded)
                    {
                        _window.Close();
                    }
                    return;
                }

                if (_window != null && _window.IsLoaded)
                {
                    if (toggleBtn != null) toggleBtn.IsChecked = true;
                    _window.SetScope(scope);
                    _window.Activate();
                    return;
                }

                _window = new StyleCheckWindow(powerPointManager, scope);

                if (toggleBtn != null)
                {
                    toggleBtn.IsChecked = true;
                    _window.Closed += (s, e) => { toggleBtn.IsChecked = false; };
                }

                // Set owner if possible
                try
                {
                    IntPtr hwnd = (IntPtr)powerPointManager.GetApplication().HWND;
                    if (hwnd != IntPtr.Zero)
                    {
                        var helper = new WindowInteropHelper(_window);
                        helper.Owner = hwnd;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"StyleCheckFeature.WindowOwner error: {ex.Message}");
                }

                _window.Show();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Error opening Style Check window: " + ex.Message);
            }
        }
    }
}
