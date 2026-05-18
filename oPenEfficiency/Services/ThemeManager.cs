using System;
using System.Linq;
using System.Windows;

namespace oPenEfficiency.Services
{
    public static class ThemeManager
    {
        public static event EventHandler<string> ThemeChanged;

        public static void ApplyTheme(string themeName)
        {
            if (string.IsNullOrEmpty(themeName)) themeName = "Dark";

            var app = Application.Current;
            if (app == null) return;

            // 1. Find all existing theme and controls dictionaries
            var existingThemes = app.Resources.MergedDictionaries
                .Where(d => d.Source != null && d.Source.OriginalString.Contains("UI/Styles/Themes/"))
                .ToList();
            
            var existingControls = app.Resources.MergedDictionaries
                .Where(d => d.Source != null && d.Source.OriginalString.Contains("UI/Styles/Controls.xaml"))
                .ToList();

            System.Diagnostics.Debug.WriteLine($"ApplyTheme: Removing {existingThemes.Count} theme dictionaries and {existingControls.Count} controls dictionaries.");

            // 2. Remove all of them
            foreach (var theme in existingThemes) app.Resources.MergedDictionaries.Remove(theme);
            foreach (var ctrl in existingControls) app.Resources.MergedDictionaries.Remove(ctrl);

            // 3. Add the new theme dictionary
            try
            {
                var themeUri = new Uri($"pack://application:,,,/oPenEfficiency;component/UI/Styles/Themes/{themeName}.xaml", UriKind.RelativeOrAbsolute);
                System.Diagnostics.Debug.WriteLine($"ApplyTheme: Loading new theme from {themeUri}");
                
                var newTheme = new ResourceDictionary { Source = themeUri };
                app.Resources.MergedDictionaries.Add(newTheme);

                // 4. Re-add Controls.xaml AFTER the theme
                var controlsUri = new Uri("pack://application:,,,/oPenEfficiency;component/UI/Styles/Controls.xaml", UriKind.RelativeOrAbsolute);
                System.Diagnostics.Debug.WriteLine($"ApplyTheme: Re-loading controls from {controlsUri}");
                
                var newControls = new ResourceDictionary { Source = controlsUri };
                app.Resources.MergedDictionaries.Add(newControls);

                // 5. Force refresh for all open WPF windows
                foreach (Window window in app.Windows)
                {
                    try
                    {
                        RefreshDictionary(window.Resources);
                        
                        // Invalidate visual to force repaint
                        window.InvalidateVisual();
                    }
                    catch { }
                }
                
                System.Diagnostics.Debug.WriteLine($"Theme '{themeName}' applied and all windows refreshed.");

                // 6. Notify listeners
                ThemeChanged?.Invoke(null, themeName);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error applying theme {themeName}: {ex.Message}");
                // Fallback to Dark if Light fails
                if (themeName != "Dark") ApplyTheme("Dark");
            }
        }

        private static void RefreshDictionary(ResourceDictionary resources)
        {
            if (resources == null) return;

            // 1. Find and replace Controls.xaml in this dictionary or its merged dictionaries
            var controlsDict = resources.MergedDictionaries
                .FirstOrDefault(d => d.Source != null && d.Source.OriginalString.Contains("UI/Styles/Controls.xaml"));

            if (controlsDict != null)
            {
                int index = resources.MergedDictionaries.IndexOf(controlsDict);
                resources.MergedDictionaries.RemoveAt(index);
                resources.MergedDictionaries.Insert(index, new ResourceDictionary
                {
                    Source = new Uri("pack://application:,,,/oPenEfficiency;component/UI/Styles/Controls.xaml", UriKind.RelativeOrAbsolute)
                });
            }

            // 2. Also check nested merged dictionaries recursively
            foreach (var merged in resources.MergedDictionaries)
            {
                // Note: We don't want to recurse into the one we just re-added
                if (merged.Source != null && merged.Source.OriginalString.Contains("UI/Styles/Controls.xaml"))
                    continue;
                
                RefreshDictionary(merged);
            }
        }
    }
}
