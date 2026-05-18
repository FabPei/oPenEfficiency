using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.UI
{
    public partial class SettingsWindow : Window
    {
        public delegate void SettingsAppliedEventHandler(object sender, EventArgs e);
        public event SettingsAppliedEventHandler SettingsApplied;

        private SettingsViewModel _viewModel;
        private PowerPointManager _ppManager;

        // Which colour slot swatch clicks target:  0=Bg  1=Font  2=BtnBg
        private int _colorTarget = 0;

        private static readonly string[] TargetLabels =
        {
            "Background Colour",
            "Section Header Text Colour",
            "Button Background Colour"
        };

        public SettingsWindow() : this(null) { }

        public SettingsWindow(PowerPointManager ppManager)
        {
            InitializeComponent();
            _ppManager = ppManager;
            _viewModel = new SettingsViewModel();
            DataContext = _viewModel;

            _viewModel.CloseAction   = () => this.Close();
            _viewModel.RebuildAction = () => SettingsApplied?.Invoke(this, EventArgs.Empty);

            Loaded += (_, __) =>
            {
                PopulateThemeSwatches();
                UpdateContrastWarning();
                UpdateCacheSizeDisplay();
                LoadSharedFolders();

                if (_viewModel != null)
                {
                    switch (_viewModel.StickyNoteStyle)
                    {
                        case "Consulting": RbStickyConsulting.IsChecked = true; break;
                        case "GreenB": RbStickyGreenB.IsChecked = true; break;
                        default: RbStickyTheme.IsChecked = true; break;
                    }
                }
            };
        }

        private void StickyStyle_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton rb && rb.Tag is string style && _viewModel != null)
            {
                _viewModel.StickyNoteStyle = style;
            }
        }

        // ══════════════════════════════════════════════════════════════════════
        // Radio buttons — select which slot swatch clicks update
        // ══════════════════════════════════════════════════════════════════════

        private void ColorTarget_Changed(object sender, RoutedEventArgs e)
        {
            // Guard: fires during InitializeComponent before named elements exist
            if (RadioBg == null || TargetHint == null) return;

            if (RadioBg.IsChecked  == true) _colorTarget = 0;
            else if (RadioFont.IsChecked == true) _colorTarget = 1;
            else                                  _colorTarget = 2;

            TargetHint.Text = TargetLabels[_colorTarget];
        }

        // ══════════════════════════════════════════════════════════════════════
        // Swatch clicks — route to the currently selected target
        // ══════════════════════════════════════════════════════════════════════

        private void Swatch_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string hex)
            {
                switch (_colorTarget)
                {
                    case 0: ApplyBackgroundColor(hex);  break;
                    case 1: ApplyFontColor(hex);       break;
                    case 2: ApplyBtnBgColor(hex);      break;
                }
            }
        }

        private Button MakeSwatch(string hex, string tooltip)
        {
            SolidColorBrush bg;
            try   { bg = (SolidColorBrush)new BrushConverter().ConvertFrom(hex); }
            catch { bg = Brushes.Gray; }

            var btn = new Button
            {
                Style      = (Style)FindResource("SwatchBtn"),
                Background = bg,
                ToolTip    = $"{tooltip}  {hex}",
                Tag        = hex
            };
            btn.Click += Swatch_Click;
            return btn;
        }

        // ══════════════════════════════════════════════════════════════════════
        // Hex text-box handlers
        // ══════════════════════════════════════════════════════════════════════

        private void BgHexBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox tb) ApplyBackgroundColor(tb.Text);
        }

        private void FontHexBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox tb) ApplyFontColor(tb.Text);
        }

        private void BtnBgHexBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox tb) ApplyBtnBgColor(tb.Text);
        }

        // ══════════════════════════════════════════════════════════════════════
        // Apply helpers
        // ══════════════════════════════════════════════════════════════════════

        private void ApplyBackgroundColor(string hex)
        {
            if (!IsValidColor(hex)) return;
            _viewModel.BackgroundColor = hex;
            UpdateContrastWarning();
        }

        private void ApplyFontColor(string hex)
        {
            if (!IsValidColor(hex)) return;
            _viewModel.SectionFontColor = hex;
            UpdateContrastWarning();
        }

        private void ApplyBtnBgColor(string hex)
        {
            // Also accept "Transparent" literal
            if (hex == "Transparent" || IsValidColor(hex))
                _viewModel.ButtonBackgroundColor = hex;
        }

        private static bool IsValidColor(string hex)
        {
            try   { new BrushConverter().ConvertFrom(hex); return true; }
            catch { return false; }
        }

        // ══════════════════════════════════════════════════════════════════════
        // Font size slider
        // ══════════════════════════════════════════════════════════════════════

        private void FontSizeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            // Guard for InitializeComponent timing
            if (_viewModel == null) return;
            _viewModel.SectionFontSize = Math.Round(e.NewValue);
        }

        // ══════════════════════════════════════════════════════════════════════
        // WCAG 2.1 contrast ratio + warning banner
        // ══════════════════════════════════════════════════════════════════════

        private void UpdateContrastWarning()
        {
            if (ContrastWarningBanner == null) return;

            double ratio = GetContrastRatio(_viewModel.BackgroundColor, _viewModel.SectionFontColor);
            if (ratio < 0) { ContrastWarningBanner.Visibility = Visibility.Collapsed; return; }

            string ratioText = ratio.ToString("F2") + ":1";

            if (ratio >= 4.5)
            {
                ContrastWarningBanner.Visibility = Visibility.Collapsed;
            }
            else if (ratio >= 3.0)
            {
                // Amber — fails AA normal text but passes AA large text
                ShowWarning(
                    background:   Color.FromRgb(0x3A, 0x2A, 0x00),
                    border:       Color.FromRgb(0xFF, 0xB9, 0x00),
                    titleColor:   Color.FromRgb(0xFF, 0xB9, 0x00),
                    detailColor:  Color.FromRgb(0xFF, 0xD9, 0x66),
                    title:        $"⚠  Low Contrast  ({ratioText})",
                    detail:       "The contrast ratio between section text and background is below the " +
                                  "WCAG AA threshold (4.5:1) for normal-sized text. Labels may be hard to read."
                );
            }
            else
            {
                // Red — fails all WCAG levels
                ShowWarning(
                    background:   Color.FromRgb(0x3A, 0x10, 0x10),
                    border:       Color.FromRgb(0xE5, 0x73, 0x73),
                    titleColor:   Color.FromRgb(0xE5, 0x73, 0x73),
                    detailColor:  Color.FromRgb(0xFF, 0xAA, 0xAA),
                    title:        $"✕  Very Low Contrast  ({ratioText})",
                    detail:       "The contrast ratio is critically low. Section labels will be nearly invisible " +
                                  "against this background. Please choose colours that differ more."
                );
            }
        }

        private void ShowWarning(Color background, Color border, Color titleColor, Color detailColor,
                                 string title, string detail)
        {
            ContrastWarningBanner.Visibility  = Visibility.Visible;
            ContrastWarningBanner.Background  = new SolidColorBrush(background);
            ContrastWarningBanner.BorderBrush = new SolidColorBrush(border);
            ContrastWarningTitle.Foreground   = new SolidColorBrush(titleColor);
            ContrastWarningTitle.Text         = title;
            ContrastWarningDetail.Foreground  = new SolidColorBrush(detailColor);
            ContrastWarningDetail.Text        = detail;
        }

        private static double GetContrastRatio(string hex1, string hex2)
        {
            try
            {
                double l1 = RelativeLuminance(hex1);
                double l2 = RelativeLuminance(hex2);
                double lighter = Math.Max(l1, l2);
                double darker  = Math.Min(l1, l2);
                return (lighter + 0.05) / (darker + 0.05);
            }
            catch { return -1; }
        }

        private static double RelativeLuminance(string hex)
        {
            Color c;
            if (hex == "Transparent")
            {
                // Treat transparent as the default sidebar background for contrast warning purposes
                c = (Color)ColorConverter.ConvertFromString("#252525");
            }
            else
            {
                c = (Color)ColorConverter.ConvertFromString(hex);
            }

            return 0.2126 * Linearize(c.R / 255.0)
                 + 0.7152 * Linearize(c.G / 255.0)
                 + 0.0722 * Linearize(c.B / 255.0);
        }

        private static double Linearize(double v)
            => v <= 0.04045 ? v / 12.92 : Math.Pow((v + 0.055) / 1.055, 2.4);

        // ══════════════════════════════════════════════════════════════════════
        // Theme swatch population
        // ══════════════════════════════════════════════════════════════════════

        private void PopulateThemeSwatches()
        {
            ThemeSwatchPanel.Children.Clear();
            ThemePlaceholder.Visibility = Visibility.Collapsed;

            var colors = GetThemeColors();
            if (colors == null || colors.Count == 0)
            {
                ThemePlaceholder.Visibility = Visibility.Visible;
                return;
            }
            foreach (var (name, hex) in colors)
                ThemeSwatchPanel.Children.Add(MakeSwatch(hex, name));
        }

        private List<(string Name, string Hex)> GetThemeColors()
        {
            try
            {
                var app  = _ppManager?.GetApplication();
                var pres = app?.ActivePresentation;
                if (pres == null) return null;

                var scheme = pres.SlideMaster.Theme.ThemeColorScheme;
                var result = new List<(string, string)>();

                var indices = new[]
                {
                    (Office.MsoThemeColorSchemeIndex.msoThemeDark1,            "Dark 1"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeLight1,           "Light 1"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeDark2,            "Dark 2"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeLight2,           "Light 2"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeAccent1,          "Accent 1"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeAccent2,          "Accent 2"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeAccent3,          "Accent 3"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeAccent4,          "Accent 4"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeAccent5,          "Accent 5"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeAccent6,          "Accent 6"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeHyperlink,        "Hyperlink"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink,"Followed Hyperlink"),
                };

                foreach (var (idx, label) in indices)
                {
                    try
                    {
                        int rgb = scheme.Colors(idx).RGB; // OLE_COLOR = BGR
                        int r = rgb & 0xFF;
                        int g = (rgb >> 8) & 0xFF;
                        int b = (rgb >> 16) & 0xFF;
                        result.Add((label, $"#{r:X2}{g:X2}{b:X2}"));
                    }
                    catch { }
                }
                return result;
            }
            catch { return null; }
        }

        // ══════════════════════════════════════════════════════════════════════
        // Layout tab — remove feature from section
        // ══════════════════════════════════════════════════════════════════════

        private void RemoveFeature_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is SidebarFeature feature)
            {
                var sectionBorder = FindSectionBorder(btn);
                if (sectionBorder?.DataContext is SidebarSection section)
                    _viewModel.RemoveFeatureFromSection(section, feature);
            }
        }

        private Border FindSectionBorder(DependencyObject child)
        {
            DependencyObject current = VisualTreeHelper.GetParent(child);
            while (current != null)
            {
                if (current is Border b && b.DataContext is SidebarSection) return b;
                current = VisualTreeHelper.GetParent(current);
            }
            return null;
        }

        private T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parent = VisualTreeHelper.GetParent(child);
            if (parent == null) return null;
            if (parent is T typed) return typed;
            return FindVisualParent<T>(parent);
        }

        private void PickSectionColor_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is SidebarSection section)
            {
                try
                {
                    using (var cd = new System.Windows.Forms.ColorDialog())
                    {
                        cd.AllowFullOpen = true;
                        cd.AnyColor = true;
                        if (!string.IsNullOrEmpty(section.Color))
                        {
                            try {
                                var wpfc = (Color)ColorConverter.ConvertFromString(section.Color);
                                cd.Color = System.Drawing.Color.FromArgb(wpfc.A, wpfc.R, wpfc.G, wpfc.B);
                            } catch { }
                        }

                        if (cd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            section.Color = $"#{cd.Color.R:X2}{cd.Color.G:X2}{cd.Color.B:X2}";
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not open color dialog: " + ex.Message);
                }
            }
        }
        
        private void PickStickyBgColor_Click(object sender, RoutedEventArgs e)
        {
            if (_viewModel == null) return;
            string color = PickColor(_viewModel.StickyNoteBackgroundColor);
            if (color != null) _viewModel.StickyNoteBackgroundColor = color;
        }

        private void PickStickyFontColor_Click(object sender, RoutedEventArgs e)
        {
            if (_viewModel == null) return;
            string color = PickColor(_viewModel.StickyNoteFontColor);
            if (color != null) _viewModel.StickyNoteFontColor = color;
        }

        private string PickColor(string initialHex)
        {
            try
            {
                var dialog = new WpfColorPickerDialog(initialHex) { Owner = this };
                if (dialog.ShowDialog() == true)
                {
                    var color = dialog.SelectedColor;
                    return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not open color dialog: " + ex.Message);
            }
            return null;
        }
        
        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
                e.Handled = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Link failed: " + ex.Message);
            }
        }

        private void CleanupCacheBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Update button state
                CleanupCacheBtn.IsEnabled = false;
                CleanupCacheBtn.Content = "Cleaning...";

                // Perform cleanup
                var cacheService = new AssetCacheService();
                var report = cacheService.CleanupStaleCache();

                // Update cache size display
                UpdateCacheSizeDisplay();

                // Show result message
                string message = $"Cache cleanup completed:\n";
                if (report.StaleCacheFoldersDeleted > 0)
                    message += $"- {report.StaleCacheFoldersDeleted} stale cache folders removed\n";
                if (report.LegacyCacheFoldersDeleted > 0)
                    message += $"- {report.LegacyCacheFoldersDeleted} legacy cache folders removed\n";
                if (report.TotalDeleted == 0)
                    message += "No cache folders to clean up.";

                CacheCleanupResult.Text = message;
                CacheCleanupResult.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {
                CacheCleanupResult.Text = $"Error during cleanup: {ex.Message}";
                CacheCleanupResult.Visibility = Visibility.Visible;
                CacheCleanupResult.Foreground = (Brush)new BrushConverter().ConvertFrom("#E57373");
            }
            finally
            {
                CleanupCacheBtn.IsEnabled = true;
                CleanupCacheBtn.Content = "Cache bereinigen";
            }
        }

        private void UpdateCacheSizeDisplay()
        {
            try
            {
                var cacheService = new AssetCacheService();
                string size = cacheService.GetCacheSizeFormatted();
                CacheSizeText.Text = $"Cache size: {size}";
            }
            catch (Exception ex)
            {
                CacheSizeText.Text = $"Cache size: Error - {ex.Message}";
            }
        }

        private void LoadSharedFolders()
        {
            try
            {
                var cacheService = new AssetCacheService();
                LstSharedFoldersSettings.ItemsSource = cacheService.GetSharedFolders();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadSharedFolders error: {ex.Message}");
            }
        }

        private void BtnAddSharedFolderSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
                {
                    dialog.Description = "Select a shared folder (OneDrive, network share) where local caching should be used";
                    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        var cacheService = new AssetCacheService();
                        cacheService.AddSharedFolder(dialog.SelectedPath);
                        LoadSharedFolders();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not add shared folder: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnRemoveSharedFolderSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LstSharedFoldersSettings.SelectedItem is string path)
                {
                    var cacheService = new AssetCacheService();
                    cacheService.RemoveSharedFolder(path);
                    LoadSharedFolders();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not remove shared folder: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private Point _dragStartPoint;

        private void PoolListBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _dragStartPoint = e.GetPosition(null);
        }

        private void PoolListBox_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                Point mousePos = e.GetPosition(null);
                Vector diff = _dragStartPoint - mousePos;

                if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    if (sender is ListBox listBox && listBox.SelectedItem is SidebarFeature feature)
                    {
                        DragDrop.DoDragDrop(listBox, feature, DragDropEffects.Move);
                    }
                }
            }
        }

        private void SectionListBox_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(SidebarFeature)))
            {
                var feature = e.Data.GetData(typeof(SidebarFeature)) as SidebarFeature;
                if (feature != null && sender is ListBox targetListBox)
                {
                    var section = targetListBox.DataContext as SidebarSection;
                    if (section != null && _viewModel.SelectedPoolFeature != null)
                    {
                        _viewModel.AddFeatureToSectionCommand.Execute(section);
                    }
                }
            }
        }
    }
}
