using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using oPenEfficiency.Features;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public partial class RatingToolbar : Window
    {
        private PowerPointManager _manager;
        private Microsoft.Office.Interop.PowerPoint.Shape _targetShape;
        private DispatcherTimer _debounceTimer;
        private bool _suppressSliderUpdate = false;

        // Current state
        private int _starCount = 5;
        private double _rating = 3.0;
        private int _iconMode = 0;   // 0=Star, 1=Circle, 2=Diamond
        private bool _hasFrame = false;
        private int _fillColorRgb = 0x07C1FF; // Gold #FFC107 in BGR for PPT

        // Parameterless constructor for XAML compilation
        public RatingToolbar() : this(null) { }

        // Standard yellow + Theme colors (BGR format for PowerPoint)
        private static readonly int[] PaletteColorsBgr = new[]
        {
            0x07C1FF,  // Gold    #FFC107 (Standard Yellow)
            0xE53935,  // Red     #E53935 (Theme Red)
            0xD40078,  // Blue    #7800D4 (Theme Blue)
            0x4CAF50,  // Green   #4CAF50 (Theme Green)
            0x809800,  // Orange  #FF9800 (Theme Orange)
            0xB027B0,  // Purple  #9C27B0 (Theme Purple)
        };

        // Matching WPF colors for the palette swatches
        private static readonly string[] PaletteColorsHex = new[]
        {
            "#FFC107", "#E53935", "#7800D4", "#4CAF50",
            "#FF9800", "#9C27B0"
        };

        public RatingToolbar(PowerPointManager manager)
        {
            InitializeCommon(manager);

            // Skip initialization if manager is null (designer/null constructor)
            if (manager == null) return;

            // Check if we should initialize from existing selection
            var selection = _manager.GetSelectedShapes();
            if (selection != null && selection.Count == 1)
            {
                var selected = selection[1];
                if (selected.AlternativeText.StartsWith(StarRatingFeature.StarRatingTag))
                {
                    ParseTag(selected.AlternativeText);
                    SyncUIFromState();
                    return;
                }
                else if (selected.Child == Office.MsoTriState.msoTrue)
                {
                    try
                    {
                        var parent = selected.ParentGroup;
                        if (parent != null && parent.AlternativeText.StartsWith(StarRatingFeature.StarRatingTag))
                        {
                            ParseTag(parent.AlternativeText);
                            SyncUIFromState();
                            return;
                        }
                    }
                    catch { }
                }
            }

            // For new star ratings, create the initial rating immediately
            ApplyRating();
        }

        public RatingToolbar(PowerPointManager manager, string altText)
        {
            InitializeCommon(manager);
            if (manager == null) return;
            ParseTag(altText);
            SyncUIFromState();
        }

        private void InitializeCommon(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;

            // Setup debounce timer (100ms)
            _debounceTimer = new DispatcherTimer();
            _debounceTimer.Interval = TimeSpan.FromMilliseconds(100);
            _debounceTimer.Tick += DebounceTimer_Tick;

            // Populate color palette
            PopulateColorPalette();

            this.Loaded += (s, e) => PositionWindow();
        }

        private void PositionWindow()
        {
            if (_manager == null) return;
            try
            {
                _manager.GetSelectionScreenTopCenterLocation(out int screenX, out int screenY);

                // Convert device pixels to WPF DIPs
                var source = PresentationSource.FromVisual(this);
                double dpiX = 1.0, dpiY = 1.0;
                if (source?.CompositionTarget != null)
                {
                    dpiX = source.CompositionTarget.TransformToDevice.M11;
                    dpiY = source.CompositionTarget.TransformToDevice.M22;
                }

                this.Left = (screenX / dpiX) - (this.Width / 2);
                this.Top = (screenY / dpiY) - this.Height - 10; // 10px above the object

                // Ensure it's not off-screen (at least not at the top)
                if (this.Top < 0) this.Top = 10;
                if (this.Left < 0) this.Left = 10;
            }
            catch { }
        }

        private void ParseTag(string altText)
        {
            // Format: oPE_StarRating|count|rating|iconMode|hasFrame|colorBgr
            var parts = altText.Split('|');
            if (parts.Length >= 6)
            {
                int.TryParse(parts[1], out _starCount);
                double.TryParse(parts[2], System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out _rating);
                int.TryParse(parts[3], out _iconMode);
                _hasFrame = parts[4] == "1";
                int.TryParse(parts[5], out _fillColorRgb);
            }
        }

        private void SyncUIFromState()
        {
            _suppressSliderUpdate = true;

            TxtStarCount.Text = _starCount.ToString();
            SliderRating.Maximum = _starCount;
            SliderRating.Value = _rating;
            TxtRatingValue.Text = _rating.ToString("F1");
            CmbIcon.SelectedIndex = _iconMode;
            ChkFrame.IsChecked = _hasFrame;

            // Find matching palette color for the well
            UpdateColorWell();

            _suppressSliderUpdate = false;
        }

        private void UpdateColorWell()
        {
            for (int i = 0; i < PaletteColorsBgr.Length; i++)
            {
                if (PaletteColorsBgr[i] == _fillColorRgb)
                {
                    var cc = (Color)ColorConverter.ConvertFromString(PaletteColorsHex[i]);
                    BtnColorWell.Background = new SolidColorBrush(cc);
                    return;
                }
            }
            // Fallback: convert BGR to WPF color
            int r = _fillColorRgb & 0xFF;
            int g = (_fillColorRgb >> 8) & 0xFF;
            int b = (_fillColorRgb >> 16) & 0xFF;
            BtnColorWell.Background = new SolidColorBrush(Color.FromRgb((byte)r, (byte)g, (byte)b));
        }

        private void PopulateColorPalette()
        {
            for (int i = 0; i < PaletteColorsHex.Length; i++)
            {
                var btn = new System.Windows.Controls.Button();
                btn.Width = 20;
                btn.Height = 20;
                btn.Margin = new Thickness(2);
                btn.Cursor = Cursors.Hand;

                var cc = (Color)ColorConverter.ConvertFromString(PaletteColorsHex[i]);
                btn.Background = new SolidColorBrush(cc);

                // Style like the color swatch
                var tmpl = new ControlTemplate(typeof(System.Windows.Controls.Button));
                var ellipse = new FrameworkElementFactory(typeof(System.Windows.Shapes.Ellipse));
                ellipse.SetBinding(System.Windows.Shapes.Ellipse.FillProperty,
                    new System.Windows.Data.Binding("Background") { RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.TemplatedParent) });
                ellipse.SetValue(System.Windows.Shapes.Ellipse.StrokeProperty, new SolidColorBrush(Color.FromRgb(224, 224, 224)));
                ellipse.SetValue(System.Windows.Shapes.Ellipse.StrokeThicknessProperty, 1.0);
                tmpl.VisualTree = ellipse;
                btn.Template = tmpl;

                int capturedBgr = PaletteColorsBgr[i];
                string capturedHex = PaletteColorsHex[i];
                btn.Click += (s, e) =>
                {
                    _fillColorRgb = capturedBgr;
                    var c = (Color)ColorConverter.ConvertFromString(capturedHex);
                    BtnColorWell.Background = new SolidColorBrush(c);
                    ColorPopup.IsOpen = false;
                    ApplyRating();
                };

                ColorPalette.Children.Add(btn);
            }
        }

        // =============== Event Handlers ===============

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try { DragMove(); } catch { }
        }

        private void Window_Deactivated(object sender, EventArgs e)
        {
            // Don't close if color popup is open
            if (ColorPopup.IsOpen) return;
            this.Close();
        }

        private void SliderRating_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_suppressSliderUpdate || _debounceTimer == null) return;
            _rating = Math.Round(SliderRating.Value, 1);
            if (TxtRatingValue != null)
                TxtRatingValue.Text = _rating.ToString("F1");

            // Debounced apply
            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        private void DebounceTimer_Tick(object sender, EventArgs e)
        {
            _debounceTimer.Stop();
            ApplyRating();
        }

        private void TxtStarCount_LostFocus(object sender, RoutedEventArgs e)
        {
            UpdateStarCount();
        }

        private void TxtStarCount_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                UpdateStarCount();
                e.Handled = true;
            }
        }

        private void UpdateStarCount()
        {
            if (int.TryParse(TxtStarCount.Text, out int newCount) && newCount >= 1 && newCount <= 10)
            {
                if (_starCount != newCount)
                {
                    _starCount = newCount;
                    _suppressSliderUpdate = true;
                    SliderRating.Maximum = _starCount;
                    if (_rating > _starCount) _rating = _starCount;
                    SliderRating.Value = _rating;
                    _suppressSliderUpdate = false;
                    ApplyRating();
                }
            }
            else
            {
                TxtStarCount.Text = _starCount.ToString();
            }
        }

        private void CmbIcon_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressSliderUpdate) return;
            _iconMode = CmbIcon.SelectedIndex;
            ApplyRating();
        }

        private void ChkFrame_Changed(object sender, RoutedEventArgs e)
        {
            if (_suppressSliderUpdate) return;
            _hasFrame = ChkFrame.IsChecked == true;
            ApplyRating();
        }

        private void BtnColorWell_Click(object sender, RoutedEventArgs e)
        {
            ColorPopup.IsOpen = !ColorPopup.IsOpen;
        }

        // =============== Apply to PowerPoint ===============

        private void ApplyRating()
        {
            try
            {
                if (_manager != null)
                {
                    StarRatingFeature.Execute(_manager, _starCount, _rating, _iconMode, _hasFrame, _fillColorRgb, _targetShape);
                    
                    var selection = _manager.GetSelectedShapes();
                    if (selection != null && selection.Count == 1)
                    {
                        var selected = selection[1];
                        if (selected.AlternativeText != null && selected.AlternativeText.StartsWith(StarRatingFeature.StarRatingTag))
                        {
                            _targetShape = selected;
                        }
                        else if (selected.Child == Office.MsoTriState.msoTrue)
                        {
                            try
                            {
                                var parent = selected.ParentGroup;
                                if (parent != null && parent.AlternativeText != null && parent.AlternativeText.StartsWith(StarRatingFeature.StarRatingTag))
                                {
                                    _targetShape = parent;
                                }
                            }
                            catch { }
                        }
                    }
                }
            }
            catch { }
        }
    }
}
