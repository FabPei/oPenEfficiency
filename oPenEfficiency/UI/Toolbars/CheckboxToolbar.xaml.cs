using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using oPenEfficiency.Features;
using oPenEfficiency.Services;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public partial class CheckboxToolbar : Window
    {
        private PowerPointManager _manager;
        private Microsoft.Office.Interop.PowerPoint.Shape _targetShape;
        private DispatcherTimer _debounceTimer;
        private bool _suppressUpdate = false;

        // Current state
        private CheckboxFeature.CheckboxState _currentState = CheckboxFeature.CheckboxState.Tick;
        private bool _hasFrame = false;
        private float _fontSize = 32f;
        private int _fillColorRgb = 0x000000;

        // Parameterless constructor for XAML compilation
        public CheckboxToolbar() : this(null) { }

        // Theme colors (BGR format for PowerPoint)
        private static readonly int[] PaletteColorsBgr = new[]
        {
            0x000000,  // Black   #000000
            0xFFFFFF,  // White   #FFFFFF
            0xE53935,  // Red     #E53935
            0x0078D4,  // Blue    #D47800 (Theme Blue BGR)
            0x4CAF50,  // Green   #50AF4C (Theme Green BGR)
            0x809800,  // Orange  #009880 (Theme Orange BGR)
            0x07C1FF,  // Gold    #FFC107 
            0xB027B0,  // Purple  #B0279C
        };

        // Matching WPF colors for the palette swatches
        private static readonly string[] PaletteColorsHex = new[]
        {
            "#000000", "#FFFFFF", "#E53935", "#0078D4", "#4CAF50", "#FF9800", "#FFC107", "#9C27B0"
        };

        public CheckboxToolbar(PowerPointManager manager)
        {
            InitializeCommon(manager);

            if (manager == null) return;

            var selection = _manager.GetSelectedShapes();
            if (selection != null && selection.Count == 1)
            {
                var selected = selection[1];
                if (selected.AlternativeText != null && selected.AlternativeText.StartsWith("oPE_Checkbox"))
                {
                    _targetShape = selected;
                    ParseTag(selected.AlternativeText);
                    SyncUIFromState();
                    return;
                }
            }

            ApplyRating();
        }

        public CheckboxToolbar(PowerPointManager manager, string altText)
        {
            InitializeCommon(manager);
            if (manager == null) return;
            
            var selection = _manager.GetSelectedShapes();
            if (selection != null && selection.Count == 1)
            {
                var selected = selection[1];
                if (selected.AlternativeText != null && selected.AlternativeText.StartsWith("oPE_Checkbox"))
                {
                    _targetShape = selected;
                }
            }

            ParseTag(altText);
            SyncUIFromState();
        }

        private void InitializeCommon(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;

            _debounceTimer = new DispatcherTimer();
            _debounceTimer.Interval = TimeSpan.FromMilliseconds(100);
            _debounceTimer.Tick += DebounceTimer_Tick;

            PopulateColorPalette();

            this.Loaded += (s, e) => PositionWindow();
        }

        private void PositionWindow()
        {
            if (_manager == null) return;
            try
            {
                _manager.GetSelectionScreenTopCenterLocation(out int screenX, out int screenY);

                var source = PresentationSource.FromVisual(this);
                double dpiX = 1.0, dpiY = 1.0;
                if (source?.CompositionTarget != null)
                {
                    dpiX = source.CompositionTarget.TransformToDevice.M11;
                    dpiY = source.CompositionTarget.TransformToDevice.M22;
                }

                this.Left = (screenX / dpiX) - (this.Width / 2);
                this.Top = (screenY / dpiY) - this.Height - 10;

                if (this.Top < 0) this.Top = 10;
                if (this.Left < 0) this.Left = 10;
            }
            catch { }
        }

        private void ParseTag(string altText)
        {
            // Format: oPE_Checkbox|state|hasFrame|fontSize|colorRgb
            var parts = altText.Split('|');
            if (parts.Length >= 5)
            {
                Enum.TryParse(parts[1], out _currentState);
                bool.TryParse(parts[2], out _hasFrame);
                float.TryParse(parts[3], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out _fontSize);
                int.TryParse(parts[4], out _fillColorRgb);
            }
        }

        private void SyncUIFromState()
        {
            _suppressUpdate = true;

            TxtFontSize.Text = _fontSize.ToString("0");
            
            switch (_currentState)
            {
                case CheckboxFeature.CheckboxState.Tick: CmbSymbol.SelectedIndex = 0; break;
                case CheckboxFeature.CheckboxState.Cross: CmbSymbol.SelectedIndex = 1; break;
                case CheckboxFeature.CheckboxState.Empty: CmbSymbol.SelectedIndex = 2; break;
            }

            ChkFrame.IsChecked = _hasFrame;

            UpdateColorWell();

            _suppressUpdate = false;
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

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try { DragMove(); } catch { }
        }

        private void Window_Deactivated(object sender, EventArgs e)
        {
            if (ColorPopup.IsOpen) return;
            this.Close();
        }

        private void DebounceTimer_Tick(object sender, EventArgs e)
        {
            _debounceTimer.Stop();
            ApplyRating();
        }

        private void TxtFontSize_LostFocus(object sender, RoutedEventArgs e)
        {
            UpdateFontSize();
        }

        private void TxtFontSize_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                UpdateFontSize();
                e.Handled = true;
            }
        }

        private void UpdateFontSize()
        {
            if (float.TryParse(TxtFontSize.Text, out float newSize) && newSize >= 8 && newSize <= 200)
            {
                if (Math.Abs(_fontSize - newSize) > 0.1f)
                {
                    _fontSize = newSize;
                    _suppressUpdate = true;
                    ApplyRating();
                    _suppressUpdate = false;
                }
            }
            else
            {
                TxtFontSize.Text = _fontSize.ToString("0");
            }
        }

        private void CmbSymbol_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressUpdate) return;
            
            if (CmbSymbol.SelectedItem is ComboBoxItem item && item.Tag != null)
            {
                if (Enum.TryParse(item.Tag.ToString(), out CheckboxFeature.CheckboxState state))
                {
                    _currentState = state;
                    ApplyRating();
                }
            }
        }

        private void ChkFrame_Changed(object sender, RoutedEventArgs e)
        {
            if (_suppressUpdate) return;
            _hasFrame = ChkFrame.IsChecked == true;
            ApplyRating();
        }

        private void BtnColorWell_Click(object sender, RoutedEventArgs e)
        {
            ColorPopup.IsOpen = !ColorPopup.IsOpen;
        }

        private void ApplyRating()
        {
            try
            {
                if (_manager != null)
                {
                    CheckboxFeature.Execute(_manager, _currentState, _hasFrame, _fontSize, _fillColorRgb);
                    
                    var selection = _manager.GetSelectedShapes();
                    if (selection != null && selection.Count == 1)
                    {
                        var selected = selection[1];
                        if (selected.AlternativeText != null && selected.AlternativeText.StartsWith("oPE_Checkbox"))
                        {
                            _targetShape = selected;
                        }
                    }
                }
            }
            catch { }
        }
    }
}