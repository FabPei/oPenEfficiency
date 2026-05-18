using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using oPenEfficiency.Features;

namespace oPenEfficiency.UI
{
    public partial class ThermometerAdorner : Window
    {
        private PowerPointManager _manager;
        private DispatcherTimer _debounceTimer;
        private bool _suppressUpdate = false;

        private float _percentage = 50f;
        private int _colorBgr = 0x3232E8; // Default Red #E83232 in BGR

        // Parameterless constructor for XAML compilation
        public ThermometerAdorner() : this(null, null) { }

        // Palette definition (BGR format)
        private static readonly int[] PaletteColorsBgr = new[]
        {
            0x3232E8, // Red #E83232
            0xD67A2B, // Blue #2B7AD6
            0x3EAF4C, // Green #4CAF3E
            0x008CFF, // Orange #FF8C00
            0x808080, // Grey #808080
            0x404040, // Dark #404040
            0x07C1FF, // Gold #FFC107
            0xD65F91  // Purple #915FD6
        };

        private static readonly string[] PaletteColorsHex = new[]
        {
            "#E83232", "#2B7AD6", "#4CAF3E", "#FF8C00",
            "#808080", "#404040", "#FFC107", "#915FD6"
        };

        public ThermometerAdorner(PowerPointManager manager, string altText)
        {
            _debounceTimer = new DispatcherTimer();
            _debounceTimer.Interval = TimeSpan.FromMilliseconds(50);
            _debounceTimer.Tick += (s, e) => {
                _debounceTimer.Stop();
                ApplyUpdate();
            };

            InitializeComponent();
            _manager = manager;

            PopulateColorPalette();
            ParseTag(altText);
            SyncUI();
        }

        private void ParseTag(string altText)
        {
            if (string.IsNullOrEmpty(altText)) return;
            var parts = altText.Split('|');
            if (parts.Length >= 2)
            {
                if (float.TryParse(parts[1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out float p))
                    _percentage = p * 100f;
                
                if (parts.Length >= 3)
                    int.TryParse(parts[2], out _colorBgr);
            }
        }

        private void PopulateColorPalette()
        {
            for (int i = 0; i < PaletteColorsHex.Length; i++)
            {
                var btn = new Button
                {
                    Width = 20,
                    Height = 20,
                    Margin = new Thickness(2),
                    Cursor = System.Windows.Input.Cursors.Hand,
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(PaletteColorsHex[i]))
                };

                // Use a simple circular template
                var tmpl = new ControlTemplate(typeof(Button));
                var ellipse = new FrameworkElementFactory(typeof(System.Windows.Shapes.Ellipse));
                ellipse.SetBinding(System.Windows.Shapes.Ellipse.FillProperty, new System.Windows.Data.Binding("Background") { RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.TemplatedParent) });
                ellipse.SetValue(System.Windows.Shapes.Ellipse.StrokeProperty, new SolidColorBrush(Color.FromRgb(224, 224, 224)));
                ellipse.SetValue(System.Windows.Shapes.Ellipse.StrokeThicknessProperty, 1.0);
                tmpl.VisualTree = ellipse;
                btn.Template = tmpl;

                int capturedBgr = PaletteColorsBgr[i];
                btn.Click += (s, e) =>
                {
                    _colorBgr = capturedBgr;
                    UpdateColorWell();
                    ColorPopup.IsOpen = false;
                    ApplyUpdate();
                };

                ColorPalette.Children.Add(btn);
            }

            // Add "Custom Color" button
            var customBtn = new Button
            {
                Width = 20, Height = 20, Margin = new Thickness(2),
                Cursor = System.Windows.Input.Cursors.Hand,
                Background = Brushes.Transparent,
                Content = new System.Windows.Controls.TextBlock { Text = "+", VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center, FontSize = 14, FontWeight = FontWeights.Bold, Foreground = Brushes.Gray }
            };
            var ctmpl = new ControlTemplate(typeof(Button));
            var cellipse = new FrameworkElementFactory(typeof(System.Windows.Shapes.Ellipse));
            cellipse.SetValue(System.Windows.Shapes.Ellipse.FillProperty, Brushes.White);
            cellipse.SetValue(System.Windows.Shapes.Ellipse.StrokeProperty, new SolidColorBrush(Color.FromRgb(224, 224, 224)));
            cellipse.SetValue(System.Windows.Shapes.Ellipse.StrokeThicknessProperty, 1.0);
            ctmpl.VisualTree = cellipse;
            var cp = new FrameworkElementFactory(typeof(ContentPresenter));
            cp.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            cp.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            cellipse.AppendChild(cp);
            customBtn.Template = ctmpl;

            customBtn.Click += (s, e) =>
            {
                var dialog = new System.Windows.Forms.ColorDialog();
                // Convert current BGR to System.Drawing.Color
                byte cr = (byte)(_colorBgr & 0xFF);
                byte cg = (byte)((_colorBgr >> 8) & 0xFF);
                byte cb = (byte)((_colorBgr >> 16) & 0xFF);
                dialog.Color = System.Drawing.Color.FromArgb(cr, cg, cb);

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var c = dialog.Color;
                    _colorBgr = c.R | (c.G << 8) | (c.B << 16);
                    UpdateColorWell();
                    ColorPopup.IsOpen = false;
                    ApplyUpdate();
                }
            };
            ColorPalette.Children.Add(customBtn);
        }

        private void SyncUI()
        {
            _suppressUpdate = true;
            if (SldPercentage != null) SldPercentage.Value = _percentage;
            if (TxtValue != null) TxtValue.Text = ((int)_percentage).ToString();
            UpdateColorWell();
            _suppressUpdate = false;
        }

        private void UpdateColorWell()
        {
            if (BtnColorWell == null) return;
            byte r = (byte)(_colorBgr & 0xFF);
            byte g = (byte)((_colorBgr >> 8) & 0xFF);
            byte b = (byte)((_colorBgr >> 16) & 0xFF);
            BtnColorWell.Background = new SolidColorBrush(Color.FromRgb(r, g, b));
        }

        private void SldPercentage_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_suppressUpdate || TxtValue == null || SldPercentage == null) return;
            _percentage = (float)SldPercentage.Value;
            TxtValue.Text = ((int)_percentage).ToString();
            TriggerUpdate();
        }

        private void TxtValue_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_suppressUpdate || TxtValue == null || SldPercentage == null) return;
            if (float.TryParse(TxtValue.Text, out float val))
            {
                if (val < 0) val = 0;
                if (val > 100) val = 100;
                _percentage = val;
                SldPercentage.Value = _percentage;
                TriggerUpdate();
            }
        }

        private void BtnColorWell_Click(object sender, RoutedEventArgs e)
        {
            ColorPopup.IsOpen = !ColorPopup.IsOpen;
        }

        private void TriggerUpdate()
        {
            if (_debounceTimer == null) return;
            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        private void ApplyUpdate()
        {
            try
            {
                var selection = _manager?.GetSelectedShapes();
                if (selection != null && selection.Count == 1)
                {
                    ThermometerFeature.UpdateThermometer(selection[1], _percentage / 100f, _colorBgr);
                }
            }
            catch { }
        }

        private void Window_Deactivated(object sender, EventArgs e)
        {
            if (ColorPopup != null && ColorPopup.IsOpen) return;
            // Closing is handled by ThisAddIn on selection change.
        }
        
        public void CloseAdorner()
        {
            this.Close();
        }
    }
}
