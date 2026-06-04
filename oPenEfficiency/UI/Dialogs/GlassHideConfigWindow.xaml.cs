using System;
using System.Windows;
using System.Windows.Input;
using oPenEfficiency.Features;
using oPenEfficiency.Services;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class GlassHideConfigWindow : Window
    {
        private PowerPointManager _manager;

        public GlassHideConfigWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            
            // Load current defaults
            SliderTransparency.Value = GlassHideFeature.DefaultTransparency * 100;
            TextTransparencyVal.Text = $"{(int)SliderTransparency.Value}%";
            
            byte r = (byte)(GlassHideFeature.DefaultColorRgb & 0xFF);
            byte g = (byte)((GlassHideFeature.DefaultColorRgb >> 8) & 0xFF);
            byte b = (byte)((GlassHideFeature.DefaultColorRgb >> 16) & 0xFF);
            TextHexColor.Text = $"#{r:X2}{g:X2}{b:X2}";
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void SliderTransparency_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (TextTransparencyVal != null)
                TextTransparencyVal.Text = $"{(int)e.NewValue}%";
        }

        private void PickColor_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.ColorDialog())
            {
                try
                {
                    var hex = TextHexColor.Text.Trim().Replace("#", "");
                    if (hex.Length == 6)
                    {
                        byte r = Convert.ToByte(hex.Substring(0, 2), 16);
                        byte g = Convert.ToByte(hex.Substring(2, 2), 16);
                        byte b = Convert.ToByte(hex.Substring(4, 2), 16);
                        dialog.Color = System.Drawing.Color.FromArgb(r, g, b);
                    }
                }
                catch { }

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TextHexColor.Text = $"#{dialog.Color.R:X2}{dialog.Color.G:X2}{dialog.Color.B:X2}";
                }
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            GlassHideFeature.DefaultTransparency = (float)(SliderTransparency.Value / 100.0);
            
            try
            {
                var hex = TextHexColor.Text.Trim().Replace("#", "");
                if (hex.Length == 6)
                {
                    byte r = Convert.ToByte(hex.Substring(0, 2), 16);
                    byte g = Convert.ToByte(hex.Substring(2, 2), 16);
                    byte b = Convert.ToByte(hex.Substring(4, 2), 16);
                    GlassHideFeature.DefaultColorRgb = (b << 16) | (g << 8) | r;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Invalid color format. Please use #RRGGBB.");
                return;
            }

            this.Close();
            
            // Auto-execute with new settings
            GlassHideFeature.Execute(_manager, GlassHideFeature.GlassMode.Single);
        }
    }
}