using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using Bitmap = System.Drawing.Bitmap;
using WinForms = System.Windows.Forms;

namespace oPenEfficiency.Excel.Features.ChartScanner
{
    /// <summary>
    /// Transparent full-screen overlay for selecting a screen region to capture
    /// (think-cell-style "snip"). The user drags a rectangle; on release the covered
    /// screen pixels are copied into <see cref="Captured"/>.
    ///
    /// Built entirely in code (no XAML) because the Excel add-in project has no WPF
    /// build targets. DPI is handled by converting the WPF selection (DIPs) to device
    /// pixels via the window's current transform before copying from the screen.
    /// </summary>
    public sealed class ScreenCaptureOverlay : Window
    {
        private readonly Canvas _canvas;
        private readonly Rectangle _selection;
        private System.Windows.Point _start;
        private bool _dragging;

        /// <summary>The captured region, or null if the user cancelled.</summary>
        public Bitmap Captured { get; private set; }

        public ScreenCaptureOverlay()
        {
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            Background = new SolidColorBrush(Color.FromArgb(80, 0, 0, 0));
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;
            Topmost = true;
            Cursor = Cursors.Cross;

            // Cover the whole virtual screen. VirtualScreen is in device pixels; convert to DIPs.
            var vs = WinForms.SystemInformation.VirtualScreen;
            double scale = GetDpiScale();
            Left = vs.Left / scale;
            Top = vs.Top / scale;
            Width = vs.Width / scale;
            Height = vs.Height / scale;

            _selection = new Rectangle
            {
                Stroke = new SolidColorBrush(Color.FromRgb(0x3B, 0x82, 0xF6)),
                StrokeThickness = 2,
                Fill = new SolidColorBrush(Color.FromArgb(40, 0x3B, 0x82, 0xF6)),
                Visibility = Visibility.Collapsed
            };

            var hint = new TextBlock
            {
                Text = "Drag a region over the chart - Esc to cancel",
                Foreground = Brushes.White,
                FontSize = 14,
                Margin = new Thickness(12)
            };

            _canvas = new Canvas();
            _canvas.Children.Add(hint);
            _canvas.Children.Add(_selection);
            Content = _canvas;

            MouseLeftButtonDown += OnDown;
            MouseMove += OnMove;
            MouseLeftButtonUp += OnUp;
            KeyDown += (s, e) => { if (e.Key == Key.Escape) { DialogResult = false; Close(); } };
        }

        private double GetDpiScale()
        {
            try
            {
                using (var g = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
                    return g.DpiX / 96.0;
            }
            catch { return 1.0; }
        }

        private void OnDown(object sender, MouseButtonEventArgs e)
        {
            _start = e.GetPosition(_canvas);
            _dragging = true;
            Canvas.SetLeft(_selection, _start.X);
            Canvas.SetTop(_selection, _start.Y);
            _selection.Width = 0;
            _selection.Height = 0;
            _selection.Visibility = Visibility.Visible;
            _canvas.CaptureMouse();
        }

        private void OnMove(object sender, MouseEventArgs e)
        {
            if (!_dragging) return;
            var p = e.GetPosition(_canvas);
            double x = Math.Min(p.X, _start.X);
            double y = Math.Min(p.Y, _start.Y);
            Canvas.SetLeft(_selection, x);
            Canvas.SetTop(_selection, y);
            _selection.Width = Math.Abs(p.X - _start.X);
            _selection.Height = Math.Abs(p.Y - _start.Y);
        }

        private void OnUp(object sender, MouseButtonEventArgs e)
        {
            if (!_dragging) return;
            _dragging = false;
            _canvas.ReleaseMouseCapture();

            double x = Canvas.GetLeft(_selection);
            double y = Canvas.GetTop(_selection);
            double w = _selection.Width;
            double h = _selection.Height;

            // Hide the overlay before grabbing the screen so it is not captured itself.
            _selection.Visibility = Visibility.Collapsed;
            Visibility = Visibility.Hidden;

            if (w < 4 || h < 4)
            {
                DialogResult = false;
                Close();
                return;
            }

            try
            {
                double scale = GetDpiScale();
                // Selection is relative to the window (which starts at the virtual-screen origin).
                int devX = (int)Math.Round((Left + x) * scale);
                int devY = (int)Math.Round((Top + y) * scale);
                int devW = (int)Math.Round(w * scale);
                int devH = (int)Math.Round(h * scale);

                var bmp = new Bitmap(devW, devH);
                using (var g = System.Drawing.Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(devX, devY, 0, 0, new System.Drawing.Size(devW, devH));
                }
                Captured = bmp;
                DialogResult = true;
            }
            catch
            {
                DialogResult = false;
            }
            finally
            {
                Close();
            }
        }
    }
}
