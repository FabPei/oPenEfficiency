using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using Brushes = System.Windows.Media.Brushes;
using Color = System.Windows.Media.Color;
using Point = System.Windows.Point;
using WinForms = System.Windows.Forms;

namespace oPenEfficiency.Excel.Features.ChartScanner
{
    /// <summary>
    /// Interactive chart digitizer (code-only WPF). The user loads a chart image
    /// (screen capture, clipboard or file), calibrates the value axis with two known
    /// reference points, then clicks the data points. Values are computed from the
    /// calibration and can be exported to Excel by the caller via <see cref="Result"/>.
    /// </summary>
    public sealed class ChartScannerWindow : Window
    {
        private enum Mode { None, Calibrate, AddPoints, SampleColor }

        private readonly IChartOcr _ocr;

        private Bitmap _image;
        private readonly Canvas _canvas = new Canvas();
        private readonly System.Windows.Controls.Image _imageControl = new System.Windows.Controls.Image();
        private readonly ComboBox _typeCombo = new ComboBox();
        private readonly TextBlock _status = new TextBlock();
        private readonly StackPanel _pointsPanel = new StackPanel();
        private readonly Button _calibrateButton = new Button();
        private readonly Button _addPointsButton = new Button();
        private readonly Button _autoDetectButton = new Button();

        private Mode _mode = Mode.None;
        private AxisCalibration _calibration;
        private IReadOnlyList<OcrNumber> _ocrNumbers = System.Array.Empty<OcrNumber>();
        private readonly List<double> _calPixels = new List<double>();
        private readonly List<double> _calValues = new List<double>();
        private readonly List<UIElement> _calOverlays = new List<UIElement>();
        private readonly List<Row> _rows = new List<Row>();

        public ChartScanResult Result { get; } = new ChartScanResult();

        private sealed class Row
        {
            public ChartDataPoint Point;
            public UIElement Marker;
            public UIElement Label;
            public FrameworkElement RowUi;
        }

        public ChartScannerWindow(IChartOcr ocr = null)
        {
            _ocr = ocr ?? new NullChartOcr();

            Title = "Chart Scanner - Digitize Chart";
            Width = 1120;
            Height = 720;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;

            Content = BuildLayout();
            UpdateChartKind();
            SetStatus("Load an image (screen capture / clipboard / file), then calibrate the axis.");
        }

        private ChartKind CurrentKind =>
            _typeCombo.SelectedIndex == 1 ? ChartKind.Bars :
            _typeCombo.SelectedIndex == 2 ? ChartKind.Lines : ChartKind.Columns;

        private AxisOrientation CurrentOrientation =>
            CurrentKind == ChartKind.Bars ? AxisOrientation.Horizontal : AxisOrientation.Vertical;

        // ---- Layout ----

        private UIElement BuildLayout()
        {
            var root = new DockPanel();

            // Top toolbar
            var toolbar = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(8) };
            toolbar.Children.Add(MakeButton("Capture Area", OnCapture));
            toolbar.Children.Add(MakeButton("From Clipboard", OnPaste));
            toolbar.Children.Add(MakeButton("Open Image...", OnOpenFile));
            toolbar.Children.Add(new Separator { Margin = new Thickness(8, 0, 8, 0) });
            toolbar.Children.Add(new TextBlock { Text = "Type:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 6, 0) });
            _typeCombo.Items.Add("Columns (vertical)");
            _typeCombo.Items.Add("Bars (horizontal)");
            _typeCombo.Items.Add("Lines");
            _typeCombo.SelectedIndex = 0;
            _typeCombo.Width = 160;
            _typeCombo.SelectionChanged += (s, e) => UpdateChartKind();
            toolbar.Children.Add(_typeCombo);
            DockPanel.SetDock(toolbar, Dock.Top);
            root.Children.Add(toolbar);

            // Right side panel
            var side = new StackPanel { Width = 300, Margin = new Thickness(8) };
            side.Children.Add(Bold("Instructions"));
            _status.TextWrapping = TextWrapping.Wrap;
            _status.Margin = new Thickness(0, 0, 0, 10);
            side.Children.Add(_status);

            _calibrateButton.Content = "1) Calibrate Value Axis";
            _calibrateButton.Margin = new Thickness(0, 0, 0, 6);
            _calibrateButton.Click += (s, e) => StartCalibration();
            side.Children.Add(_calibrateButton);

            _addPointsButton.Content = "2) Capture Data Points";
            _addPointsButton.Margin = new Thickness(0, 0, 0, 6);
            _addPointsButton.Click += (s, e) => ToggleAddPoints();
            side.Children.Add(_addPointsButton);

            _autoDetectButton.Content = "Auto-Detect Bars";
            _autoDetectButton.Margin = new Thickness(0, 0, 0, 6);
            _autoDetectButton.Click += (s, e) => StartAutoDetect();
            side.Children.Add(_autoDetectButton);

            side.Children.Add(MakeButton("Clear All Points", (s, e) => ClearPoints()));

            side.Children.Add(Bold("Data Points"));
            var scroller = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Height = 320,
                Content = _pointsPanel
            };
            side.Children.Add(scroller);

            var okButton = MakeButton("Send to Excel", (s, e) => Accept());
            okButton.Margin = new Thickness(0, 10, 0, 6);
            side.Children.Add(okButton);
            side.Children.Add(MakeButton("Cancel", (s, e) => { DialogResult = false; Close(); }));

            DockPanel.SetDock(side, Dock.Right);
            root.Children.Add(side);

            // Center image canvas
            _imageControl.Stretch = Stretch.None;
            Canvas.SetLeft(_imageControl, 0);
            Canvas.SetTop(_imageControl, 0);
            _canvas.Children.Add(_imageControl);
            _canvas.Background = Brushes.Gainsboro;
            _canvas.MouseLeftButtonDown += OnCanvasClick;

            var center = new ScrollViewer
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = _canvas
            };
            root.Children.Add(center);
            return root;
        }

        private static TextBlock Bold(string text) =>
            new TextBlock { Text = text, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 6, 0, 4) };

        private Button MakeButton(string text, RoutedEventHandler onClick)
        {
            var b = new Button { Content = text, Margin = new Thickness(0, 0, 6, 0), Padding = new Thickness(8, 3, 8, 3) };
            b.Click += onClick;
            return b;
        }

        private void SetStatus(string text) => _status.Text = text;

        // ---- Image sources ----

        private void OnCapture(object sender, RoutedEventArgs e)
        {
            Hide();
            try
            {
                var overlay = new ScreenCaptureOverlay();
                if (overlay.ShowDialog() == true && overlay.Captured != null)
                    SetImage(overlay.Captured);
            }
            finally { Show(); Activate(); }
        }

        private void OnPaste(object sender, RoutedEventArgs e)
        {
            try
            {
                if (WinForms.Clipboard.ContainsImage())
                {
                    using (var img = WinForms.Clipboard.GetImage())
                        SetImage(ImageInterop.Clone(img));
                }
                else SetStatus("No image found in the clipboard.");
            }
            catch (Exception ex) { SetStatus("Could not read the clipboard: " + ex.Message); }
        }

        private void OnOpenFile(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Images|*.png;*.jpg;*.jpeg;*.bmp;*.gif|All Files|*.*"
            };
            if (dlg.ShowDialog() == true)
            {
                try
                {
                    using (var tmp = new Bitmap(dlg.FileName))
                        SetImage(ImageInterop.Clone(tmp));
                }
                catch (Exception ex) { SetStatus("Could not load the image: " + ex.Message); }
            }
        }

        private void SetImage(Bitmap bitmap)
        {
            if (bitmap == null) return;
            ClearPoints();
            ResetCalibration();

            _image?.Dispose();
            _image = bitmap;
            _imageControl.Source = ImageInterop.ToBitmapSource(bitmap);
            _imageControl.Width = bitmap.Width;
            _imageControl.Height = bitmap.Height;
            _canvas.Width = bitmap.Width;
            _canvas.Height = bitmap.Height;

            try { _ocrNumbers = _ocr.ReadNumbers(bitmap, System.Drawing.Rectangle.Empty); }
            catch { _ocrNumbers = System.Array.Empty<OcrNumber>(); }

            string ocrNote = _ocr.IsAvailable
                ? $" OCR enabled ({_ocrNumbers.Count} numbers detected) - suggested reference values are available."
                : " (OCR unavailable - please enter the values manually.)";
            SetStatus("Image loaded. Click 'Calibrate Value Axis' and set two reference points." + ocrNote);
        }

        // ---- Chart kind ----

        private void UpdateChartKind()
        {
            Result.Kind = CurrentKind;
            // Orientation may change; existing calibration/points become invalid.
            ResetCalibration();
            ClearPoints();
            _autoDetectButton.Content = CurrentKind == ChartKind.Lines
                ? "Auto-Detect Line"
                : "Auto-Detect Bars";
            SetStatus(CurrentKind == ChartKind.Bars
                ? "Bars: calibrate the horizontal value axis with two points."
                : "Calibrate the vertical value axis with two points.");
        }

        // ---- Calibration ----

        private void StartCalibration()
        {
            if (_image == null) { SetStatus("Please load an image first."); return; }
            ResetCalibration();
            _mode = Mode.Calibrate;
            HighlightModeButtons();
            SetStatus("Calibration: click the 1st reference point on the value axis (for example, a gridline with a known value).");
        }

        private void ResetCalibration()
        {
            _calibration = null;
            _calPixels.Clear();
            _calValues.Clear();
            foreach (var o in _calOverlays) _canvas.Children.Remove(o);
            _calOverlays.Clear();
        }

        private void HandleCalibrationClick(Point p)
        {
            double axisPixel = CurrentOrientation == AxisOrientation.Vertical ? p.Y : p.X;
            int index = _calPixels.Count + 1;
            string suggestion = NearestOcrValue(p);
            string entry = PromptDialog.Show(this, "Reference Value",
                $"Value at reference point {index} (axis label):", suggestion);
            double? value = ParseNumber(entry);
            if (value == null) { SetStatus("Cancelled - calibration is incomplete."); _mode = Mode.None; HighlightModeButtons(); return; }

            _calPixels.Add(axisPixel);
            _calValues.Add(value.Value);
            DrawCalibrationMarker(p, value.Value);

            if (_calPixels.Count == 1)
            {
                SetStatus("Click the 2nd reference point (another known value on the value axis).");
            }
            else
            {
                var cal = new AxisCalibration(CurrentOrientation,
                    _calPixels[0], _calValues[0], _calPixels[1], _calValues[1]);
                if (!cal.IsValid)
                {
                    SetStatus("The two points are too close together. Please calibrate again.");
                    ResetCalibration();
                    return;
                }
                _calibration = cal;
                _mode = Mode.None;
                HighlightModeButtons();
                SetStatus("Axis calibrated. Click 'Capture Data Points' and mark the bars or data points.");
            }
        }

        private void DrawCalibrationMarker(Point p, double value)
        {
            var dot = new Ellipse
            {
                Width = 10, Height = 10,
                Fill = new SolidColorBrush(Color.FromRgb(0xF5, 0x9E, 0x0B)),
                Stroke = Brushes.White, StrokeThickness = 1
            };
            Canvas.SetLeft(dot, p.X - 5);
            Canvas.SetTop(dot, p.Y - 5);
            _canvas.Children.Add(dot);
            _calOverlays.Add(dot);

            // Guide line across the whole image at the reference position.
            var line = new Line
            {
                Stroke = new SolidColorBrush(Color.FromArgb(120, 0xF5, 0x9E, 0x0B)),
                StrokeThickness = 1
            };
            if (CurrentOrientation == AxisOrientation.Vertical)
            { line.X1 = 0; line.X2 = _canvas.Width; line.Y1 = p.Y; line.Y2 = p.Y; }
            else
            { line.Y1 = 0; line.Y2 = _canvas.Height; line.X1 = p.X; line.X2 = p.X; }
            _canvas.Children.Add(line);
            _calOverlays.Add(line);

            var tag = new TextBlock { Text = value.ToString("0.####", CultureInfo.CurrentCulture), Foreground = Brushes.Black, FontSize = 11 };
            Canvas.SetLeft(tag, p.X + 8);
            Canvas.SetTop(tag, p.Y - 8);
            _canvas.Children.Add(tag);
            _calOverlays.Add(tag);
        }

        // ---- Data points ----

        private void ToggleAddPoints()
        {
            if (_image == null) { SetStatus("Please load an image first."); return; }
            if (_calibration == null || !_calibration.IsValid) { SetStatus("Please calibrate the value axis first."); return; }
            _mode = _mode == Mode.AddPoints ? Mode.None : Mode.AddPoints;
            HighlightModeButtons();
            SetStatus(_mode == Mode.AddPoints
                ? "Capture mode active: click the tip of each bar or each data point. Click again to stop."
                : "Capture paused.");
        }

        // ---- Automatic bar detection ----

        private void StartAutoDetect()
        {
            if (_image == null) { SetStatus("Please load an image first."); return; }
            if (_calibration == null || !_calibration.IsValid) { SetStatus("Please calibrate the value axis first."); return; }
            _mode = Mode.SampleColor;
            HighlightModeButtons();
            SetStatus(CurrentKind == ChartKind.Lines
                ? "Click the line to sample its color - after that you will be asked for the number of data points."
                : "Click a bar or column to sample its color - then all matching bars will be detected.");
        }

        private void RunAutoDetect(System.Drawing.Color color)
        {
            _mode = Mode.None;
            HighlightModeButtons();
            try
            {
                if (CurrentKind == ChartKind.Lines) { RunLineDetect(color); return; }

                var detected = BarDetector.Detect(_image, color, 60.0, CurrentKind);
                if (detected.Count == 0)
                {
                    SetStatus("No bars of this color were found. Check the color or tolerance, or add points manually.");
                    return;
                }
                ClearPoints();
                foreach (var pt in detected)
                    AddDataPoint(new Point(pt.X, pt.Y));
                SetStatus($"{detected.Count} bars detected. Please review or correct them, then send them to Excel.");
            }
            catch (Exception ex)
            {
                SetStatus("Auto-detection failed: " + ex.Message);
            }
        }

        private void RunLineDetect(System.Drawing.Color color)
        {
            string answer = PromptDialog.Show(this, "Data Points",
                "Number of data points on the line (for example, the number of X categories):", "12");
            if (answer == null) { SetStatus("Line detection cancelled."); return; }
            if (!int.TryParse(answer.Trim(), out int count) || count < 1)
            {
                SetStatus("Please enter a valid count (>= 1).");
                return;
            }

            var traced = LineTracer.Trace(_image, color, 60.0, count);
            if (traced.Count == 0)
            {
                SetStatus("No line of this color was found. Check the color or tolerance, or add points manually.");
                return;
            }
            ClearPoints();
            foreach (var pt in traced)
                AddDataPoint(new Point(pt.X, pt.Y));
            SetStatus($"{traced.Count} points detected on the line. Please review or correct them, then send them to Excel.");
        }

        private System.Drawing.Color SamplePixel(Point p)
        {
            int x = Math.Max(0, Math.Min(_image.Width - 1, (int)p.X));
            int y = Math.Max(0, Math.Min(_image.Height - 1, (int)p.Y));
            return _image.GetPixel(x, y);
        }

        private void AddDataPoint(Point p)
        {
            double value = _calibration.ValueAt(p.X, p.Y);
            var point = new ChartDataPoint
            {
                PixelX = p.X,
                PixelY = p.Y,
                Value = value,
                Category = (_rows.Count + 1).ToString()
            };
            Result.Points.Add(point);

            var marker = new Ellipse
            {
                Width = 10, Height = 10,
                Fill = new SolidColorBrush(Color.FromRgb(0x3B, 0x82, 0xF6)),
                Stroke = Brushes.White, StrokeThickness = 1
            };
            Canvas.SetLeft(marker, p.X - 5);
            Canvas.SetTop(marker, p.Y - 5);
            _canvas.Children.Add(marker);

            var idx = new TextBlock { Text = (_rows.Count + 1).ToString(), FontSize = 10, Foreground = Brushes.Black };
            Canvas.SetLeft(idx, p.X + 7);
            Canvas.SetTop(idx, p.Y - 16);
            _canvas.Children.Add(idx);

            var row = new Row { Point = point, Marker = marker, Label = idx };
            row.RowUi = BuildRowUi(row);
            _rows.Add(row);
            _pointsPanel.Children.Add(row.RowUi);
            SetStatus($"{_rows.Count} point(s) captured.");
        }

        private FrameworkElement BuildRowUi(Row row)
        {
            var panel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 2, 0, 2) };
            panel.Children.Add(new TextBlock { Text = "#" + (_rows.Count + 1), Width = 26, VerticalAlignment = VerticalAlignment.Center });

            var cat = new TextBox { Text = row.Point.Category, Width = 120, Margin = new Thickness(2, 0, 2, 0) };
            cat.TextChanged += (s, e) => row.Point.Category = cat.Text;
            panel.Children.Add(cat);

            var val = new TextBlock
            {
                Text = row.Point.Value.ToString("0.####", CultureInfo.CurrentCulture),
                Width = 80, VerticalAlignment = VerticalAlignment.Center
            };
            panel.Children.Add(val);

            var del = new Button { Content = "×", Width = 24, Padding = new Thickness(0) };
            del.Click += (s, e) => RemoveRow(row);
            panel.Children.Add(del);
            return panel;
        }

        private void RemoveRow(Row row)
        {
            _canvas.Children.Remove(row.Marker);
            _canvas.Children.Remove(row.Label);
            _pointsPanel.Children.Remove(row.RowUi);
            _rows.Remove(row);
            Result.Points.Remove(row.Point);
            SetStatus($"{_rows.Count} point(s) captured.");
        }

        private void ClearPoints()
        {
            foreach (var r in _rows)
            {
                _canvas.Children.Remove(r.Marker);
                _canvas.Children.Remove(r.Label);
            }
            _rows.Clear();
            _pointsPanel.Children.Clear();
            Result.Points.Clear();
        }

        // ---- Canvas dispatch ----

        private void OnCanvasClick(object sender, MouseButtonEventArgs e)
        {
            if (_image == null) return;
            var p = e.GetPosition(_canvas);
            if (p.X < 0 || p.Y < 0 || p.X > _canvas.Width || p.Y > _canvas.Height) return;

            if (_mode == Mode.Calibrate) HandleCalibrationClick(p);
            else if (_mode == Mode.AddPoints) AddDataPoint(p);
            else if (_mode == Mode.SampleColor) RunAutoDetect(SamplePixel(p));
        }

        private void HighlightModeButtons()
        {
            _calibrateButton.FontWeight = _mode == Mode.Calibrate ? FontWeights.Bold : FontWeights.Normal;
            _addPointsButton.FontWeight = _mode == Mode.AddPoints ? FontWeights.Bold : FontWeights.Normal;
        }

        private void Accept()
        {
            if (Result.Points.Count == 0) { SetStatus("No data points have been captured yet."); return; }
            Result.Kind = CurrentKind;
            DialogResult = true;
            Close();
        }

        // ---- Helpers ----

        /// <summary>Returns the OCR-recognised number nearest the click (within a radius) as text, or "".</summary>
        private string NearestOcrValue(Point p)
        {
            if (_ocrNumbers == null || _ocrNumbers.Count == 0) return string.Empty;
            const double maxDistance = 120.0;
            double best = double.MaxValue;
            double bestValue = 0;
            bool found = false;
            foreach (var n in _ocrNumbers)
            {
                double dx = n.CenterX - p.X;
                double dy = n.CenterY - p.Y;
                double d = Math.Sqrt(dx * dx + dy * dy);
                if (d < best && d <= maxDistance) { best = d; bestValue = n.Value; found = true; }
            }
            return found ? bestValue.ToString("0.####", CultureInfo.CurrentCulture) : string.Empty;
        }

        private static double? ParseNumber(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            text = text.Trim();
            if (double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out double v)) return v;
            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return v;
            return null;
        }

        protected override void OnClosed(EventArgs e)
        {
            _image?.Dispose();
            _image = null;
            base.OnClosed(e);
        }
    }
}
