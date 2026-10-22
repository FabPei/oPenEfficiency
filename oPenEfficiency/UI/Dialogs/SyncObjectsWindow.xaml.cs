using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Features;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    public partial class SyncObjectsWindow : Window
    {
        private PowerPointManager _manager;
        private readonly HashSet<int> _selectedIndices = new HashSet<int>();
        private readonly string _tempDir;
        
        // Thumbnail display dimensions
        private const double ThumbW = 112.0;
        private const double ThumbH = 63.0;

        public SyncObjectsWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            SourceInitialized += (sender, args) => FitToOwnerWorkArea();

            _tempDir = Path.Combine(Path.GetTempPath(), "oPenEfficiency_Sync_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);

            Loaded += OnLoaded;
            Loaded += (sender, args) => UpdateHeaderCaption();
            HeaderDragRegion.SizeChanged += (sender, args) => UpdateHeaderCaption();
            Closed += OnClosed;
        }

        private void FitToOwnerWorkArea()
        {
            var helper = new System.Windows.Interop.WindowInteropHelper(this);
            var screen = System.Windows.Forms.Screen.FromHandle(helper.Owner);
            var dpi = VisualTreeHelper.GetDpi(this);
            double availableWidth = Math.Max(1, (screen.WorkingArea.Width / dpi.DpiScaleX) - 32);
            double availableHeight = Math.Max(1, (screen.WorkingArea.Height / dpi.DpiScaleY) - 32);
            MinWidth = Math.Min(MinWidth, availableWidth);
            MinHeight = Math.Min(MinHeight, availableHeight);
            MaxWidth = availableWidth;
            MaxHeight = availableHeight;
            Width = Math.Min(Width, availableWidth);
            Height = Math.Min(Height, availableHeight);
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
                e.Handled = true;
            }
        }

        private void UpdateHeaderCaption()
        {
            var chrome = System.Windows.Shell.WindowChrome.GetWindowChrome(this);
            if (chrome == null || HeaderDragRegion.ActualHeight <= 0) return;
            // REG-SYNC-WINDOW-UI-001: native caption hit testing includes the shell's top padding
            // and follows header wrapping/DPI layout. Close remains an interactive client control.
            chrome.CaptionHeight = HeaderDragRegion.TranslatePoint(
                new System.Windows.Point(0, HeaderDragRegion.ActualHeight), this).Y;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(LoadThumbnails),
                System.Windows.Threading.DispatcherPriority.Background);
        }

        private void OnClosed(object sender, EventArgs e)
        {
            try { Directory.Delete(_tempDir, true); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "SyncObjectsWindow.Cleanup"); }
        }

        private void LoadThumbnails()
        {
            try
            {
                var app = _manager.GetApplication();
                var pres = app.ActivePresentation;
                if (pres == null) return;

                int count = pres.Slides.Count;
                if (count == 0) return;

                int activeIndex = -1;
                try
                {
                    if (app.ActiveWindow?.View?.Slide != null)
                        activeIndex = ((Slide)app.ActiveWindow.View.Slide).SlideIndex;
                }
                catch { /* A stale candidate shape is skipped while populating the synchronization list. */ }

                for (int i = 1; i <= count; i++)
                {
                    int slideIndex = i;
                    string thumbPath = Path.Combine(_tempDir, $"thumb_{i}.jpg");

                    try { pres.Slides[i].Export(thumbPath, "JPG", 320, 180); }
                    catch (Exception ex) { ExceptionLogger.Log(ex, $"SyncObjectsWindow.LoadThumbnails slide {i}"); continue; }

                    if (!File.Exists(thumbPath)) continue;

                    BitmapImage bmp = null;
                    try
                    {
                        bmp = new BitmapImage();
                        bmp.BeginInit();
                        bmp.CacheOption = BitmapCacheOption.OnLoad;
                        bmp.UriSource = new Uri(thumbPath);
                        bmp.EndInit();
                        bmp.Freeze();
                    }
                    catch (Exception ex) { ExceptionLogger.Log(ex, $"SyncObjectsWindow.BitmapLoad {slideIndex}"); continue; }

                    var capturedBmp = bmp;
                    int capturedIndex = slideIndex;
                    bool isSelected = capturedIndex != activeIndex; // Select all but active by default

                    Dispatcher.Invoke(() => AddSlideItem(capturedIndex, capturedBmp, isSelected));
                }

                Dispatcher.Invoke(() =>
                {
                    LoadingPanel.Visibility = Visibility.Collapsed;
                    SlidesScrollViewer.Visibility = Visibility.Visible;
                });
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SyncObjectsWindow.LoadThumbnails");
            }
        }

        private void AddSlideItem(int slideIndex, BitmapImage thumbnail, bool isSelectedInitially)
        {
            var stack = new StackPanel();
            stack.Children.Add(new System.Windows.Controls.Image
            {
                Source = thumbnail,
                Width = ThumbW,
                Height = ThumbH,
                Stretch = Stretch.Uniform
            });
            var numberLabel = new TextBlock
            {
                Text = $"Slide {slideIndex}",
                FontSize = (double)FindResource("TypographyCaptionSize"),
                HorizontalAlignment = HorizontalAlignment.Center
            };
            numberLabel.SetResourceReference(TextBlock.ForegroundProperty, "TextSecondaryBrush");
            stack.Children.Add(numberLabel);

            var item = new System.Windows.Controls.CheckBox
            {
                Content = stack,
                Width = ThumbW + 36,
                MinHeight = ThumbH + 26,
                Margin = (Thickness)FindResource("Space4"),
                VerticalContentAlignment = VerticalAlignment.Center,
                IsChecked = isSelectedInitially,
                ToolTip = $"Target slide {slideIndex} (Space to select or clear)"
            };
            item.Style = (Style)FindResource(typeof(System.Windows.Controls.CheckBox));
            System.Windows.Automation.AutomationProperties.SetName(item, $"Target slide {slideIndex}");
            if (isSelectedInitially) _selectedIndices.Add(slideIndex);
            else _selectedIndices.Remove(slideIndex);
            item.Checked += (sender, args) => _selectedIndices.Add(slideIndex);
            item.Unchecked += (sender, args) => _selectedIndices.Remove(slideIndex);
            SlidesPanel.Children.Add(item);
        }

        private void BtnDistribute_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedIndices.Count == 0)
            {
                ShowStatus("No target slides selected.", "WarningBrush");
                return;
            }

            bool success = SyncObjectsFeature.DistributeShapesToSlides(_manager, _selectedIndices.ToList());
            if (success)
            {
                ShowStatus("Shapes distributed successfully!", "SuccessBrush");
            }
            else
            {
                ShowStatus("Failed to distribute. Ensure shapes are selected.", "ErrorBrush");
            }
        }

        private void BtnSyncFormat_Click(object sender, RoutedEventArgs e) => Sync(true, false, false);
        private void BtnSyncPosition_Click(object sender, RoutedEventArgs e) => Sync(false, true, false);
        private void BtnSyncContent_Click(object sender, RoutedEventArgs e) => Sync(false, false, true);

        private void Sync(bool format, bool position, bool content)
        {
            bool success = SyncObjectsFeature.SyncShapeProperties(_manager, format, position, content);
            if (success)
            {
                ShowStatus("Synchronization complete!", "SuccessBrush");
            }
            else
            {
                ShowStatus("Sync failed. Check if a valid tagged shape is selected.", "ErrorBrush");
            }
        }

        private void ShowStatus(string message, string brushKey)
        {
            TxtStatus.Text = message;
            TxtStatus.SetResourceReference(TextBlock.ForegroundProperty, brushKey);
            TxtStatus.Visibility = Visibility.Visible;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
