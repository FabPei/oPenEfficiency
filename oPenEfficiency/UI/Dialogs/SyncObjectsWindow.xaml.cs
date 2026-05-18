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

            _tempDir = Path.Combine(Path.GetTempPath(), "oPenEfficiency_Sync_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);

            Loaded += OnLoaded;
            Closed += OnClosed;
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
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
                catch { }

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
            var itemBorder = new Border
            {
                Width = ThumbW + 8,
                Height = ThumbH + 26,
                Margin = new Thickness(5),
                Padding = new Thickness(3),
                BorderThickness = new Thickness(2),
                BorderBrush = Brushes.Transparent,
                CornerRadius = new CornerRadius(4),
                Cursor = Cursors.Hand,
                Tag = slideIndex
            };

            var stack = new StackPanel { Orientation = Orientation.Vertical };

            var imgGrid = new Grid
            {
                Width = ThumbW,
                Height = ThumbH
            };

            var img = new System.Windows.Controls.Image
            {
                Source = thumbnail,
                Width = ThumbW,
                Height = ThumbH,
                Stretch = Stretch.Fill,
                SnapsToDevicePixels = true
            };

            var selectionOverlay = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(100, 59, 130, 246)),
                Visibility = Visibility.Collapsed
            };

            var checkmark = new TextBlock
            {
                Text = "✓",
                Foreground = Brushes.White,
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Visibility = Visibility.Collapsed
            };

            imgGrid.Children.Add(img);
            imgGrid.Children.Add(selectionOverlay);
            imgGrid.Children.Add(checkmark);

            var numberLabel = new TextBlock
            {
                Text = slideIndex.ToString(),
                FontSize = 10,
                Foreground = new SolidColorBrush(Color.FromRgb(0x9C, 0xA3, 0xAF)),
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 4, 0, 0)
            };

            stack.Children.Add(imgGrid);
            stack.Children.Add(numberLabel);
            itemBorder.Child = stack;

            Action<bool> setSelectionState = (selected) =>
            {
                if (selected)
                {
                    _selectedIndices.Add(slideIndex);
                    itemBorder.BorderBrush = (Brush)FindResource("AccentBrush") ?? new SolidColorBrush(Color.FromRgb(59, 130, 246));
                    itemBorder.Background = new SolidColorBrush(Color.FromArgb(20, 59, 130, 246));
                    selectionOverlay.Visibility = Visibility.Visible;
                    checkmark.Visibility = Visibility.Visible;
                }
                else
                {
                    _selectedIndices.Remove(slideIndex);
                    itemBorder.BorderBrush = Brushes.Transparent;
                    itemBorder.Background = Brushes.Transparent;
                    selectionOverlay.Visibility = Visibility.Collapsed;
                    checkmark.Visibility = Visibility.Collapsed;
                }
            };

            setSelectionState(isSelectedInitially);

            itemBorder.MouseLeftButtonUp += (s, e) =>
            {
                bool isCurrentlySelected = _selectedIndices.Contains(slideIndex);
                setSelectionState(!isCurrentlySelected);
            };

            itemBorder.MouseEnter += (s, e) =>
            {
                if (!_selectedIndices.Contains(slideIndex))
                {
                    itemBorder.BorderBrush = (Brush)FindResource("BorderLightBrush") ?? Brushes.LightGray;
                }
            };
            itemBorder.MouseLeave += (s, e) =>
            {
                if (!_selectedIndices.Contains(slideIndex))
                {
                    itemBorder.BorderBrush = Brushes.Transparent;
                }
            };

            SlidesPanel.Children.Add(itemBorder);
        }

        private void BtnDistribute_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedIndices.Count == 0)
            {
                ShowStatus("No target slides selected.", Colors.Orange);
                return;
            }

            bool success = SyncObjectsFeature.DistributeShapesToSlides(_manager, _selectedIndices.ToList());
            if (success)
            {
                ShowStatus("Shapes distributed successfully!", Colors.LightGreen);
            }
            else
            {
                ShowStatus("Failed to distribute. Ensure shapes are selected.", Colors.Red);
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
                ShowStatus("Synchronization complete!", Colors.LightGreen);
            }
            else
            {
                ShowStatus("Sync failed. Check if a valid tagged shape is selected.", Colors.Red);
            }
        }

        private void ShowStatus(string message, Color color)
        {
            TxtStatus.Text = message;
            TxtStatus.Foreground = new SolidColorBrush(color);
            TxtStatus.Visibility = Visibility.Visible;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}