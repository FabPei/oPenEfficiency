using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    public partial class SlidePasteWindow : Window
    {
        // ─── Constants ────────────────────────────────────────────────────────
        private const float CmToPt = oPenEfficiency.Core.Services.SnapshotGridLayoutPolicy.CmToPt;
        private const float MarginCm = 1.5f;
        private const float GapCm = 0.3f;
        private const float DefaultWidthCm = 7.11f;
        private const float DefaultHeightCm = 4.0f;

        // Thumbnail display dimensions in the dialog (16:9 proportional)
        private const double ThumbW = 112.0;
        private const double ThumbH = 63.0;

        // ─── Fields ────────────────────────────────────────────────────────────
        private readonly PowerPointManager _manager;
        private readonly HashSet<int> _selectedSlideIds = new HashSet<int>();
        private readonly string _tempDir;
        private Presentation _sourcePresentation;

        // ─── Construction ─────────────────────────────────────────────────────
        public SlidePasteWindow(PowerPointManager manager, bool addLinks = false, bool addShadows = false)
        {
            InitializeComponent();
            _manager = manager;
            ChkAddLink.IsChecked = addLinks;
            ChkAddShadow.IsChecked = addShadows;

            _tempDir = Path.Combine(Path.GetTempPath(), "oPenEfficiency_SlidePaste_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);

            Loaded += OnLoaded;
            Closed += OnClosed;
        }

        public void SetOptions(bool addLinks, bool addShadows)
        {
            ChkAddLink.IsChecked = addLinks;
            ChkAddShadow.IsChecked = addShadows;
        }

        // ─── Window events ────────────────────────────────────────────────────
        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            // Defer thumbnail generation so the window renders first
            Dispatcher.BeginInvoke(new Action(LoadThumbnails),
                System.Windows.Threading.DispatcherPriority.Background);
        }

        private void OnClosed(object sender, EventArgs e)
        {
            // Clean up exported temp files
            try { Directory.Delete(_tempDir, recursive: true); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "SlidePasteWindow.OnClosed.Cleanup"); }
        }

        // ─── Thumbnail loading ────────────────────────────────────────────────
        private void LoadThumbnails()
        {
            try
            {
                var app = _manager.GetApplication();
                var pres = app.ActivePresentation;
                if (pres == null) return;
                _sourcePresentation = pres;

                int count = pres.Slides.Count;
                if (count == 0) return;

                // Export each slide as a small JPG for the thumbnail grid
                for (int i = 1; i <= count; i++)
                {
                    int slideIndex = i; // capture for lambda
                    string thumbPath = Path.Combine(_tempDir, $"thumb_{i}.jpg");

                    try
                    {
                        pres.Slides[i].Export(thumbPath, "JPG", 320, 180);
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, $"SlidePasteWindow.LoadThumbnails slide {i}");
                        continue;
                    }

                    if (!File.Exists(thumbPath)) continue;

                    // Load image (must keep stream open until freeze)
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
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, $"SlidePasteWindow.LoadThumbnail.BitmapLoad {slideIndex}");
                        continue;
                    }

                    var capturedBmp = bmp;
                    int capturedIndex = slideIndex;

                    int slideId = pres.Slides[i].SlideID;
                    Dispatcher.Invoke(() => AddSlideItem(capturedIndex, slideId, capturedBmp));
                }

                Dispatcher.Invoke(() =>
                {
                    LoadingPanel.Visibility = Visibility.Collapsed;
                    SlidesScrollViewer.Visibility = Visibility.Visible;
                });
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SlidePasteWindow.LoadThumbnails");
            }
        }

        private void AddSlideItem(int slideIndex, int slideId, BitmapImage thumbnail)
        {
            // Outer container
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
                Tag = slideId
            };

            var stack = new StackPanel { Orientation = Orientation.Vertical };

            // Thumbnail image with selection overlay via Grid
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

            // Selection overlay (blue tint)
            var selectionOverlay = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(100, 59, 130, 246)),
                Visibility = Visibility.Collapsed
            };

            // Checkmark shown when selected
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

            // Slide number label
            var numberLabel = new TextBlock
            {
                Text = slideIndex.ToString(),
                FontSize = 10,
                Foreground = new SolidColorBrush(Color.FromRgb(0x9C, 0xA3, 0xAF)),
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 3, 0, 0)
            };

            stack.Children.Add(imgGrid);
            stack.Children.Add(numberLabel);
            itemBorder.Child = stack;

            // Click handler
            itemBorder.MouseLeftButtonDown += (s, e) =>
            {
                int id = (int)itemBorder.Tag;
                if (_selectedSlideIds.Contains(id))
                {
                    _selectedSlideIds.Remove(id);
                    itemBorder.BorderBrush = Brushes.Transparent;
                    selectionOverlay.Visibility = Visibility.Collapsed;
                    checkmark.Visibility = Visibility.Collapsed;
                }
                else
                {
                    _selectedSlideIds.Add(id);
                    itemBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(0x3B, 0x82, 0xF6));
                    selectionOverlay.Visibility = Visibility.Visible;
                    checkmark.Visibility = Visibility.Visible;
                }
                UpdateStatus();
            };

            // Hover effect
            itemBorder.MouseEnter += (s, e) =>
            {
                if (!_selectedSlideIds.Contains((int)itemBorder.Tag))
                    itemBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(0x3F, 0x3F, 0x4A));
            };
            itemBorder.MouseLeave += (s, e) =>
            {
                if (!_selectedSlideIds.Contains((int)itemBorder.Tag))
                    itemBorder.BorderBrush = Brushes.Transparent;
            };

            SlidesPanel.Children.Add(itemBorder);
        }

        private void UpdateStatus()
        {
            int n = _selectedSlideIds.Count;
            if (n == 0)
            {
                TxtStatus.Text = "No slides selected";
                BtnInsert.Content = "Insert";
                BtnInsert.IsEnabled = false;
            }
            else
            {
                TxtStatus.Text = $"{n} slide{(n == 1 ? "" : "s")} selected";
                BtnInsert.Content = $"Insert ({n})";
                BtnInsert.IsEnabled = true;
            }
        }

        // ─── Insert ───────────────────────────────────────────────────────────
        private void BtnInsert_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedSlideIds.Count == 0) return;

            var createdShapes = new List<Shape>();

            try
            {
                var app = _manager.GetApplication();
                var pres = app.ActivePresentation;
                if (pres == null || _sourcePresentation == null || !IsSameComObject(pres, _sourcePresentation))
                {
                    MessageBox.Show(
                        "The active presentation changed while Insert Slide Snapshots was open. Return to the original presentation or reopen Insert Slide Snapshots.",
                        "Insert Slide Snapshots",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                var targetSlide = _manager.GetCurrentSlide();
                if (targetSlide == null)
                {
                    MessageBox.Show("No active slide found. Please navigate to a slide first.",
                        "Insert Slide Snapshots", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                bool addLink = ChkAddLink.IsChecked == true;
                bool addShadow = ChkAddShadow.IsChecked == true;

                var sourceSlides = new List<Slide>();
                foreach (int slideId in _selectedSlideIds)
                {
                    try
                    {
                        var sourceSlide = pres.Slides.FindBySlideID(slideId);
                        if (sourceSlide != null) sourceSlides.Add(sourceSlide);
                    }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "SlidePasteWindow.ResolveSourceSlide"); }
                }
                sourceSlides = sourceSlides.OrderBy(slide => slide.SlideIndex).ToList();
                if (sourceSlides.Count == 0)
                {
                    MessageBox.Show("The selected source slides no longer exist.", "Insert Slide Snapshots", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Layout calculation
                float margin = MarginCm * CmToPt;
                float gap = GapCm * CmToPt;
                float defaultW = DefaultWidthCm * CmToPt;
                float defaultH = DefaultHeightCm * CmToPt;
                float ratio = defaultW / defaultH;

                float slideWidth = pres.PageSetup.SlideWidth;
                float slideHeight = pres.PageSetup.SlideHeight;
                float availW = slideWidth - 2 * margin;
                float availH = slideHeight - 2 * margin;

                int n = sourceSlides.Count;
                // Reduce image size until all images fit within the available slide area
                oPenEfficiency.Core.Services.SnapshotGridLayoutPolicy.FitGrid(
                    n, availW, availH, gap, defaultW, defaultH, ratio,
                    out float W, out float H, out int finalCols);

                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "SlidePasteWindow.StartNewUndoEntry"); }

                int skippedCount = 0;
                for (int i = 0; i < sourceSlides.Count; i++)
                {
                    Slide sourceSlide = sourceSlides[i];
                    int slideIndex = sourceSlide.SlideIndex;
                    int slideId = sourceSlide.SlideID;
                    string exportPath = Path.Combine(_tempDir, $"export_{slideId}.png");

                    try
                    {
                        sourceSlide.Export(exportPath, "PNG", 960, 540);
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, $"SlidePasteWindow.BtnInsert_Click.Export slide {slideIndex}");
                        skippedCount++;
                        continue;
                    }

                    if (!File.Exists(exportPath))
                    {
                        skippedCount++;
                        continue;
                    }

                    int layoutIndex = createdShapes.Count;
                    int col = layoutIndex % finalCols;
                    int row = layoutIndex / finalCols;
                    float left = margin + col * (W + gap);
                    float top = margin + row * (H + gap);

                    Shape pic = targetSlide.Shapes.AddPicture(
                        exportPath,
                        Office.MsoTriState.msoFalse,  // LinkToFile = false
                        Office.MsoTriState.msoTrue,   // SaveWithDocument = true
                        left, top, W, H);
                    pic.Name = $"oPE_SlidePaste_{slideId}";
                    pic.Tags.Add("oPE_SlidePasteSourceSlideId", slideId.ToString());
                    createdShapes.Add(pic);

                    if (addShadow)
                    {
                        try
                        {
                            pic.Shadow.Visible = Office.MsoTriState.msoTrue;
                            pic.Shadow.OffsetX = 3f;
                            pic.Shadow.OffsetY = 3f;
                            pic.Shadow.Blur = 6f;
                            pic.Shadow.Transparency = 0.5f;
                        }
                        catch (Exception ex)
                        {
                            ExceptionLogger.Log(ex, "SlidePasteWindow.Shadow");
                        }
                    }

                    if (addLink)
                    {
                        try
                        {
                            var actionSetting = pic.ActionSettings[PpMouseActivation.ppMouseClick];
                            actionSetting.Action = PpActionType.ppActionHyperlink;
                            actionSetting.Hyperlink.SubAddress = $"{slideId},{slideIndex},";
                        }
                        catch (Exception ex)
                        {
                            ExceptionLogger.Log(ex, $"SlidePasteWindow.Link slide {slideIndex}");
                        }
                    }
                }

                if (createdShapes.Count == 0)
                {
                    MessageBox.Show("No slide images could be inserted. Review the error log for export details.", "Insert Slide Snapshots", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (skippedCount > 0)
                {
                    MessageBox.Show(
                        $"Inserted {createdShapes.Count} slide image(s); {skippedCount} source slide(s) were skipped because export failed.",
                        "Insert Slide Snapshots — Incomplete",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
                Close();
            }
            catch (Exception ex)
            {
                for (int i = createdShapes.Count - 1; i >= 0; i--)
                {
                    try { createdShapes[i].Delete(); }
                    catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "SlidePasteWindow.BtnInsert_Click.Rollback"); }
                }
                ExceptionLogger.Log(ex, "SlidePasteWindow.BtnInsert_Click");
                MessageBox.Show($"An error occurred while inserting slides:\n{ex.Message}",
                    "Insert Slide Snapshots", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static bool IsSameComObject(object first, object second)
        {
            IntPtr firstPointer = IntPtr.Zero;
            IntPtr secondPointer = IntPtr.Zero;
            try
            {
                firstPointer = Marshal.GetIUnknownForObject(first);
                secondPointer = Marshal.GetIUnknownForObject(second);
                return firstPointer == secondPointer;
            }
            finally
            {
                if (firstPointer != IntPtr.Zero) Marshal.Release(firstPointer);
                if (secondPointer != IntPtr.Zero) Marshal.Release(secondPointer);
            }
        }
    }
}
