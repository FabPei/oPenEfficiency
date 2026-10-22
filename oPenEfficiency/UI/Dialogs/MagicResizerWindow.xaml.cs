using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using oPenEfficiency.Core.Services;
using oPenEfficiency.Features;
using oPenEfficiency.Utils;
using ScaleAnchor = oPenEfficiency.Core.Services.MagicResizerScalingPolicy.ScaleAnchor;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public partial class MagicResizerWindow : Window
    {
        private readonly PowerPointManager _manager;
        private readonly PowerPoint.Application _application;

        public MagicResizerWindow(PowerPointManager manager)
        {
            _manager = manager ?? throw new ArgumentNullException(nameof(manager));
            InitializeComponent();
            _application = _manager.GetApplication();
            if (_application != null)
                new WindowInteropHelper(this).Owner = new IntPtr(_application.HWND);
            LoadInitialDimensions();
            SourceInitialized += (sender, args) => FitToOwnerWorkArea();
        }

        private void FitToOwnerWorkArea()
        {
            var helper = new WindowInteropHelper(this);
            var screen = System.Windows.Forms.Screen.FromHandle(helper.Owner);
            var dpi = VisualTreeHelper.GetDpi(this);
            double availableWidth = Math.Max(1, (screen.WorkingArea.Width / dpi.DpiScaleX) - 32);
            double availableHeight = Math.Max(1, (screen.WorkingArea.Height / dpi.DpiScaleY) - 32);
            MaxWidth = availableWidth;
            MaxHeight = availableHeight;
            Width = Math.Min(Width, availableWidth);
        }

        private void Window_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape)
            {
                Close();
                e.Handled = true;
            }
        }

        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left &&
                e.ButtonState == System.Windows.Input.MouseButtonState.Pressed)
                this.DragMove();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void LoadInitialDimensions()
        {
            var shapes = _manager.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return;

            var master = shapes[1];
            double wCm = MagicResizerScalingPolicy.PointsToCentimetres(master.Width);
            double hCm = MagicResizerScalingPolicy.PointsToCentimetres(master.Height);
            TxtWidth.Text = Math.Round(wCm, 2).ToString();
            TxtHeight.Text = Math.Round(hCm, 2).ToString();
        }

        private void BtnResize_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var shapes = _manager.GetSelectedShapes();
                // REG-MAGICRESIZER-GROUP-001: a single group is a valid scale target and scales as one unit.
                if (shapes == null || shapes.Count == 0)
                {
                    MessageBox.Show(
                        "Select at least one shape.",
                        "Scale Selected Shapes",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }
                var master = shapes[1]; // First shape as reference if needed

                // Determine Scaling Factors
                double scaleX;
                double scaleY;

                if (RadioPercentage.IsChecked == true)
                {
                    if (!double.TryParse(TxtPercentage.Text, out double pct) ||
                        !MagicResizerScalingPolicy.TryPercentageScaleFactors(pct, out scaleX, out scaleY))
                    {
                        MessageBox.Show("Enter a percentage greater than zero.", "Scale Selected Shapes", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }
                else
                {
                    // Absolute mode: calculate factor based on Master shape
                    if (!double.TryParse(TxtWidth.Text, out double wTargetCm) ||
                        !double.TryParse(TxtHeight.Text, out double hTargetCm) ||
                        !MagicResizerScalingPolicy.TryAbsoluteScaleFactors(wTargetCm, hTargetCm, master.Width, master.Height, out scaleX, out scaleY))
                    {
                        MessageBox.Show("Enter a width and height greater than zero.", "Scale Selected Shapes", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }

                if (MagicResizerScalingPolicy.IsNoOp(scaleX, scaleY)) return;

                // Determine Anchor logic (the nine anchor points of the 3x3 grid)
                ScaleAnchor anchor = ScaleAnchor.Center;
                if (AnchorTL.IsChecked == true) anchor = ScaleAnchor.TopLeft;
                else if (AnchorTC.IsChecked == true) anchor = ScaleAnchor.Top;
                else if (AnchorTR.IsChecked == true) anchor = ScaleAnchor.TopRight;
                else if (AnchorCL.IsChecked == true) anchor = ScaleAnchor.Left;
                else if (AnchorCR.IsChecked == true) anchor = ScaleAnchor.Right;
                else if (AnchorBL.IsChecked == true) anchor = ScaleAnchor.BottomLeft;
                else if (AnchorBC.IsChecked == true) anchor = ScaleAnchor.Bottom;
                else if (AnchorBR.IsChecked == true) anchor = ScaleAnchor.BottomRight;

                // Options
                bool scaleFont = SwitchFont.IsChecked == true;
                bool scaleLine = SwitchLine.IsChecked == true;
                bool scaleMargin = SwitchMargin.IsChecked == true;
                bool scaleIndent = SwitchIndent.IsChecked == true;

                try { _application?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ScaleSelectedShapes.StartNewUndoEntry"); }

                // Apply resizing
                // We iterate 1..Count (COM collection is 1-based)
                double avgFactor = MagicResizerScalingPolicy.ScalarFactor(scaleX, scaleY); // Use average for scalar properties
                for (int i = 1; i <= shapes.Count; i++)
                {
                    var shp = shapes[i];

                    // Calculate the scaled bounds around the anchor before writing them back
                    MagicResizerScalingPolicy.ComputeScaledBounds(
                        shp.Left, shp.Top, shp.Width, shp.Height,
                        scaleX, scaleY, anchor,
                        out float newLeft, out float newTop, out float newW, out float newH);

                    // Resize
                    shp.Width = newW;
                    shp.Height = newH;

                    // Reposition to maintain anchor
                    shp.Left = newLeft;
                    shp.Top = newTop;

                    // --- Smart Scaling ---

                    // 1. Line Weight
                    if (scaleLine && shp.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                        shp.Line.Weight = MagicResizerScalingPolicy.ScaleScalar(shp.Line.Weight, avgFactor);
                    }

                    // 2. Text Properties
                    if (shp.HasTextFrame == Office.MsoTriState.msoTrue && shp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        var textRange = shp.TextFrame.TextRange;

                        // Font Size
                        if (scaleFont)
                        {
                            // Ideally iterate runs if mixed sizes, but simplistic approach first:
                            textRange.Font.Size = MagicResizerScalingPolicy.ScaledFontSize(textRange.Font.Size, avgFactor);
                        }

                        // Margins
                        if (scaleMargin)
                        {
                            shp.TextFrame.MarginLeft = MagicResizerScalingPolicy.ScaleScalar(shp.TextFrame.MarginLeft, scaleX);
                            shp.TextFrame.MarginRight = MagicResizerScalingPolicy.ScaleScalar(shp.TextFrame.MarginRight, scaleX);
                            shp.TextFrame.MarginTop = MagicResizerScalingPolicy.ScaleScalar(shp.TextFrame.MarginTop, scaleY);
                            shp.TextFrame.MarginBottom = MagicResizerScalingPolicy.ScaleScalar(shp.TextFrame.MarginBottom, scaleY);
                        }

                        // Indents & Bullets
                        if (scaleIndent)
                        {
                            try
                            {
                                // Apply to the whole range's format directly
                                textRange.ParagraphFormat.Bullet.RelativeSize =
                                    MagicResizerScalingPolicy.ScaleScalar(textRange.ParagraphFormat.Bullet.RelativeSize, avgFactor);

                                // For indents (LeftIndent, FirstLineIndent), we might need to iterate paragraphs
                                // For Indents, PowerPoint doesn't expose strict LeftIndent/FirstLineIndent on ParagraphFormat easily
                                // without going through complex Ruler objects or IndentLevels.
                                // We will skip specific indent scaling for now to fix compilation.
                                // textRange.ParagraphFormat.LeftIndent *= (float)scaleX; // Not available
                                // textRange.ParagraphFormat.FirstLineIndent *= (float)scaleX; // Not available
                            }
                            catch { /* A stale preview shape is skipped while restoring the remaining preview. */ }
                        }
                    }
                }

                // Optional: Close or keep open?
                // this.Close();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ScaleSelectedShapes.Resize");
                MessageBox.Show(
                    "The selected shapes could not be scaled completely. " + ex.Message,
                    "Scale Selected Shapes",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void RadioMode_Checked(object sender, RoutedEventArgs e)
        {
            if (PanelPercentage == null || PanelAbsolute == null) return; // During init

            if (RadioPercentage.IsChecked == true)
            {
                PanelPercentage.Visibility = Visibility.Visible;
                PanelAbsolute.Visibility = Visibility.Collapsed;
            }
            else
            {
                PanelPercentage.Visibility = Visibility.Collapsed;
                PanelAbsolute.Visibility = Visibility.Visible;
                // Refresh dimensions in case selection changed?
                LoadInitialDimensions();
            }
        }
    }
}
