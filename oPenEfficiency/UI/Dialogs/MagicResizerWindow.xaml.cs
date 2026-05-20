using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public partial class MagicResizerWindow : Window
    {
        private PowerPoint.Application _application;

        public MagicResizerWindow()
        {
            InitializeComponent();
            _application = Globals.ThisAddIn.Application;
            LoadInitialDimensions();
        }

        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
                this.DragMove();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void LoadInitialDimensions()
        {
            try
            {
                var selection = _application.ActiveWindow.Selection;
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 0)
                {
                    // Master shape is the first one
                    var master = selection.ShapeRange[1];
                    // Convert points to cm (1 inch = 72 points, 1 inch = 2.54 cm)
                    // points * 2.54 / 72
                    double wCm = master.Width * 2.54 / 72.0;
                    double hCm = master.Height * 2.54 / 72.0;

                    TxtWidth.Text = Math.Round(wCm, 2).ToString();
                    TxtHeight.Text = Math.Round(hCm, 2).ToString();
                }
            }
            catch { }
        }

        private void BtnResize_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selection = _application.ActiveWindow.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count == 0)
                    return;

                var shapes = selection.ShapeRange;
                var master = shapes[1]; // First shape as reference if needed

                // Determine Scaling Factors
                double scaleX = 1.0;
                double scaleY = 1.0;

                if (RadioPercentage.IsChecked == true)
                {
                    if (double.TryParse(TxtPercentage.Text, out double pct))
                    {
                        scaleX = pct / 100.0;
                        scaleY = pct / 100.0;
                    }
                }
                else
                {
                    // Absolute mode: calculate factor based on Master shape
                    if (double.TryParse(TxtWidth.Text, out double wTargetCm) && master.Width > 0)
                    {
                        double wTargetPts = wTargetCm * 72.0 / 2.54;
                        scaleX = wTargetPts / master.Width;
                    }
                    if (double.TryParse(TxtHeight.Text, out double hTargetCm) && master.Height > 0)
                    {
                        double hTargetPts = hTargetCm * 72.0 / 2.54;
                        scaleY = hTargetPts / master.Height;
                    }
                }

                if (scaleX == 1.0 && scaleY == 1.0) return;

                // Determine Anchor logic
                // 0=Left, 0.5=Center, 1=Right
                double anchorX = 0.5; 
                double anchorY = 0.5;

                if (AnchorTL.IsChecked == true || AnchorCL.IsChecked == true || AnchorBL.IsChecked == true) anchorX = 0.0;
                if (AnchorTR.IsChecked == true || AnchorCR.IsChecked == true || AnchorBR.IsChecked == true) anchorX = 1.0;
                
                if (AnchorTL.IsChecked == true || AnchorTC.IsChecked == true || AnchorTR.IsChecked == true) anchorY = 0.0;
                if (AnchorBL.IsChecked == true || AnchorBC.IsChecked == true || AnchorBR.IsChecked == true) anchorY = 1.0;

                // Options
                bool scaleFont = SwitchFont.IsChecked == true;
                bool scaleLine = SwitchLine.IsChecked == true;
                bool scaleMargin = SwitchMargin.IsChecked == true;
                bool scaleIndent = SwitchIndent.IsChecked == true;

                // Apply resizing
                // We iterate 1..Count (COM collection is 1-based)
                for (int i = 1; i <= shapes.Count; i++)
                {
                    var shp = shapes[i];

                    // Calculate old center/anchor before resize
                    float oldW = shp.Width;
                    float oldH = shp.Height;
                    float oldL = shp.Left;
                    float oldT = shp.Top;

                    // Compute fixed point based on anchor
                    float fixedX = oldL + (oldW * (float)anchorX);
                    float fixedY = oldT + (oldH * (float)anchorY);

                    // Resize
                    float newW = oldW * (float)scaleX;
                    float newH = oldH * (float)scaleY;

                    shp.Width = newW;
                    shp.Height = newH;

                    // Reposition to maintain anchor
                    // fixedX = newL + (newW * anchorX)  =>  newL = fixedX - (newW * anchorX)
                    shp.Left = fixedX - (newW * (float)anchorX);
                    shp.Top = fixedY - (newH * (float)anchorY);

                    // --- Smart Scaling ---
                    double avgFactor = (scaleX + scaleY) / 2.0; // Use average for scalar properties

                    // 1. Line Weight
                    if (scaleLine && shp.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                        shp.Line.Weight *= (float)avgFactor;
                    }

                    // 2. Text Properties
                    try
                    {
                        if (shp.HasTextFrame == Office.MsoTriState.msoTrue && shp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            var textRange = shp.TextFrame.TextRange;

                            // Font Size
                            if (scaleFont)
                            {
                                // Ideally iterate runs if mixed sizes, but simplistic approach first:
                                float newSize = textRange.Font.Size * (float)avgFactor;
                                if (newSize < 1) newSize = 1;
                                textRange.Font.Size = newSize;
                            }

                            // Margins
                            if (scaleMargin)
                            {
                                shp.TextFrame.MarginLeft *= (float)scaleX;
                                shp.TextFrame.MarginRight *= (float)scaleX;
                                shp.TextFrame.MarginTop *= (float)scaleY;
                                shp.TextFrame.MarginBottom *= (float)scaleY;
                            }

                            // Indents & Bullets
                            if (scaleIndent)
                            {
                                try
                                {
                                    // Apply to the whole range's format directly
                                    textRange.ParagraphFormat.Bullet.RelativeSize *= (float)avgFactor;
                                }
                                catch { }
                            }
                        }
                    }
                    catch { }                }
                
                // Optional: Close or keep open?
                // this.Close(); 
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error resizing: " + ex.Message);
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
