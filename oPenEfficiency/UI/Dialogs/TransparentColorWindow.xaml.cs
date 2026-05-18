using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI
{
    public partial class TransparentColorWindow : Window
    {
        private PowerPointManager _manager;
        private System.Drawing.Color _selectedColor = System.Drawing.Color.Empty;

        public TransparentColorWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;

            // Set Owner
            try
            {
                IntPtr hwnd = (IntPtr)_manager.GetApplication().HWND;
                if (hwnd != IntPtr.Zero)
                {
                    new WindowInteropHelper(this).Owner = hwnd;
                }
            } catch { }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnPickColor_Click(object sender, RoutedEventArgs e)
        {
            Window overlay = new Window
            {
                WindowStyle = WindowStyle.None,
                AllowsTransparency = true,
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(1, 255, 255, 255)),
                Cursor = System.Windows.Input.Cursors.Cross,
                Topmost = true,
                Left = SystemParameters.VirtualScreenLeft,
                Top = SystemParameters.VirtualScreenTop,
                Width = SystemParameters.VirtualScreenWidth,
                Height = SystemParameters.VirtualScreenHeight,
                ShowInTaskbar = false
            };

            overlay.MouseDown += (s, ev) =>
            {
                try
                {
                    var point = System.Windows.Forms.Control.MousePosition;
                    using (Bitmap bmp = new Bitmap(1, 1))
                    {
                        using (Graphics g = Graphics.FromImage(bmp))
                        {
                            g.CopyFromScreen(point, new System.Drawing.Point(0, 0), new System.Drawing.Size(1, 1));
                        }
                        _selectedColor = bmp.GetPixel(0, 0);
                        
                        // Update UI
                        ColorPreview.Background = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(_selectedColor.R, _selectedColor.G, _selectedColor.B));
                    }
                }
                catch { }
                finally
                {
                    overlay.Close();
                }
            };

            overlay.KeyDown += (s, ev) =>
            {
                if (ev.Key == Key.Escape) overlay.Close();
            };

            overlay.ShowDialog();
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedColor.IsEmpty)
            {
                MessageBox.Show("Please pick a color first.");
                return;
            }

            try
            {
                var app = _manager.GetApplication();
                if (app.ActiveWindow.Selection.Type != PpSelectionType.ppSelectionShapes && 
                    app.ActiveWindow.Selection.Type != PpSelectionType.ppSelectionText)
                {
                    MessageBox.Show("Please select exactly one shape to process.");
                    return;
                }

                var shapeRange = app.ActiveWindow.Selection.ShapeRange;
                if (shapeRange.Count != 1)
                {
                    MessageBox.Show("Please select exactly one shape. You have selected " + shapeRange.Count);
                    return;
                }

                Microsoft.Office.Interop.PowerPoint.Shape targetShape = shapeRange[1];

                string templateTemp = Path.GetTempFileName();
                File.Delete(templateTemp);
                string tempIn = templateTemp + "_in.png";
                string tempOut = templateTemp + "_out.png";

                // Export current shape
                targetShape.Export(tempIn, PpShapeFormat.ppShapeFormatPNG);

                // Process image transparency
                using (Bitmap src = new Bitmap(tempIn))
                {
                    using (Bitmap dest = new Bitmap(src.Width, src.Height, PixelFormat.Format32bppArgb))
                    {
                        using (Graphics g = Graphics.FromImage(dest))
                        {
                            g.DrawImage(src, 0, 0);
                        }

                        double tolerancePercent = ToleranceSlider.Value / 100.0;
                        float limitSq = (float)(3 * 255 * 255 * (tolerancePercent * tolerancePercent)); // Cartesian distance limit^2
                        
                        Rectangle rect = new Rectangle(0, 0, dest.Width, dest.Height);
                        BitmapData bmpData = dest.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);

                        int bytes = Math.Abs(bmpData.Stride) * dest.Height;
                        byte[] rgbValues = new byte[bytes];
                        System.Runtime.InteropServices.Marshal.Copy(bmpData.Scan0, rgbValues, 0, bytes);

                        for (int i = 0; i < rgbValues.Length; i += 4)
                        {
                            byte b = rgbValues[i];
                            byte g = rgbValues[i+1];
                            byte r = rgbValues[i+2];
                            byte a = rgbValues[i+3];

                            if (a > 0)
                            {
                                float distSq = (r - _selectedColor.R)*(r - _selectedColor.R) +
                                               (g - _selectedColor.G)*(g - _selectedColor.G) +
                                               (b - _selectedColor.B)*(b - _selectedColor.B);

                                if (distSq <= limitSq)
                                {
                                    rgbValues[i + 3] = 0; // Make transparent
                                }
                            }
                        }

                        System.Runtime.InteropServices.Marshal.Copy(rgbValues, 0, bmpData.Scan0, bytes);
                        dest.UnlockBits(bmpData);

                        dest.Save(tempOut, ImageFormat.Png);
                    }
                }

                // Import replacement picture
                Slide slide = (Slide)targetShape.Parent;
                Shape newShape = slide.Shapes.AddPicture(tempOut, 
                    Microsoft.Office.Core.MsoTriState.msoFalse, 
                    Microsoft.Office.Core.MsoTriState.msoTrue, 
                    targetShape.Left, targetShape.Top, targetShape.Width, targetShape.Height);

                newShape.Rotation = targetShape.Rotation;
                
                // Try aligning z-order precisely
                while (newShape.ZOrderPosition > targetShape.ZOrderPosition + 1)
                {
                    newShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendBackward);
                }

                targetShape.Delete();
                newShape.Select();
                
                // Cleanup temp files safely without throwing if in use
                try { File.Delete(tempIn); } catch { }
                try { File.Delete(tempOut); } catch { }

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to erase color: " + ex.Message);
            }
        }
    }
}
