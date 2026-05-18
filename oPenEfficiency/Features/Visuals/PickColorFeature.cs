using System;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnPickColor",
        Name = "Color Picker",
        Tooltip = "Pick color from screen",
        IconData = "m17.58 3.08-2.6 2.6L18.33 9l2.6-2.6c.39-.39.39-1.02 0-1.41l-1.94-1.91c-.38-.39-1.02-.39-1.41 0zM13.56 6.1 3 16.66V21h4.34l10.56-10.56L13.56 6.1z",
        Color = "#06B6D4",
        Description = "Pick a color from anywhere on your screen and apply it to the Fill or Font of selected objects.",
        DetailedHelpText = "### Color Picker\nActivates a screen color picker eyedropper. Click any pixel on screen to sample its RGB color and apply it as the fill of the selected shapes.",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class PickColorFeature
    {
        public enum ApplyTarget
        {
            Fill,
            Line,
            Text
        }

        public static bool Execute(PowerPointManager manager, ApplyTarget target = ApplyTarget.Fill)
        {
            if (manager == null) return false;

            // Create a transparent overlay window to capture the click
            Window overlay = new Window
            {
                WindowStyle = WindowStyle.None,
                AllowsTransparency = true,
                // CRITICAL: Pure Transparent is click-through. 
                // We use an almost invisible color to capture the mouse.
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(1, 255, 255, 255)),
                Cursor = System.Windows.Input.Cursors.Cross,
                Topmost = true,
                Left = System.Windows.SystemParameters.VirtualScreenLeft,
                Top = System.Windows.SystemParameters.VirtualScreenTop,
                Width = System.Windows.SystemParameters.VirtualScreenWidth,
                Height = System.Windows.SystemParameters.VirtualScreenHeight,
                ShowInTaskbar = false,
                ShowActivated = true
            };

            bool success = false;

            overlay.MouseDown += (s, e) =>
            {
                try
                {
                    var point = System.Windows.Forms.Control.MousePosition;
                    bool colorApplied = false;

                    // 1. Try to get the exact color from a PowerPoint shape using RangeFromPoint
                    try
                    {
                        var activeWin = manager.GetApplication().ActiveWindow;
                        if (activeWin != null)
                        {
                            // RangeFromPoint expects screen coordinates in pixels
                            object hitObject = activeWin.RangeFromPoint(point.X, point.Y);
                            if (hitObject is Microsoft.Office.Interop.PowerPoint.Shape hitShape)
                            {
                                int rgb = 0;
                                bool found = false;

                                if (target == ApplyTarget.Fill && hitShape.Fill.Visible == Office.MsoTriState.msoTrue)
                                {
                                    rgb = hitShape.Fill.ForeColor.RGB;
                                    found = true;
                                }
                                else if (target == ApplyTarget.Line && hitShape.Line.Visible == Office.MsoTriState.msoTrue)
                                {
                                    rgb = hitShape.Line.ForeColor.RGB;
                                    found = true;
                                }
                                else if (target == ApplyTarget.Text && hitShape.HasTextFrame == Office.MsoTriState.msoTrue)
                                {
                                    rgb = hitShape.TextFrame.TextRange.Font.Color.RGB;
                                    found = true;
                                }

                                if (found)
                                {
                                    ApplyColorToSelection(manager, rgb, target);
                                    colorApplied = true;
                                    success = true;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "PickColorFeature.HitTest");
                    }

                    // 2. Fallback: Pixel-perfect screen color picking
                    if (!colorApplied)
                    {
                        using (Bitmap bmp = new Bitmap(1, 1))
                        {
                            using (Graphics g = Graphics.FromImage(bmp))
                            {
                                g.CopyFromScreen(point, new System.Drawing.Point(0, 0), new System.Drawing.Size(1, 1));
                            }
                            System.Drawing.Color color = bmp.GetPixel(0, 0);
                            int pptColor = color.R | (color.G << 8) | (color.B << 16);
                            ApplyColorToSelection(manager, pptColor, target);
                            success = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "PickColorFeature.Execute");
                }
                finally
                {
                    overlay.Close();
                }
            };

            // Allow ESC to cancel
            overlay.KeyDown += (s, e) =>
            {
                if (e.Key == System.Windows.Input.Key.Escape)
                {
                    overlay.Close();
                }
            };

            overlay.ShowDialog();
            return success;
        }

        private static bool ApplyColorToSelection(PowerPointManager manager, int rgbColor, ApplyTarget target)
        {
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count == 0) return false;

            try
            {
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    try
                    {
                        if (target == ApplyTarget.Fill && shape.Fill != null)
                        {
                            shape.Fill.ForeColor.RGB = rgbColor;
                            shape.Fill.Visible = Office.MsoTriState.msoTrue;
                            shape.Fill.Solid();
                        }
                        else if (target == ApplyTarget.Line && shape.Line != null)
                        {
                            shape.Line.ForeColor.RGB = rgbColor;
                            shape.Line.Visible = Office.MsoTriState.msoTrue;
                        }
                        else if (target == ApplyTarget.Text && shape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            shape.TextFrame.TextRange.Font.Color.RGB = rgbColor;
                        }
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, $"PickColorFeature.Apply.{target}");
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PickColorFeature.ApplyColorToSelection");
                return false;
            }
        }
    }
}
