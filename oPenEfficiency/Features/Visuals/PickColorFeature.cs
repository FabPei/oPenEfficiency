using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
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
        Tooltip = "Pick an absolute RGB color from the screen",
        IconData = "m17.58 3.08-2.6 2.6L18.33 9l2.6-2.6c.39-.39.39-1.02 0-1.41l-1.94-1.91c-.38-.39-1.02-.39-1.41 0zM13.56 6.1 3 16.66V21h4.34l10.56-10.56L13.56 6.1z",
        Color = "#06B6D4",
        Description = "Samples an exact screen pixel and applies its absolute RGB value to compatible properties of the selected shapes.",
        DetailedHelpText = "### Color Picker\n\nSelect one or more shapes, then click the picker and choose any visible screen pixel. The primary action changes fill foreground colors; right-click provides line and text targets. Fill styles such as gradients are preserved, text applies to the complete text of each compatible selected shape, Escape cancels without changes, and each successful multi-shape application is one PowerPoint Undo step.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
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
            if (manager == null || !HasCompatibleTarget(manager, target)) return false;

            Presentation sourcePresentation = manager.GetCurrentPresentation();
            if (sourcePresentation == null) return false;

            System.Drawing.Point? sampledPoint = null;
            Window overlay = null;
            try
            {
                overlay = CreateCaptureOverlay();
                overlay.PreviewMouseLeftButtonDown += (sender, args) =>
                {
                    sampledPoint = System.Windows.Forms.Control.MousePosition;
                    overlay.Close();
                };
                overlay.KeyDown += (sender, args) =>
                {
                    if (args.Key == Key.Escape)
                        overlay.Close();
                };

                overlay.ShowDialog();
                if (!sampledPoint.HasValue) return false;

                int sampledColor;
                using (var bitmap = new Bitmap(1, 1))
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    graphics.CopyFromScreen(
                        sampledPoint.Value,
                        new System.Drawing.Point(0, 0),
                        new System.Drawing.Size(1, 1));
                    var color = bitmap.GetPixel(0, 0);
                    sampledColor = color.R | (color.G << 8) | (color.B << 16);
                }

                var activePresentation = manager.GetCurrentPresentation();
                if (activePresentation == null || !IsSameComObject(activePresentation, sourcePresentation))
                {
                    MessageBox.Show(
                        "The active presentation changed while the color picker was open. No color was applied.",
                        "Color Picker",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return false;
                }

                return ApplyColorToSelection(manager, sampledColor, target, "PickColorFeature.Apply");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PickColorFeature.Execute");
                try { overlay?.Close(); } catch { /* Preserve the primary picker error if overlay cleanup also fails. */ }
                MessageBox.Show(
                    "The screen color could not be picked or applied.",
                    "Color Picker",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return false;
            }
        }

        internal static bool ApplyColorToSelection(
            PowerPointManager manager,
            int rgbColor,
            ApplyTarget target,
            string logContext)
        {
            if (manager == null) return false;
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return false;

            var states = new List<ColorTargetState>();
            try
            {
                for (int i = 1; i <= shapes.Count; i++)
                {
                    var state = ColorTargetState.TryCreate(shapes[i], target);
                    if (state != null) states.Add(state);
                }

                if (states.Count == 0) return false;

                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, logContext + ".StartNewUndoEntry"); }

                foreach (var state in states)
                    state.Apply(rgbColor);

                return true;
            }
            catch (Exception ex)
            {
                foreach (var state in states)
                {
                    try { state.Restore(); }
                    catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, logContext + ".Rollback"); }
                }

                ExceptionLogger.Log(ex, logContext);
                MessageBox.Show(
                    "The color could not be applied. Any partial changes were rolled back where possible.",
                    "Color Picker",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return false;
            }
        }

        private static bool HasCompatibleTarget(PowerPointManager manager, ApplyTarget target)
        {
            var shapes = manager.GetSelectedShapes();
            if (shapes == null) return false;

            for (int i = 1; i <= shapes.Count; i++)
            {
                try
                {
                    if (target != ApplyTarget.Text ||
                        shapes[i].HasTextFrame == Office.MsoTriState.msoTrue)
                        return true;
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "PickColorFeature.HasCompatibleTarget");
                }
            }

            return false;
        }

        private static Window CreateCaptureOverlay()
        {
            return new Window
            {
                Title = "Pick screen color (Esc to cancel)",
                WindowStyle = WindowStyle.None,
                AllowsTransparency = true,
                // VisualColorException: capture must leave the screen pixels visible.
                Background = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromArgb(1, 255, 255, 255)),
                Cursor = Cursors.Cross,
                Topmost = true,
                Left = SystemParameters.VirtualScreenLeft,
                Top = SystemParameters.VirtualScreenTop,
                Width = SystemParameters.VirtualScreenWidth,
                Height = SystemParameters.VirtualScreenHeight,
                ShowInTaskbar = false,
                ShowActivated = true
            };
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

        private sealed class ColorTargetState
        {
            private readonly Shape _shape;
            private readonly ApplyTarget _target;
            private readonly Office.MsoTriState _visible;
            private readonly ColorFormatState _color;

            private ColorTargetState(
                Shape shape,
                ApplyTarget target,
                Office.MsoTriState visible,
                ColorFormatState color)
            {
                _shape = shape;
                _target = target;
                _visible = visible;
                _color = color;
            }

            public static ColorTargetState TryCreate(Shape shape, ApplyTarget target)
            {
                if (shape == null) return null;

                try
                {
                    if (target == ApplyTarget.Fill)
                    {
                        return new ColorTargetState(
                            shape,
                            target,
                            shape.Fill.Visible,
                            new ColorFormatState(shape.Fill.ForeColor));
                    }

                    if (target == ApplyTarget.Line)
                    {
                        return new ColorTargetState(
                            shape,
                            target,
                            shape.Line.Visible,
                            new ColorFormatState(shape.Line.ForeColor));
                    }

                    if (shape.HasTextFrame != Office.MsoTriState.msoTrue) return null;
                    // No RTF probe: PowerPoint's TextRange object model exposes no RTF member
                    // (verified by PIA reflection - Text, Font, ParagraphFormat, Runs and friends
                    // only), so the former dynamic .Rtf read threw a RuntimeBinderException and
                    // logged an error entry on every text-target apply. Restore therefore always
                    // takes the font-color branch - which is the branch that always ran, because
                    // the captured value could never become non-empty.
                    return new ColorTargetState(
                        shape,
                        target,
                        Office.MsoTriState.msoTrue,
                        new ColorFormatState(shape.TextFrame.TextRange.Font.Color));
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "PickColorFeature.CaptureTarget");
                    return null;
                }
            }

            public void Apply(int rgbColor)
            {
                if (_target == ApplyTarget.Fill)
                {
                    _shape.Fill.ForeColor.RGB = rgbColor;
                    _shape.Fill.Visible = Office.MsoTriState.msoTrue;
                }
                else if (_target == ApplyTarget.Line)
                {
                    _shape.Line.ForeColor.RGB = rgbColor;
                    _shape.Line.Visible = Office.MsoTriState.msoTrue;
                }
                else
                {
                    _shape.TextFrame.TextRange.Font.Color.RGB = rgbColor;
                }
            }

            public void Restore()
            {
                if (_target == ApplyTarget.Fill)
                {
                    _color.Restore(_shape.Fill.ForeColor);
                    _shape.Fill.Visible = _visible;
                }
                else if (_target == ApplyTarget.Line)
                {
                    _color.Restore(_shape.Line.ForeColor);
                    _shape.Line.Visible = _visible;
                }
                else
                {
                    _color.Restore(_shape.TextFrame.TextRange.Font.Color);
                }
            }
        }

        private sealed class ColorFormatState
        {
            public ColorFormatState(ColorFormat color)
            {
                Rgb = color.RGB;
                ThemeColor = color.ObjectThemeColor;
                TintAndShade = color.TintAndShade;
            }

            private int Rgb { get; }
            private Office.MsoThemeColorIndex ThemeColor { get; }
            private float TintAndShade { get; }

            public void Restore(ColorFormat color)
            {
                if (ThemeColor != Office.MsoThemeColorIndex.msoNotThemeColor &&
                    ThemeColor != Office.MsoThemeColorIndex.msoThemeColorMixed)
                {
                    color.ObjectThemeColor = ThemeColor;
                    color.TintAndShade = TintAndShade;
                }
                else
                {
                    color.RGB = Rgb;
                }
            }
        }
    }
}
