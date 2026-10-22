using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    /// <summary>
    /// Compact Width / Height control embedded in the sidebar SIZE section.
    /// Reads and writes the selected PowerPoint shape's dimensions in centimetres.
    /// 1 cm = 28.3464567 PowerPoint points.
    /// </summary>
    public partial class ShapeSizePanel : UserControl
    {
        private const double StepCm   = 0.1;

        private PowerPointManager _ppManager;
        private bool _isUpdating = false;
        private string _displayedWidthText;
        private string _displayedHeightText;

        public ShapeSizePanel(PowerPointManager ppManager)
        {
            InitializeComponent();
            _ppManager = ppManager;
        }

        // ─── Public API ───────────────────────────────────────────────────────────

        /// <summary>
        /// Call this whenever the PowerPoint selection changes.
        /// </summary>
        public void OnSelectionChange(Selection sel)
        {
            try
            {
                UpdateDisplay(GetShapesFromSelection(sel));
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ShapeSizePanel.OnSelectionChange");
                ClearDisplay();
            }
        }

        /// <summary>
        /// Update the display with the given shape's dimensions (called while editing text).
        /// </summary>
        public void UpdateSize(Shape shape)
        {
            try
            {
                if (shape == null) return;
                UpdateDisplay(new List<Shape> { shape });
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ShapeSizePanel.UpdateSize"); }
        }

        public void FocusWidth()
        {
            FocusDimension(TxtWidth);
        }

        public void FocusHeight()
        {
            FocusDimension(TxtHeight);
        }

        private void FocusDimension(TextBox textBox)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                textBox.Focus();
                Keyboard.Focus(textBox);
                textBox.SelectAll();
            }), System.Windows.Threading.DispatcherPriority.Input);
        }

        // ─── Spin buttons ─────────────────────────────────────────────────────────

        private void BtnWidthUp_Click(object sender, RoutedEventArgs e)  => AdjustWidth(+StepCm);
        private void BtnWidthDown_Click(object sender, RoutedEventArgs e) => AdjustWidth(-StepCm);
        private void BtnHeightUp_Click(object sender, RoutedEventArgs e)  => AdjustHeight(+StepCm);
        private void BtnHeightDown_Click(object sender, RoutedEventArgs e) => AdjustHeight(-StepCm);

        private void AdjustWidth(double deltaCm)
        {
            AdjustDimension(deltaCm, isWidth: true);
        }

        private void AdjustHeight(double deltaCm)
        {
            AdjustDimension(deltaCm, isWidth: false);
        }

        // ─── Text box keyboard / focus handlers ───────────────────────────────────

        private void TxtWidth_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) { ApplyFromTextBox(TxtWidth, isWidth: true); e.Handled = true; }
            if (e.Key == Key.Escape) { RefreshFromSelection(); e.Handled = true; }
        }

        private void TxtHeight_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) { ApplyFromTextBox(TxtHeight, isWidth: false); e.Handled = true; }
            if (e.Key == Key.Escape) { RefreshFromSelection(); e.Handled = true; }
        }

        private void TxtWidth_LostFocus(object sender, RoutedEventArgs e)  => ApplyFromTextBox(TxtWidth,  isWidth: true);
        private void TxtHeight_LostFocus(object sender, RoutedEventArgs e) => ApplyFromTextBox(TxtHeight, isWidth: false);

        // ─── Core apply logic ─────────────────────────────────────────────────────

        private void ApplyFromTextBox(TextBox box, bool isWidth)
        {
            if (_isUpdating) return;
            string displayedText = isWidth ? _displayedWidthText : _displayedHeightText;
            if (string.Equals(box.Text, displayedText, StringComparison.Ordinal)) return;
            if (!TryParseDisplay(box.Text, out double cm) || !Core.Services.ShapeSizePolicy.IsValidTypedCm(cm)) { RefreshFromSelection(); return; }
            if (isWidth) ApplyWidth(cm); else ApplyHeight(cm);
        }

        private void ApplyWidth(double cm)
        {
            ApplyDimension(cm, isWidth: true);
        }

        private void ApplyHeight(double cm)
        {
            ApplyDimension(cm, isWidth: false);
        }

        private void ApplyDimension(double cm, bool isWidth)
        {
            try
            {
                var sel = _ppManager?.GetApplication()?.ActiveWindow?.Selection;
                if (sel == null) return;

                var shapes = GetShapesFromSelection(sel);
                float targetPoints = (float)CmToPts(cm);
                ApplyDimensionChanges(shapes, isWidth, shape => targetPoints);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, isWidth ? "ShapeSizePanel.ApplyWidth" : "ShapeSizePanel.ApplyHeight");
                RefreshFromSelection();
            }
        }

        private void AdjustDimension(double deltaCm, bool isWidth)
        {
            try
            {
                var sel = _ppManager?.GetApplication()?.ActiveWindow?.Selection;
                if (sel == null) return;
                var shapes = GetShapesFromSelection(sel);
                ApplyDimensionChanges(
                    shapes,
                    isWidth,
                    shape => (float)Core.Services.ShapeSizePolicy.TargetAfterAdjust(isWidth ? shape.Width : shape.Height, deltaCm));
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, isWidth ? "ShapeSizePanel.AdjustWidth" : "ShapeSizePanel.AdjustHeight");
                RefreshFromSelection();
            }
        }

        private void ApplyDimensionChanges(IList<Shape> shapes, bool isWidth, Func<Shape, float> getTarget)
        {
            if (shapes == null || shapes.Count == 0) return;

            var states = new List<DimensionState>();
            foreach (var shape in shapes)
            {
                float target = getTarget(shape);
                float current = isWidth ? shape.Width : shape.Height;
                if (!Core.Services.ShapeSizePolicy.NeedsDimensionChange(current, target)) continue;
                states.Add(new DimensionState(shape, target));
            }
            if (states.Count == 0)
            {
                UpdateDisplay(shapes);
                return;
            }

            try
            {
                try { _ppManager?.GetApplication()?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ShapeSizePanel.StartNewUndoEntry"); }

                foreach (var state in states)
                {
                    if (isWidth) state.Shape.Width = state.Target;
                    else state.Shape.Height = state.Target;
                }
                UpdateDisplay(shapes);
            }
            catch
            {
                foreach (var state in states)
                {
                    try
                    {
                        state.Shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        state.Shape.Width = state.Width;
                        state.Shape.Height = state.Height;
                        state.Shape.Left = state.Left;
                        state.Shape.Top = state.Top;
                        state.Shape.LockAspectRatio = state.LockAspectRatio;
                    }
                    catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "ShapeSizePanel.Rollback"); }
                }
                throw;
            }
        }

        private static List<Shape> GetShapesFromSelection(Selection sel)
        {
            var shapes = new List<Shape>();
            if (sel == null) return shapes;

            if (sel.Type == PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
            {
                for (int i = 1; i <= sel.ShapeRange.Count; i++) shapes.Add(sel.ShapeRange[i]);
            }
            else if (sel.Type == PpSelectionType.ppSelectionText)
            {
                var textRange = sel.TextRange;
                if (textRange != null && textRange.Parent is Shape parentShape) shapes.Add(parentShape);
            }
            return shapes;
        }

        // ─── Display helpers ──────────────────────────────────────────────────────

        private void UpdateDisplay(IList<Shape> shapes)
        {
            if (shapes == null || shapes.Count == 0)
            {
                ClearDisplay();
                return;
            }

            double? widthCm = PtsToCm(shapes[0].Width);
            double? heightCm = PtsToCm(shapes[0].Height);
            for (int i = 1; i < shapes.Count; i++)
            {
                if (!Core.Services.ShapeSizePolicy.AreUniform(shapes[0].Width, shapes[i].Width)) widthCm = null;
                if (!Core.Services.ShapeSizePolicy.AreUniform(shapes[0].Height, shapes[i].Height)) heightCm = null;
            }
            SetDisplay(widthCm, heightCm);
        }

        private void SetDisplay(double? widthCm, double? heightCm)
        {
            try
            {
                _isUpdating = true;
                _displayedWidthText = widthCm.HasValue ? FormatCm(widthCm.Value) : "Mixed";
                _displayedHeightText = heightCm.HasValue ? FormatCm(heightCm.Value) : "Mixed";
                TxtWidth.Text = _displayedWidthText;
                TxtHeight.Text = _displayedHeightText;
            }
            finally { _isUpdating = false; }
        }

        private void ClearDisplay()
        {
            try
            {
                _isUpdating = true;
                _displayedWidthText = "—";
                _displayedHeightText = "—";
                TxtWidth.Text = _displayedWidthText;
                TxtHeight.Text = _displayedHeightText;
            }
            finally { _isUpdating = false; }
        }

        private void RefreshFromSelection()
        {
            try
            {
                var sel = _ppManager?.GetApplication()?.ActiveWindow?.Selection;
                if (sel == null)
                {
                    ClearDisplay();
                    return;
                }

                UpdateDisplay(GetShapesFromSelection(sel));
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ShapeSizePanel.RefreshFromSelection");
                ClearDisplay();
            }
        }

        // ─── Utility ──────────────────────────────────────────────────────────────

        private static double PtsToCm(double pts) => Core.Services.ShapeSizePolicy.PointsToCm(pts);
        private static double CmToPts(double cm)  => Core.Services.ShapeSizePolicy.CmToPoints(cm);

        private static string FormatCm(double cm) => Core.Services.ShapeSizePolicy.FormatCm(cm);

        private static bool TryParseDisplay(string text, out double cm)
            => Core.Services.ShapeSizePolicy.TryParseDisplay(text, out cm);

        private sealed class DimensionState
        {
            public Shape Shape { get; }
            public float Target { get; }
            public float Left { get; }
            public float Top { get; }
            public float Width { get; }
            public float Height { get; }
            public Microsoft.Office.Core.MsoTriState LockAspectRatio { get; }

            public DimensionState(Shape shape, float target)
            {
                Shape = shape;
                Target = target;
                Left = shape.Left;
                Top = shape.Top;
                Width = shape.Width;
                Height = shape.Height;
                LockAspectRatio = shape.LockAspectRatio;
            }
        }
    }
}
