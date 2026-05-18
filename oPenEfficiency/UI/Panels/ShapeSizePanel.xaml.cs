using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI
{
    /// <summary>
    /// Compact Width / Height control embedded in the sidebar SIZE section.
    /// Reads and writes the selected PowerPoint shape's dimensions in centimetres.
    /// 1 cm = 28.3464567 PowerPoint points.
    /// </summary>
    public partial class ShapeSizePanel : UserControl
    {
        private const double PtsPerCm = 28.3464567;
        private const double StepCm   = 0.1;

        private PowerPointManager _ppManager;
        private bool _isUpdating = false;

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
                if (sel == null)
                {
                    ClearDisplay();
                    return;
                }

                float widthPts = 0;
                float heightPts = 0;
                bool hasValidSelection = false;

                if (sel.Type == PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
                {
                    // For multi-selection, show the bounding box (width/height of the range)
                    widthPts = sel.ShapeRange.Width;
                    heightPts = sel.ShapeRange.Height;
                    hasValidSelection = true;
                }
                else if (sel.Type == PpSelectionType.ppSelectionText)
                {
                    // Text is selected inside a shape - show the parent shape's dimensions
                    var textRange = sel.TextRange;
                    if (textRange != null && textRange.Parent is Shape parentShape)
                    {
                        widthPts = parentShape.Width;
                        heightPts = parentShape.Height;
                        hasValidSelection = true;
                    }
                }

                if (hasValidSelection)
                {
                    SetDisplay(PtsToCm(widthPts), PtsToCm(heightPts));
                }
                else
                {
                    ClearDisplay();
                }
            }
            catch
            {
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
                SetDisplay(PtsToCm(shape.Width), PtsToCm(shape.Height));
            }
            catch { }
        }

        // ─── Spin buttons ─────────────────────────────────────────────────────────

        private void BtnWidthUp_Click(object sender, RoutedEventArgs e)  => AdjustWidth(+StepCm);
        private void BtnWidthDown_Click(object sender, RoutedEventArgs e) => AdjustWidth(-StepCm);
        private void BtnHeightUp_Click(object sender, RoutedEventArgs e)  => AdjustHeight(+StepCm);
        private void BtnHeightDown_Click(object sender, RoutedEventArgs e) => AdjustHeight(-StepCm);

        private void AdjustWidth(double deltaCm)
        {
            if (!TryParseDisplay(TxtWidth.Text, out double current)) return;
            ApplyWidth(Math.Max(0.01, current + deltaCm));
        }

        private void AdjustHeight(double deltaCm)
        {
            if (!TryParseDisplay(TxtHeight.Text, out double current)) return;
            ApplyHeight(Math.Max(0.01, current + deltaCm));
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
            if (!TryParseDisplay(box.Text, out double cm) || cm <= 0) { RefreshFromSelection(); return; }
            if (isWidth) ApplyWidth(cm); else ApplyHeight(cm);
        }

        private void ApplyWidth(double cm)
        {
            try
            {
                var sel = _ppManager?.GetApplication()?.ActiveWindow?.Selection;
                if (sel == null) return;

                var shape = GetShapeFromSelection(sel);
                if (shape == null) return;

                shape.Width = (float)CmToPts(cm);
                SetDisplay(cm, PtsToCm(shape.Height));
            }
            catch { RefreshFromSelection(); }
        }

        private void ApplyHeight(double cm)
        {
            try
            {
                var sel = _ppManager?.GetApplication()?.ActiveWindow?.Selection;
                if (sel == null) return;

                var shape = GetShapeFromSelection(sel);
                if (shape == null) return;

                shape.Height = (float)CmToPts(cm);
                SetDisplay(PtsToCm(shape.Width), cm);
            }
            catch { RefreshFromSelection(); }
        }

        /// <summary>
        /// Gets the shape from selection, handling both shape selection and text range selection.
        /// </summary>
        private Shape GetShapeFromSelection(Selection sel)
        {
            if (sel.Type == PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
            {
                // For single shape selection, return the first shape
                // ShapeRange[1] returns a Shape object
                return sel.ShapeRange[1];
            }
            else if (sel.Type == PpSelectionType.ppSelectionText)
            {
                var textRange = sel.TextRange;
                if (textRange != null && textRange.Parent is Shape parentShape)
                {
                    return parentShape;
                }
            }
            return null;
        }

        // ─── Display helpers ──────────────────────────────────────────────────────

        private void SetDisplay(double widthCm, double heightCm)
        {
            _isUpdating = true;
            TxtWidth.Text  = FormatCm(widthCm);
            TxtHeight.Text = FormatCm(heightCm);
            _isUpdating = false;
        }

        private void ClearDisplay()
        {
            _isUpdating = true;
            TxtWidth.Text  = "—";
            TxtHeight.Text = "—";
            _isUpdating = false;
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

                var shape = GetShapeFromSelection(sel);
                if (shape != null)
                {
                    SetDisplay(PtsToCm(shape.Width), PtsToCm(shape.Height));
                }
                else
                {
                    ClearDisplay();
                }
            }
            catch { ClearDisplay(); }
        }

        // ─── Utility ──────────────────────────────────────────────────────────────

        private static double PtsToCm(double pts) => pts / PtsPerCm;
        private static double CmToPts(double cm)  => cm  * PtsPerCm;

        private static string FormatCm(double cm)
            => cm.ToString("0.##", CultureInfo.CurrentCulture) + " cm";

        private static bool TryParseDisplay(string text, out double cm)
        {
            // Strip unit and whitespace, handle both "." and "," decimal separators
            text = (text ?? "").Replace("cm", "").Trim();
            return double.TryParse(text,
                NumberStyles.Any,
                CultureInfo.CurrentCulture,
                out cm)
                || double.TryParse(text,
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out cm);
        }
    }
}
