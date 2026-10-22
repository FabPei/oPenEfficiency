using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    public partial class SizeAdjustControl : UserControl
    {
        private const double PointsPerCm = 28.346;
        private bool _isUpdatingFromCode;
        private PowerPointManager _manager;
        private PowerPoint.Application _application;

        public SizeAdjustControl()
        {
            InitializeComponent();
            Loaded += SizeAdjustControl_Loaded;
            Unloaded += SizeAdjustControl_Unloaded;
            TxtWidth.LostKeyboardFocus += (s, e) => ApplyWidthFromText();
            TxtHeight.LostKeyboardFocus += (s, e) => ApplyHeightFromText();
            TxtWidth.TextChanged += (s, e) => ApplyWidthFromText(immediate: true);
            TxtHeight.TextChanged += (s, e) => ApplyHeightFromText(immediate: true);
        }

        private void SizeAdjustControl_Loaded(object sender, RoutedEventArgs e)
        {
            _application = Globals.ThisAddIn?.Application;
            if (_application == null) return;

            _manager = new PowerPointManager(_application);
            try
            {
                _application.WindowSelectionChange += Application_WindowSelectionChange;
            }
            catch (Exception ex) { ReportOptionalSizeFailure(ex, "SizeAdjustControl.Loaded.SubscribeSelectionChange"); }
            RefreshFromSelection();
        }

        private void SizeAdjustControl_Unloaded(object sender, RoutedEventArgs e)
        {
            if (_application == null) return;

            try
            {
                _application.WindowSelectionChange -= Application_WindowSelectionChange;
            }
            catch (Exception ex) { ReportOptionalSizeFailure(ex, "SizeAdjustControl.Unloaded.UnsubscribeSelectionChange"); }
            finally
            {
                _application = null;
                _manager = null;
            }
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            RefreshFromSelection();
        }

        private void RefreshFromSelection()
        {
            if (_manager == null) return;
            try
            {
                var shapes = _manager.GetSelectedShapes();
                if (shapes != null && shapes.Count >= 1)
                {
                    var s = shapes[1];
                    _isUpdatingFromCode = true;
                    TxtWidth.Text = ToCmString(s.Width);
                    TxtHeight.Text = ToCmString(s.Height);
                }
                else
                {
                    _isUpdatingFromCode = true;
                    TxtWidth.Text = "";
                    TxtHeight.Text = "";
                }
            }
            catch (Exception ex) { ReportOptionalSizeFailure(ex, "SizeAdjustControl.RefreshFromSelection"); }
            finally
            {
                _isUpdatingFromCode = false;
            }
        }

        private string ToCmString(float points)
        {
            double cm = points / PointsPerCm;
            return cm.ToString("0.##", CultureInfo.InvariantCulture);
        }

        private float ToPoints(string cmText, float fallback)
        {
            if (double.TryParse(cmText, NumberStyles.Float, CultureInfo.InvariantCulture, out var cm))
            {
                return (float)(cm * PointsPerCm);
            }
            return fallback;
        }

        private void ApplyWidthFromText(bool immediate = false)
        {
            if (_isUpdatingFromCode) return;
            try
            {
                var shapes = _manager?.GetSelectedShapes();
                if (shapes == null || shapes.Count == 0) return;
                var s = shapes[1];
                float current = s.Width;
                float newPoints = ToPoints(TxtWidth.Text, current);
                if (Math.Abs(newPoints - current) > 0.01f)
                {
                    s.Width = newPoints;
                }
            }
            catch (Exception ex) { ReportOptionalSizeFailure(ex, "SizeAdjustControl.WidthChanged"); }
            finally
            {
                if (!immediate) RefreshFromSelection();
            }
        }

        private void ApplyHeightFromText(bool immediate = false)
        {
            if (_isUpdatingFromCode) return;
            try
            {
                var shapes = _manager?.GetSelectedShapes();
                if (shapes == null || shapes.Count == 0) return;
                var s = shapes[1];
                float current = s.Height;
                float newPoints = ToPoints(TxtHeight.Text, current);
                if (Math.Abs(newPoints - current) > 0.01f)
                {
                    s.Height = newPoints;
                }
            }
            catch (Exception ex) { ReportOptionalSizeFailure(ex, "SizeAdjustControl.HeightChanged"); }
            finally
            {
                if (!immediate) RefreshFromSelection();
            }
        }

        private void NudgeWidth(double deltaCm)
        {
            try
            {
                var shapes = _manager?.GetSelectedShapes();
                if (shapes == null || shapes.Count == 0) return;
                var s = shapes[1];
                s.Width = (float)(s.Width + deltaCm * PointsPerCm);
            }
            catch (Exception ex) { ReportOptionalSizeFailure(ex, "SizeAdjustControl.AdjustWidth"); }
            finally
            {
                RefreshFromSelection();
            }
        }

        private void NudgeHeight(double deltaCm)
        {
            try
            {
                var shapes = _manager?.GetSelectedShapes();
                if (shapes == null || shapes.Count == 0) return;
                var s = shapes[1];
                s.Height = (float)(s.Height + deltaCm * PointsPerCm);
            }
            catch (Exception ex) { ReportOptionalSizeFailure(ex, "SizeAdjustControl.AdjustHeight"); }
            finally
            {
                RefreshFromSelection();
            }
        }

        private static void ReportOptionalSizeFailure(Exception exception, string context)
        {
            if (!ExceptionPolicy.IsExpectedComFailure(exception))
                ExceptionLogger.Log(exception, context);
            // A stale selection leaves the control unchanged and available for the next selection event.
        }

        private void BtnWidthUp_Click(object sender, RoutedEventArgs e) => NudgeWidth(+0.1);
        private void BtnWidthDown_Click(object sender, RoutedEventArgs e) => NudgeWidth(-0.1);
        private void BtnHeightUp_Click(object sender, RoutedEventArgs e) => NudgeHeight(+0.1);
        private void BtnHeightDown_Click(object sender, RoutedEventArgs e) => NudgeHeight(-0.1);
    }
}
