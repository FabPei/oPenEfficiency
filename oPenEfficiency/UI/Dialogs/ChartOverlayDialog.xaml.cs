using System;
using System.Windows;
using Microsoft.Office.Core;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class ChartOverlayDialog : Window
    {
        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
                this.DragMove();
        }
        private PowerPointManager _manager;
        private Microsoft.Office.Interop.PowerPoint.Chart _chart;
        private Microsoft.Office.Interop.PowerPoint.Shape _chartShape;

        private int _pointCount = 0;

        public ChartOverlayDialog(PowerPointManager manager, Microsoft.Office.Interop.PowerPoint.Chart chart, Microsoft.Office.Interop.PowerPoint.Shape chartShape)
        {
            InitializeComponent();
            _manager = manager;
            _chart = chart;
            _chartShape = chartShape;

            LoadDataPoints();
        }

        private void LoadDataPoints()
        {
            try
            {
                // Attempt to read the first series to see how many points there are
                dynamic seriesCollection = _chart.SeriesCollection(1);
                dynamic values = seriesCollection.Values;
                if (values != null)
                {
                    Array valArray = (Array)values;
                    _pointCount = valArray.Length;

                    for (int i = 1; i <= _pointCount; i++)
                    {
                        StartPointCombo.Items.Add($"Point {i}");
                        EndPointCombo.Items.Add($"Point {i}");
                    }
                    if (_pointCount >= 2)
                    {
                        StartPointCombo.SelectedIndex = 0;
                        EndPointCombo.SelectedIndex = _pointCount - 1;
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ChartOverlayDialog.LoadDataPoints");
                System.Windows.Forms.MessageBox.Show("Could not read chart data automatically. Please ensure a standard data series exists.", "Warning", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        private void OverlayType_Checked(object sender, RoutedEventArgs e)
        {
            if (PointsGroupBox == null) return;
            
            bool isPointToPoint = OverlayDeltaRadio.IsChecked == true || 
                                  OverlayCagrRadio.IsChecked == true || 
                                  OverlayPercentRadio.IsChecked == true;
                                  
            PointsGroupBox.IsEnabled = isPointToPoint;
            if (BracketStylePanel != null) BracketStylePanel.IsEnabled = isPointToPoint;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int startIndex = StartPointCombo.SelectedIndex + 1; // 1-based indexing for COM
                int endIndex = EndPointCombo.SelectedIndex + 1;

                if (startIndex >= endIndex)
                {
                    System.Windows.Forms.MessageBox.Show("Start point must be before End point.", "Invalid Input", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

                ApplyOverlay(startIndex, endIndex);
                this.Close();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ChartOverlayDialog.ApplyButton_Click");
                System.Windows.Forms.MessageBox.Show($"Error applying overlay: {ex.Message}", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void ApplyOverlay(int sIndex, int eIndex)
        {
            dynamic seriesObject = _chart.SeriesCollection();
            int seriesCount = seriesObject.Count;
            dynamic series = seriesObject.Item(1);

            // Compute Stacked Values (sum of all series for the index)
            double startVal = 0;
            double endVal = 0;
            double[] allIndexSums = new double[_pointCount];

            for (int s = 1; s <= seriesCount; s++)
            {
                Array sVals = (Array)seriesObject.Item(s).Values;
                startVal += Convert.ToDouble(sVals.GetValue(sIndex));
                endVal += Convert.ToDouble(sVals.GetValue(eIndex));

                for (int i = 0; i < _pointCount; i++)
                {
                    allIndexSums[i] += Convert.ToDouble(sVals.GetValue(i + 1));
                }
            }

            var slide = (Microsoft.Office.Interop.PowerPoint.Slide)_chartShape.Parent;

            if (OverlayReferenceRadio.IsChecked == true)
            {
                // Average Level Line
                double sum = 0;
                foreach (double v in allIndexSums) sum += v;
                double mean = sum / _pointCount;

                DrawLevelLine(slide, mean, "Average: " + (InsertRoundedCheckBox.IsChecked == true ? Math.Round(mean, 1) : mean));
                return;
            }
            if (OverlayOutlierRadio.IsChecked == true)
            {
                // Outlier Detection (StdDev > 1.5)
                double sum = 0;
                foreach (double v in allIndexSums) sum += v;
                double mean = sum / _pointCount;
                double sumSq = 0;
                foreach (double v in allIndexSums) sumSq += Math.Pow(v - mean, 2);
                double stdDev = Math.Sqrt(sumSq / _pointCount);
                
                double thresholdUpper = mean + 1.2 * stdDev;
                double thresholdLower = mean - 1.2 * stdDev;

                bool foundOutlier = false;
                for (int i = 0; i < _pointCount; i++)
                {
                    if (allIndexSums[i] > thresholdUpper || allIndexSums[i] < thresholdLower)
                    {
                        DrawOutlierOval(slide, i + 1, series);
                        foundOutlier = true;
                    }
                }
                if (!foundOutlier) System.Windows.Forms.MessageBox.Show("No massive outliers detected.");
                return;
            }

            // Normal Point-to-Point Overlay
            string labelText = "";
            bool isPositive = (endVal >= startVal);

            if (OverlayDeltaRadio.IsChecked == true)
            {
                double delta = endVal - startVal;
                if (InsertRoundedCheckBox.IsChecked == true) delta = Math.Round(delta, 1);
                labelText = (delta > 0 ? "+" : "") + delta.ToString();
            }
            else if (OverlayCagrRadio.IsChecked == true)
            {
                double years = eIndex - sIndex;
                if (years == 0) years = 1;
                double cagr = (Math.Pow(endVal / startVal, 1.0 / years) - 1.0) * 100.0;
                if (InsertRoundedCheckBox.IsChecked == true) cagr = Math.Round(cagr, 1);
                labelText = "CAGR " + (cagr > 0 ? "+" : "") + cagr.ToString() + "%";
            }
            else 
            {
                double pct = ((endVal - startVal) / Math.Abs(startVal)) * 100.0;
                if (InsertRoundedCheckBox.IsChecked == true) pct = Math.Round(pct, 1);
                labelText = (pct > 0 ? "+" : "") + pct.ToString() + "%";
            }

            DrawBracketOverlay(slide, series, sIndex, eIndex, labelText, isPositive);
        }

        private void DrawBracketOverlay(Microsoft.Office.Interop.PowerPoint.Slide slide, dynamic series, int sIndex, int eIndex, string labelText, bool isPositive)
        {
            dynamic startPoint = series.Points(sIndex);
            dynamic endPoint = series.Points(eIndex);

            float x1, y1, x2, y2;
            try 
            {
                x1 = _chartShape.Left + (float)startPoint.Left + ((float)startPoint.Width / 2);
                y1 = _chartShape.Top + (float)startPoint.Top - 15f; 
                x2 = _chartShape.Left + (float)endPoint.Left + ((float)endPoint.Width / 2);
                y2 = _chartShape.Top + (float)endPoint.Top - 15f;
            } 
            catch 
            {
                float widthStep = _chartShape.Width / _pointCount;
                x1 = _chartShape.Left + (sIndex - 0.5f) * widthStep;
                x2 = _chartShape.Left + (eIndex - 0.5f) * widthStep;
                y1 = _chartShape.Top + 20f;
                y2 = _chartShape.Top + 20f;
            }

            int bracketStyle = BracketStyleCombo.SelectedIndex; // 0=Arrows, 1=Square, 2=Curved

            var line = slide.Shapes.AddLine(x1, y1, x2, y1); 
            line.Line.Weight = 1.5f;

            Microsoft.Office.Interop.PowerPoint.Shape leg1 = null;
            Microsoft.Office.Interop.PowerPoint.Shape leg2 = null;

            if (bracketStyle == 0 || bracketStyle == 1)
            {
                leg1 = slide.Shapes.AddLine(x1, y1, x1, y1 + 10f); 
                leg2 = slide.Shapes.AddLine(x2, y1, x2, y2); 
                if (bracketStyle == 0) // Arrow
                {
                    line.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                    leg2.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                }
            }
            if (bracketStyle == 2)
            {
                // Curved (Simulated via smoothing)
                line.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
            }
            
            var txt = slide.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, x1 + ((x2 - x1) / 2) - 30f, y1 - 20f, 60f, 20f);
            txt.TextFrame.TextRange.Text = labelText;
            txt.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            txt.TextFrame.TextRange.Font.Size = 12;
            txt.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;

            if (ColorCodeCheckBox.IsChecked == true)
            {
                int color = isPositive ? System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.MediumSeaGreen) : System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.IndianRed);
                line.Line.ForeColor.RGB = color;
                if (leg1 != null) { leg1.Line.ForeColor.RGB = color; leg2.Line.ForeColor.RGB = color; }
                txt.TextFrame.TextRange.Font.Color.RGB = color;
            }

            string[] sNames;
            if (leg1 == null) sNames = new string[] { line.Name, txt.Name };
            else sNames = new string[] { line.Name, leg1.Name, leg2.Name, txt.Name };

            slide.Shapes.Range(sNames).Group().Tags.Add("ChartOverlayTracker", "Bracket");
        }

        private void DrawLevelLine(Microsoft.Office.Interop.PowerPoint.Slide slide, double meanValue, string text)
        {
            try
            {
                dynamic ax = _chart.Axes(Microsoft.Office.Interop.PowerPoint.XlAxisType.xlValue);
                var pa = _chart.PlotArea;
                double min = ax.MinimumScale;
                double max = ax.MaximumScale;
                
                double ratio = (meanValue - min) / (max - min);
                float yPx = _chartShape.Top + (float)pa.Top + (float)pa.Height - (float)(pa.Height * ratio);

                float x1 = _chartShape.Left + (float)pa.Left;
                float x2 = x1 + (float)pa.Width;

                var line = slide.Shapes.AddLine(x1, yPx, x2, yPx);
                line.Line.DashStyle = MsoLineDashStyle.msoLineDash;
                line.Line.Weight = 2f;
                line.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);

                var txt = slide.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, x2 + 10f, yPx - 10f, 100f, 20f);
                txt.TextFrame.TextRange.Text = text;
                txt.TextFrame.TextRange.Font.Size = 10;
                txt.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                
                slide.Shapes.Range(new string[] { line.Name, txt.Name }).Group().Tags.Add("ChartOverlayTracker", "Reference");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DrawLevelLine");
                System.Windows.Forms.MessageBox.Show("Could not calculate exact chart axis geometry. Is this a standard numeric chart?", "Warning");
            }
        }

        private void DrawOutlierOval(Microsoft.Office.Interop.PowerPoint.Slide slide, int COMIndex, dynamic series)
        {
            try
            {
                dynamic p = series.Points(COMIndex);
                float x = _chartShape.Left + (float)p.Left + ((float)p.Width / 2);
                float y = _chartShape.Top + (float)p.Top;

                var oval = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, x - 20f, y - 20f, 40f, 40f);
                oval.Fill.Visible = MsoTriState.msoFalse;
                oval.Line.Weight = 2.5f;
                oval.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                oval.Line.DashStyle = MsoLineDashStyle.msoLineDash;
                oval.Tags.Add("ChartOverlayTracker", "Outlier");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DrawOutlierOval");
            }
        }
    }
}
