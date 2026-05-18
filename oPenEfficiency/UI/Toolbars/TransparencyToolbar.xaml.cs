using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using oPenEfficiency.Features.Visuals;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI.Toolbars
{
    public partial class TransparencyToolbar : Window
    {
        private readonly PowerPointManager _manager;
        private List<Shape> _targetShapes;
        private bool _isUpdatingFromPreset = false;

        public TransparencyToolbar(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            
            _targetShapes = new List<Shape>();
            float initialTransparency = 0f;
            bool foundTransparency = false;

            try 
            {
                var selection = manager.GetSelectedShapes();
                if (selection != null)
                {
                    for (int i = 1; i <= selection.Count; i++)
                    {
                        var shape = selection[i];
                        _targetShapes.Add(shape);

                        if (!foundTransparency)
                        {
                            try
                            {
                                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoLine || shape.Connector == Microsoft.Office.Core.MsoTriState.msoTrue)
                                {
                                    if (shape.Line != null) { initialTransparency = shape.Line.Transparency; foundTransparency = true; }
                                }
                                else if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || shape.Type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture)
                                {
                                    try { if (shape.PictureFormat != null) { initialTransparency = ((dynamic)shape.PictureFormat).Transparency; foundTransparency = true; } }
                                    catch { if (shape.Fill != null) { initialTransparency = shape.Fill.Transparency; foundTransparency = true; } }
                                }
                                else
                                {
                                    if (shape.Fill != null) { initialTransparency = shape.Fill.Transparency; foundTransparency = true; }
                                }
                            }
                            catch { }
                        }
                    }
                }
            } catch { }

            if (foundTransparency)
            {
                _isUpdatingFromPreset = true;
                double initialValue = Math.Round(initialTransparency * 100f);
                TransparencySlider.Value = initialValue;
                TransparencyValueText.Text = $"{initialValue}%";
                
                foreach (var child in PresetsPanel.Children)
                {
                    if (child is ToggleButton btn && btn.Tag != null)
                    {
                        if (double.TryParse(btn.Tag.ToString(), out double tagVal))
                        {
                            btn.IsChecked = Math.Abs(initialValue - tagVal) < 0.1;
                        }
                    }
                }
                
                _isUpdatingFromPreset = false;
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void BtnPreset_Click(object sender, RoutedEventArgs e)
        {
            if (sender is ToggleButton btn && btn.Tag != null)
            {
                if (float.TryParse(btn.Tag.ToString(), out float val))
                {
                    _isUpdatingFromPreset = true;
                    TransparencySlider.Value = val;
                    TransparencyValueText.Text = $"{val}%";
                    
                    float t = val / 100f;
                    TransparencyFeature.SetTransparency(_manager, _targetShapes, t);
                    
                    foreach (var child in PresetsPanel.Children)
                    {
                        if (child is ToggleButton tBtn)
                        {
                            tBtn.IsChecked = (tBtn == btn);
                        }
                    }
                    
                    _isUpdatingFromPreset = false;
                }
            }
        }

        private void TransparencySlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_isUpdatingFromPreset) return;

            float val = (float)e.NewValue;
            if (TransparencyValueText != null)
            {
                TransparencyValueText.Text = $"{Math.Round(val)}%";
            }

            float t = val / 100f;
            TransparencyFeature.SetTransparency(_manager, _targetShapes, t);

            if (PresetsPanel != null)
            {
                foreach (var child in PresetsPanel.Children)
                {
                    if (child is ToggleButton btn && btn.Tag != null)
                    {
                        if (double.TryParse(btn.Tag.ToString(), out double tagVal))
                        {
                            btn.IsChecked = Math.Abs(val - tagVal) < 0.1;
                        }
                    }
                }
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        protected override void OnDeactivated(EventArgs e)
        {
            base.OnDeactivated(e);
            try
            {
                this.Close();
            }
            catch { }
        }
    }
}
