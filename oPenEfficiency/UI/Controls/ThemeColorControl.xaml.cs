using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public partial class ThemeColorControl : UserControl
    {
        private PowerPointManager _manager;
        private static Color[] _savedCustomColors = new Color[20];

        public event Action<int> ColorSelected;

        public ThemeColorControl()
        {
            InitializeComponent();
        }

        public bool IsVerticalLayout { get; set; } = false;

        public void Initialize(PowerPointManager manager, bool vertical = false)
        {
            _manager = manager;
            IsVerticalLayout = vertical;
            
            StandardLayout.Visibility = IsVerticalLayout ? Visibility.Collapsed : Visibility.Visible;
            VerticalLayout.Visibility = IsVerticalLayout ? Visibility.Visible : Visibility.Collapsed;
            
            ApplySettings();
            LoadThemeColors();
        }

        private void ApplySettings()
        {
            try
            {
                var config = oPenEfficiency.JsonService.Load<oPenEfficiency.Models.SidebarConfig>("sidebar_config.json");
                if (config == null) return;

                if (!string.IsNullOrEmpty(config.BackgroundColor))
                {
                    try { this.Background = (Brush)new BrushConverter().ConvertFromString(config.BackgroundColor); } catch { }
                }

                Brush txtBrush = null;
                if (!string.IsNullOrEmpty(config.SectionFontColor))
                {
                    try { txtBrush = (Brush)new BrushConverter().ConvertFromString(config.SectionFontColor); } catch { }
                }

                double fontSize = config.SectionFontSize > 0 ? config.SectionFontSize : 9;

                void ApplyTextBlockStyle(Panel panel)
                {
                    if (panel == null) return;
                    foreach (var child in panel.Children)
                    {
                        if (child is TextBlock tb)
                        {
                            if (txtBrush != null) tb.Foreground = txtBrush;
                            tb.FontSize = fontSize;
                        }
                        else if (child is Panel pnl)
                        {
                            ApplyTextBlockStyle(pnl);
                        }
                    }
                }

                ApplyTextBlockStyle(StandardLayout);
                ApplyTextBlockStyle(VerticalLayout);
            }
            catch { }
        }

        private void LoadThemeColors()
        {
            try
            {
                var app = _manager.GetApplication();
                if (app.ActivePresentation == null) return;

                var theme = app.ActivePresentation.SlideMaster.Theme;
                var scheme = theme.ThemeColorScheme;

                Office.MsoThemeColorSchemeIndex[] indices = {
                    Office.MsoThemeColorSchemeIndex.msoThemeDark1,
                    Office.MsoThemeColorSchemeIndex.msoThemeLight1,
                    Office.MsoThemeColorSchemeIndex.msoThemeDark2,
                    Office.MsoThemeColorSchemeIndex.msoThemeLight2,
                    Office.MsoThemeColorSchemeIndex.msoThemeAccent1,
                    Office.MsoThemeColorSchemeIndex.msoThemeAccent2,
                    Office.MsoThemeColorSchemeIndex.msoThemeAccent3,
                    Office.MsoThemeColorSchemeIndex.msoThemeAccent4,
                    Office.MsoThemeColorSchemeIndex.msoThemeAccent5,
                    Office.MsoThemeColorSchemeIndex.msoThemeAccent6
                };

                float[] tints = { 0f, 0.8f, 0.6f, 0.4f, -0.25f, -0.5f };

                GridFill.Children.Clear();
                GridBorder.Children.Clear();
                GridText.Children.Clear();
                GridVertical.Children.Clear();
                GridCustom1.Children.Clear();
                GridCustom2.Children.Clear();
                GridCustomVertical.Children.Clear();

                if (IsVerticalLayout)
                {
                    // Vertical logic: Group by color variation
                    foreach (var idx in indices)
                    {
                        foreach (var t in tints)
                        {
                            int rgb = scheme.Colors(idx).RGB;
                            Color baseColor = ConvertRgb(rgb);
                            Color color = ApplyTint(baseColor, t);

                            GridVertical.Children.Add(CreateColorButton(color, "Fill", idx, t));
                            GridVertical.Children.Add(CreateColorButton(color, "Border", idx, t));
                            GridVertical.Children.Add(CreateColorButton(color, "Text", idx, t));
                        }
                    }
                }
                else
                {
                    // Standard logic: Group by type
                    foreach (var t in tints)
                    {
                        foreach (var idx in indices)
                        {
                            int rgb = scheme.Colors(idx).RGB;
                            Color baseColor = ConvertRgb(rgb);
                            Color color = ApplyTint(baseColor, t);

                            GridFill.Children.Add(CreateColorButton(color, "Fill", idx, t));
                            GridBorder.Children.Add(CreateColorButton(color, "Border", idx, t));
                            GridText.Children.Add(CreateColorButton(color, "Text", idx, t));
                        }
                    }
                }

                for (int i = 0; i < 20; i++)
                {
                    var btn = CreateCustomButton(i);
                    if (IsVerticalLayout)
                    {
                        GridCustomVertical.Children.Add(btn);
                    }
                    else
                    {
                        if (i < 10) GridCustom1.Children.Add(btn);
                        else GridCustom2.Children.Add(btn);
                    }
                }
            }
            catch { }
        }

        private Color ApplyTint(Color color, float factor)
        {
            if (factor == 0) return color;
            if (factor > 0)
            {
                return Color.FromRgb(
                    (byte)(color.R + (255 - color.R) * factor),
                    (byte)(color.G + (255 - color.G) * factor),
                    (byte)(color.B + (255 - color.B) * factor)
                );
            }
            else
            {
                float f = 1 + factor;
                return Color.FromRgb(
                    (byte)(color.R * f),
                    (byte)(color.G * f),
                    (byte)(color.B * f)
                );
            }
        }

        private Button CreateColorButton(Color color, string type, Office.MsoThemeColorSchemeIndex idx, float tint)
        {
            Button btn = new Button
            {
                Width = 22, Height = 22, Margin = new Thickness(0.5),
                Cursor = Cursors.Hand, Background = Brushes.Transparent, BorderThickness = new Thickness(0),
                Padding = new Thickness(0), ToolTip = $"{type} - {idx} (Tint: {tint})"
            };

            Border border = new Border
            {
                Width = 18, Height = 18, CornerRadius = new CornerRadius(3),
                HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center
            };

            if (type == "Fill")
            {
                border.Background = new SolidColorBrush(color);
                border.BorderBrush = new SolidColorBrush(Color.FromRgb(80, 80, 80));
                border.BorderThickness = new Thickness(0.5);
            }
            else if (type == "Border")
            {
                border.BorderBrush = new SolidColorBrush(color);
                border.BorderThickness = new Thickness(2);
                border.Background = Brushes.Transparent;
            }
            else if (type == "Text")
            {
                border.Background = Brushes.Transparent;
                TextBlock tb = new TextBlock
                {
                    Text = "T", Foreground = new SolidColorBrush(color),
                    FontSize = 12, FontWeight = FontWeights.Bold,
                    HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center
                };
                border.Child = tb;
            }

            btn.Content = border;
            btn.Click += (s, e) => ApplyColor(idx, type, tint);
            return btn;
        }

        private void ApplyColor(Office.MsoThemeColorSchemeIndex idx, string type, float tint)
        {
            try
            {
                var app = _manager.GetApplication();
                var scheme = app.ActivePresentation.SlideMaster.Theme.ThemeColorScheme;
                int rgb = scheme.Colors(idx).RGB;
                
                // If we have a tint, we should probably resolve it to an absolute RGB for the "Replace" window
                // But PowerPoint's RGB property for theme colors is handled automatically if we set ObjectThemeColor.
                // For the "Replace" window, we want the absolute RGB.
                
                Color baseColor = ConvertRgb(rgb);
                Color tinted = ApplyTint(baseColor, tint);
                int finalRgb = (tinted.R) | (tinted.G << 8) | (tinted.B << 16);

                ColorSelected?.Invoke(finalRgb);

                var shapes = _manager.GetSelectedShapes();
                if (shapes == null || shapes.Count == 0) return;

                for (int i = 1; i <= shapes.Count; i++)
                {
                    var shape = shapes[i];
                    if (type == "Fill")
                    {
                        shape.Fill.ForeColor.ObjectThemeColor = (Office.MsoThemeColorIndex)idx;
                        shape.Fill.ForeColor.TintAndShade = tint;
                        shape.Fill.Visible = Office.MsoTriState.msoTrue;
                    }
                    else if (type == "Border")
                    {
                        shape.Line.ForeColor.ObjectThemeColor = (Office.MsoThemeColorIndex)idx;
                        shape.Line.ForeColor.TintAndShade = tint;
                        shape.Line.Visible = Office.MsoTriState.msoTrue;
                    }
                    else if (type == "Text")
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            shape.TextFrame.TextRange.Font.Color.ObjectThemeColor = (Office.MsoThemeColorIndex)idx;
                            shape.TextFrame.TextRange.Font.Color.TintAndShade = tint;
                        }
                    }
                }
            }
            catch { }
        }

        private Button CreateCustomButton(int index)
        {
            Color color = _savedCustomColors[index];
            bool hasColor = color.A > 0;

            Button btn = new Button
            {
                Width = 22, Height = 22, Margin = new Thickness(0.5),
                Cursor = Cursors.Hand, Background = Brushes.Transparent, BorderThickness = new Thickness(0),
                Padding = new Thickness(0), ToolTip = "Right click to set color, left click to apply"
            };

            Border border = new Border
            {
                Width = 18, Height = 18, CornerRadius = new CornerRadius(3),
                BorderBrush = new SolidColorBrush(Color.FromRgb(80, 80, 80)),
                BorderThickness = new Thickness(0.5)
            };

            if (hasColor)
            {
                border.Background = new SolidColorBrush(color);
            }
            else
            {
                border.Background = new SolidColorBrush(Color.FromRgb(45, 45, 45));
                border.Child = new TextBlock 
                { 
                    Text = "+", Foreground = Brushes.Gray, FontSize = 10, 
                    HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center 
                };
            }

            btn.Content = border;
            btn.Click += (s, e) => {
                if (_savedCustomColors[index].A > 0)
                    ApplyCustomColor(_savedCustomColors[index]);
            };

            btn.MouseRightButtonUp += (s, e) => {
                using (var dialog = new System.Windows.Forms.ColorDialog())
                {
                    if (hasColor) dialog.Color = System.Drawing.Color.FromArgb(color.R, color.G, color.B);
                    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        var sc = dialog.Color;
                        _savedCustomColors[index] = Color.FromRgb(sc.R, sc.G, sc.B);
                        RefreshCustomButton(btn, index);
                    }
                }
            };

            return btn;
        }

        private void RefreshCustomButton(Button btn, int index)
        {
            Color color = _savedCustomColors[index];
            if (btn.Content is Border border)
            {
                border.Background = new SolidColorBrush(color);
                border.Child = null;
            }
        }

        private void ApplyCustomColor(Color color)
        {
            try
            {
                int rgb = (color.B << 16) | (color.G << 8) | color.R;
                ColorSelected?.Invoke(rgb);

                var shapes = _manager.GetSelectedShapes();
                if (shapes == null || shapes.Count == 0) return;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    shapes[i].Fill.ForeColor.RGB = rgb;
                    shapes[i].Fill.Visible = Office.MsoTriState.msoTrue;
                }
            }
            catch { }
        }

        private Color ConvertRgb(int rgb)
        {
            byte r = (byte)(rgb & 0xFF);
            byte g = (byte)((rgb >> 8) & 0xFF);
            byte b = (byte)((rgb >> 16) & 0xFF);
            return Color.FromRgb(r, g, b);
        }
    }
}
