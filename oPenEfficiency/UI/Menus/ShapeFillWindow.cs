using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using oPenEfficiency.Features;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI.Menus
{
    internal sealed class ShapeFillWindow : Window
    {
        private readonly PowerPointManager _manager;
        private readonly ShapeFillService _service;

        public ShapeFillWindow(PowerPointManager manager)
        {
            _manager = manager ?? throw new ArgumentNullException(nameof(manager));
            _service = new ShapeFillService(manager);
            Title = "Shape Fill";
            Width = 276;
            SizeToContent = SizeToContent.Height;
            WindowStyle = WindowStyle.ToolWindow;
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;
            Topmost = true;
            Content = BuildContent();
            PreviewKeyDown += (sender, args) => { if (args.Key == Key.Escape) Close(); };
            Deactivated += (sender, args) => Close();
            Loaded += (sender, args) => PositionNearPointer();
        }

        private UIElement BuildContent()
        {
            var panel = new StackPanel { Margin = new Thickness(10) };
            panel.Children.Add(Header("Theme Colors"));
            foreach (var theme in ReadThemeColors())
            {
                var row = new WrapPanel { Margin = new Thickness(0, 1, 0, 1) };
                foreach (float tint in ShapeFillPalette.Tints)
                {
                    int rgb = ShapeFillPalette.ApplyTint(theme.Rgb, tint);
                    row.Children.Add(Swatch(rgb, theme.Name + " " + ShapeFillPalette.TintLabel(tint),
                        () => _service.ApplyTheme(theme.Index, tint)));
                }
                panel.Children.Add(row);
            }

            panel.Children.Add(Header("Standard Colors"));
            var standard = new WrapPanel();
            foreach (int rgb in ShapeFillPalette.StandardRgb)
                standard.Children.Add(Swatch(rgb, ShapeFillPalette.ToHex(rgb), () => _service.ApplyRgb(ToBgr(rgb))));
            panel.Children.Add(standard);

            var recent = ReadRecentColors().ToList();
            if (recent.Count > 0)
            {
                panel.Children.Add(Header("Recent Colors"));
                var recentPanel = new WrapPanel();
                foreach (int rgb in recent)
                    recentPanel.Children.Add(Swatch(rgb, "Recent " + ShapeFillPalette.ToHex(rgb), () => _service.ApplyRgb(ToBgr(rgb))));
                panel.Children.Add(recentPanel);
            }

            panel.Children.Add(new Separator { Margin = new Thickness(0, 6, 0, 4) });
            panel.Children.Add(Command("No Fill", () => _service.ApplyNoFill()));
            panel.Children.Add(Command("More Fill Colors...", OpenMoreColors));
            panel.Children.Add(Command("Eyedropper", () => PickFillColorFeature.Execute(_manager)));
            return panel;
        }

        private static TextBlock Header(string text) => new TextBlock
        {
            Text = text,
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(2, 5, 0, 3)
        };

        private Button Swatch(int rgb, string name, Func<bool> apply)
        {
            var button = new Button
            {
                Width = 38,
                Height = 24,
                Margin = new Thickness(1),
                Background = new SolidColorBrush(Color.FromRgb((byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb)),
                ToolTip = name,
                BorderThickness = new Thickness(1)
            };
            System.Windows.Automation.AutomationProperties.SetName(button, name);
            button.Click += (sender, args) => { if (apply()) Close(); };
            return button;
        }

        private Button Command(string name, Func<bool> action)
        {
            var button = new Button
            {
                Content = name,
                HorizontalContentAlignment = HorizontalAlignment.Left,
                MinHeight = 28,
                Margin = new Thickness(0, 1, 0, 1),
                Padding = new Thickness(8, 3, 8, 3)
            };
            System.Windows.Automation.AutomationProperties.SetName(button, name);
            button.Click += (sender, args) => { Close(); action(); };
            return button;
        }

        private bool OpenMoreColors()
        {
            var picker = new WpfColorPickerDialog("#FFFFFF") { Owner = Owner };
            if (picker.ShowDialog() != true) return false;
            var color = picker.SelectedColor;
            return _service.ApplyRgb(System.Drawing.ColorTranslator.ToOle(
                System.Drawing.Color.FromArgb(color.R, color.G, color.B)));
        }

        private IEnumerable<ThemeEntry> ReadThemeColors()
        {
            var app = _manager.GetApplication();
            var scheme = app?.ActivePresentation?.SlideMaster?.Theme?.ThemeColorScheme;
            var mappings = new[]
            {
                Tuple.Create(Office.MsoThemeColorSchemeIndex.msoThemeDark1, Office.MsoThemeColorIndex.msoThemeColorText1, "Dark 1"),
                Tuple.Create(Office.MsoThemeColorSchemeIndex.msoThemeLight1, Office.MsoThemeColorIndex.msoThemeColorBackground1, "Light 1"),
                Tuple.Create(Office.MsoThemeColorSchemeIndex.msoThemeDark2, Office.MsoThemeColorIndex.msoThemeColorText2, "Dark 2"),
                Tuple.Create(Office.MsoThemeColorSchemeIndex.msoThemeLight2, Office.MsoThemeColorIndex.msoThemeColorBackground2, "Light 2"),
                Tuple.Create(Office.MsoThemeColorSchemeIndex.msoThemeAccent1, Office.MsoThemeColorIndex.msoThemeColorAccent1, "Accent 1"),
                Tuple.Create(Office.MsoThemeColorSchemeIndex.msoThemeAccent2, Office.MsoThemeColorIndex.msoThemeColorAccent2, "Accent 2"),
                Tuple.Create(Office.MsoThemeColorSchemeIndex.msoThemeAccent3, Office.MsoThemeColorIndex.msoThemeColorAccent3, "Accent 3"),
                Tuple.Create(Office.MsoThemeColorSchemeIndex.msoThemeAccent4, Office.MsoThemeColorIndex.msoThemeColorAccent4, "Accent 4"),
                Tuple.Create(Office.MsoThemeColorSchemeIndex.msoThemeAccent5, Office.MsoThemeColorIndex.msoThemeColorAccent5, "Accent 5"),
                Tuple.Create(Office.MsoThemeColorSchemeIndex.msoThemeAccent6, Office.MsoThemeColorIndex.msoThemeColorAccent6, "Accent 6")
            };
            foreach (var map in mappings)
            {
                int bgr = scheme == null ? 0x808080 : scheme.Colors(map.Item1).RGB;
                var color = System.Drawing.ColorTranslator.FromOle(bgr);
                yield return new ThemeEntry(map.Item2, map.Item3, (color.R << 16) | (color.G << 8) | color.B);
            }
        }

        private static IEnumerable<int> ReadRecentColors()
        {
            var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
            foreach (string value in config?.RecentShapeFillColors ?? new List<string>())
            {
                if (ShapeFillPalette.TryParseHex(value, out int rgb)) yield return rgb;
            }
        }

        private static int ToBgr(int rgb) => ((rgb & 0xFF) << 16) | (rgb & 0x00FF00) | ((rgb >> 16) & 0xFF);

        private void PositionNearPointer()
        {
            if (!ThisAddIn.GetCursorPos(out var point)) return;
            using (var graphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
            {
                Left = point.X * 96d / graphics.DpiX + 10;
                Top = point.Y * 96d / graphics.DpiY + 10;
            }
            var area = SystemParameters.WorkArea;
            Left = Math.Max(area.Left, Math.Min(Left, area.Right - ActualWidth));
            Top = Math.Max(area.Top, Math.Min(Top, area.Bottom - ActualHeight));
        }

        private sealed class ThemeEntry
        {
            public ThemeEntry(Office.MsoThemeColorIndex index, string name, int rgb) { Index = index; Name = name; Rgb = rgb; }
            public Office.MsoThemeColorIndex Index { get; }
            public string Name { get; }
            public int Rgb { get; }
        }
    }

}
