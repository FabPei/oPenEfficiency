using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using oPenEfficiency.Utils;
using Xl = global::Microsoft.Office.Interop.Excel;

namespace oPenEfficiency.Excel.Features.Data
{
    /// <summary>One invalid-date finding shown in the results window.</summary>
    public sealed class DateFinding
    {
        public string Address { get; set; }
        public string Value { get; set; }
        public string Reason { get; set; }
    }

    /// <summary>
    /// Modeless floating window listing every invalid date (address, value, reason). Clicking
    /// a row jumps to and selects that cell in Excel. Built entirely in code (no XAML) because
    /// the Excel add-in project has no WPF build targets. It is modeless so the user can keep
    /// interacting with the worksheet while it stays open.
    /// </summary>
    public sealed class DateFindingsWindow : Window
    {
        private readonly Xl.Application _app;
        private readonly Xl.Worksheet _sheet;

        private static readonly Brush PanelBg = Frozen(Color.FromRgb(0x1F, 0x29, 0x37));
        private static readonly Brush RowBg = Frozen(Color.FromRgb(0x11, 0x18, 0x27));
        private static readonly Brush RowHover = Frozen(Color.FromRgb(0x37, 0x41, 0x51));
        private static readonly Brush TextMain = Frozen(Color.FromRgb(0xF9, 0xFA, 0xFB));
        private static readonly Brush TextDim = Frozen(Color.FromRgb(0x9C, 0xA3, 0xAF));
        private static readonly Brush Accent = Frozen(Color.FromRgb(0xF5, 0x9E, 0x0B));

        public DateFindingsWindow(Xl.Application app, Xl.Worksheet sheet, IList<DateFinding> findings)
        {
            _app = app;
            _sheet = sheet;

            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            ShowInTaskbar = false;
            Topmost = true;
            Width = 460;
            Height = 540;
            WindowStartupLocation = WindowStartupLocation.Manual;

            var wa = SystemParameters.WorkArea;
            Left = wa.Right - Width - 24;
            Top = wa.Top + 48;

            Content = BuildContent(findings);
        }

        private UIElement BuildContent(IList<DateFinding> findings)
        {
            var root = new Border
            {
                CornerRadius = new CornerRadius(12),
                Background = PanelBg,
                BorderBrush = Frozen(Color.FromRgb(0x37, 0x41, 0x51)),
                BorderThickness = new Thickness(1),
                Effect = new DropShadowEffect { BlurRadius = 18, ShadowDepth = 2, Opacity = 0.5, Color = Colors.Black }
            };

            var dock = new DockPanel { Margin = new Thickness(14) };

            // Header (draggable) + close button
            var header = new DockPanel();
            header.MouseLeftButtonDown += (s, e) => { if (e.ButtonState == MouseButtonState.Pressed) DragMove(); };

            var close = new Button
            {
                Content = "×",
                Width = 26,
                Height = 26,
                Foreground = TextDim,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                FontSize = 16,
                Cursor = Cursors.Hand
            };
            close.Click += (s, e) => Close();
            DockPanel.SetDock(close, Dock.Right);
            header.Children.Add(close);

            header.Children.Add(new TextBlock
            {
                Text = $"Invalid Date Values ({findings.Count})",
                Foreground = TextMain,
                FontSize = 15,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center
            });
            DockPanel.SetDock(header, Dock.Top);
            dock.Children.Add(header);

            var hint = new TextBlock
            {
                Text = "Click an entry to jump to the cell.",
                Foreground = TextDim,
                FontSize = 11,
                Margin = new Thickness(0, 6, 0, 10)
            };
            DockPanel.SetDock(hint, Dock.Top);
            dock.Children.Add(hint);

            var list = new StackPanel();
            foreach (var f in findings) list.Children.Add(BuildRow(f));

            dock.Children.Add(new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = list
            });

            root.Child = dock;
            return root;
        }

        private FrameworkElement BuildRow(DateFinding f)
        {
            var stack = new StackPanel { Margin = new Thickness(10, 6, 10, 6) };
            stack.Children.Add(new TextBlock
            {
                Text = f.Address + "   =   " + f.Value,
                Foreground = TextMain,
                FontWeight = FontWeights.SemiBold,
                TextTrimming = TextTrimming.CharacterEllipsis
            });
            stack.Children.Add(new TextBlock
            {
                Text = f.Reason,
                Foreground = Accent,
                FontSize = 11,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 2, 0, 0)
            });

            var button = new Button
            {
                Content = stack,
                Background = RowBg,
                BorderThickness = new Thickness(0),
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(0, 0, 0, 6),
                Padding = new Thickness(2),
                Cursor = Cursors.Hand,
                Tag = f
            };
            button.MouseEnter += (s, e) => button.Background = RowHover;
            button.MouseLeave += (s, e) => button.Background = RowBg;
            button.Click += (s, e) => NavigateTo(f);

            // Round the row corners via a hosting border.
            return new Border
            {
                CornerRadius = new CornerRadius(8),
                ClipToBounds = true,
                Child = button,
                Margin = new Thickness(0, 0, 0, 6)
            };
        }

        private void NavigateTo(DateFinding f)
        {
            try
            {
                var range = _sheet.Range[f.Address];
                _app.Goto(range, true);
                try { range.Select(); } catch { /* Navigation already succeeded; selection emphasis is optional. */ }
                try { _app.ActiveWindow.Activate(); } catch { /* Excel may reject activation while another modal window owns focus. */ }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DateFindingsWindow.NavigateTo");
            }
        }

        private static Brush Frozen(Color c)
        {
            var b = new SolidColorBrush(c);
            b.Freeze();
            return b;
        }
    }
}
