using System;
using System.Linq;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Threading;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.UI.Controls;

namespace oPenEfficiency.UI.Panels
{
    internal static class SidebarControlFactory
    {
        public static Button CreateButton(
            string id,
            FeatureDisplayInfo info,
            Brush background,
            double iconSize,
            string iconColorMode,
            Func<string, object> findResource,
            Action<Button> execute,
            Action<ButtonBase> attachContextMenu)
        {
            var button = new Button { Name = id };
            Configure(button, id, info, background, iconSize, iconColorMode, findResource);
            button.Click += (sender, args) => execute(button);
            attachContextMenu(button);
            return button;
        }

        public static ToggleButton CreateToggleButton(
            string id,
            FeatureDisplayInfo info,
            Brush background,
            double iconSize,
            string iconColorMode,
            Func<string, object> findResource,
            Action<ToggleButton> execute,
            Action<ButtonBase> attachContextMenu)
        {
            var button = new ToggleButton { Name = id };
            Configure(button, id, info, background, iconSize, iconColorMode, findResource);
            button.Click += (sender, args) => execute(button);
            attachContextMenu(button);
            return button;
        }

        private static void Configure(
            ButtonBase button,
            string id,
            FeatureDisplayInfo info,
            Brush background,
            double iconSize,
            string iconColorMode,
            Func<string, object> findResource)
        {
            button.Style = (Style)findResource(button is ToggleButton
                ? "GridIconButtonToggleButton"
                : "GridIconButton");
            button.ToolTip = CreateTooltip(button, info, findResource);
            AutomationProperties.SetName(button, info.Name);
            AutomationProperties.SetHelpText(button,
                string.IsNullOrWhiteSpace(info.Tooltip) ? info.Description : info.Tooltip);
            button.Background = background;
            button.Content = CreateIcon(button, id, info.Color, iconSize, iconColorMode);
        }

        private static object CreateTooltip(
            ButtonBase button,
            FeatureDisplayInfo info,
            Func<string, object> findResource)
        {
            var panel = new StackPanel
            {
                MaxWidth = 250,
                Margin = (Thickness)findResource("Space2")
            };
            panel.Children.Add(new TextBlock
            {
                Text = info.Name,
                FontWeight = FontWeights.Bold,
                FontSize = (double)findResource("TypographyBodySize"),
                Foreground = (Brush)findResource("TextPrimaryBrush"),
                TextWrapping = TextWrapping.Wrap
            });

            if (!string.IsNullOrWhiteSpace(info.Tooltip) &&
                !IsNormalizedDuplicate(info.Name, info.Tooltip))
            {
                panel.Children.Add(new TextBlock
                {
                    Text = info.Tooltip,
                    FontSize = (double)findResource("TypographyCaptionSize"),
                    Foreground = (Brush)findResource("TextSecondaryBrush"),
                    TextWrapping = TextWrapping.Wrap,
                    Margin = (Thickness)findResource("RelatedControlGap")
                });
            }

            if (info.Variants != null && info.Variants.Count > 0)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = "Right-click for more options",
                    FontSize = (double)findResource("TypographyMicroSize"),
                    Foreground = (Brush)findResource("TextMutedBrush"),
                    TextWrapping = TextWrapping.Wrap,
                    Margin = (Thickness)findResource("RelatedControlGap")
                });
            }

            if (string.IsNullOrEmpty(info.Description)) return panel;

            var separator = new Border
            {
                Height = 1,
                Background = (Brush)findResource("BorderBrush"),
                Margin = (Thickness)findResource("Space4"),
                Visibility = Visibility.Collapsed
            };
            var description = new TextBlock
            {
                Text = info.Description,
                TextWrapping = TextWrapping.Wrap,
                FontSize = (double)findResource("TypographyCaptionSize"),
                Foreground = (Brush)findResource("TextSecondaryBrush"),
                LineHeight = 14,
                Visibility = Visibility.Collapsed
            };
            panel.Children.Add(separator);
            panel.Children.Add(description);

            var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1.5) };
            timer.Tick += (sender, args) =>
            {
                description.Visibility = Visibility.Visible;
                separator.Visibility = Visibility.Visible;
                timer.Stop();
            };
            button.MouseEnter += (sender, args) => timer.Start();
            button.MouseLeave += (sender, args) =>
            {
                timer.Stop();
                description.Visibility = Visibility.Collapsed;
                separator.Visibility = Visibility.Collapsed;
            };
            return panel;
        }

        private static bool IsNormalizedDuplicate(string left, string right)
        {
            return string.Equals(Normalize(left), Normalize(right), StringComparison.Ordinal);
        }

        private static string Normalize(string value)
        {
            return new string((value ?? string.Empty)
                .Where(char.IsLetterOrDigit)
                .Select(char.ToLowerInvariant)
                .ToArray());
        }

        private static Grid CreateIcon(
            ButtonBase button,
            string id,
            string color,
            double iconSize,
            string iconColorMode)
        {
            var container = new Grid
            {
                Width = iconSize + 4,
                Height = iconSize + 4,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };

            try
            {
                var brush = (Brush)new BrushConverter().ConvertFrom(color);
                var icon = new FeatureIcon
                {
                    FeatureId = id,
                    IconColor = brush,
                    ColorMode = IconColorModes.Normalize(iconColorMode),

                    Width = FeatureIconSizing.ResolveGlyphSize(iconSize),
                    Height = FeatureIconSizing.ResolveGlyphSize(iconSize),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                container.Children.Add(icon);

                Action updateState = () =>
                {
                    icon.IsEnabled = button.IsEnabled;
                    icon.VisualState = !button.IsEnabled
                        ? FeatureIconVisualState.Disabled
                        : button is ToggleButton toggle && toggle.IsChecked == true
                            ? FeatureIconVisualState.Selected
                            : FeatureIconVisualState.Normal;
                };
                button.IsEnabledChanged += (sender, args) => updateState();
                button.MouseEnter += (sender, args) =>
                {

                    if (button.IsEnabled) icon.VisualState = FeatureIconVisualState.Hover;
                };
                button.MouseLeave += (sender, args) =>
                {

                    updateState();
                };
                button.PreviewMouseLeftButtonDown += (sender, args) =>
                {
                    if (button.IsEnabled) icon.VisualState = FeatureIconVisualState.Pressed;
                };
                button.PreviewMouseLeftButtonUp += (sender, args) =>
                {
                    if (button.IsEnabled) icon.VisualState = FeatureIconVisualState.Hover;
                };
                if (button is ToggleButton toggleButton)
                {
                    toggleButton.Checked += (sender, args) => updateState();
                    toggleButton.Unchecked += (sender, args) => updateState();
                }
                updateState();
            }
            catch
            {
                container.Children.Add(new TextBlock { Text = "!" });
            }

            return container;
        }
    }
}
