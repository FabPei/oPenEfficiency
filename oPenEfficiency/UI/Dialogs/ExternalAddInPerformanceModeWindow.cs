using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Shell;
using oPenEfficiency.Services;

namespace oPenEfficiency.UI.Dialogs
{
    internal sealed class ExternalAddInPerformanceModeWindow : Window
    {
        private readonly ExternalAddInPerformanceModeService _service;
        private readonly StackPanel _items = new StackPanel();
        private readonly List<Row> _rows = new List<Row>();
        private WindowChrome _chrome;
        private Grid _headerDragRegion;

        public ExternalAddInPerformanceModeWindow(ExternalAddInPerformanceModeService service)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            Title = "Configure Performance Mode";
            Width = 560;
            MinWidth = 420;
            MinHeight = 240;
            MaxHeight = 700;
            SizeToContent = SizeToContent.Height;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ShowInTaskbar = false;
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            SetResourceReference(StyleProperty, "CompactTaskWindowStyle");
            SetResourceReference(ForegroundProperty, "TextPrimaryBrush");
            _chrome = new WindowChrome
            {
                CaptionHeight = 32,
                ResizeBorderThickness = new Thickness(0),
                GlassFrameThickness = new Thickness(0),
                CornerRadius = new CornerRadius(8),
                UseAeroCaptionButtons = false
            };
            WindowChrome.SetWindowChrome(this, _chrome);
            Content = BuildLayout();
            Refresh();
            Loaded += (sender, args) => UpdateCaptionHeight();
            _headerDragRegion.SizeChanged += (sender, args) => UpdateCaptionHeight();
        }

        private void UpdateCaptionHeight()
        {
            if (_headerDragRegion == null || _headerDragRegion.ActualHeight <= 0) return;
            System.Windows.Point bottom = _headerDragRegion.TranslatePoint(
                new System.Windows.Point(0, _headerDragRegion.ActualHeight), this);
            _chrome.CaptionHeight = Math.Max(0, bottom.Y);
        }

        private UIElement BuildLayout()
        {
            // Compact Task Window shell: the content-hugging surface of Format Numbers / QR Code.
            var shell = new Border();
            shell.SetResourceReference(FrameworkElement.StyleProperty, "CompactTaskShell");

            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            _headerDragRegion = new Grid();
            _headerDragRegion.SetResourceReference(FrameworkElement.MinHeightProperty, "ModalHeaderHeight");
            _headerDragRegion.SetResourceReference(FrameworkElement.MarginProperty, "RelatedControlGap");
            _headerDragRegion.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            _headerDragRegion.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            var title = new TextBlock
            {
                Text = "CONFIGURE PERFORMANCE MODE",
                VerticalAlignment = VerticalAlignment.Center
            };
            title.SetResourceReference(FrameworkElement.StyleProperty, "DialogTitleText");
            Grid.SetColumn(title, 0);
            _headerDragRegion.Children.Add(title);
            var closeButton = new Button();
            closeButton.SetResourceReference(FrameworkElement.StyleProperty, "WindowCloseButton");
            closeButton.Click += (sender, args) => Close();
            closeButton.ToolTip = "Close";
            System.Windows.Automation.AutomationProperties.SetName(closeButton, "Close Configure Performance Mode");
            WindowChrome.SetIsHitTestVisibleInChrome(closeButton, true);
            Grid.SetColumn(closeButton, 1);
            _headerDragRegion.Children.Add(closeButton);
            Grid.SetRow(_headerDragRegion, 0);
            grid.Children.Add(_headerDragRegion);

            var content = new StackPanel();
            var intro = new TextBlock
            {
                Text = "Performance Mode temporarily pauses the external PowerPoint COM add-ins you select here. Nothing is selected automatically, and openEfficiency is always excluded. Enabling the mode pauses the selected active add-ins; disabling it or closing PowerPoint restores only the add-ins paused in this session.",
                TextWrapping = TextWrapping.Wrap
            };
            intro.SetResourceReference(FrameworkElement.StyleProperty, "ContextMetaText");
            intro.SetResourceReference(FrameworkElement.MarginProperty, "SectionGapCompact");
            content.Children.Add(intro);
            content.Children.Add(_items);

            var scroll = new ScrollViewer
            {
                Content = content,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Focusable = false
            };
            Grid.SetRow(scroll, 1);
            grid.Children.Add(scroll);

            var footer = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            footer.SetResourceReference(FrameworkElement.MarginProperty, "SectionGapCompact");
            footer.Children.Add(Button("Refresh", (s, e) => Refresh()));
            footer.Children.Add(Button("Select none", (s, e) => _rows.ForEach(row => row.CheckBox.IsChecked = false)));
            footer.Children.Add(Button("Select active", (s, e) => _rows.ForEach(row => row.CheckBox.IsChecked = row.State.IsConnected)));
            footer.Children.Add(Button("Cancel", (s, e) => Close()));
            footer.Children.Add(Button("Save", Save));
            Grid.SetRow(footer, 2);
            grid.Children.Add(footer);

            shell.Child = grid;
            return shell;
        }

        private void Refresh()
        {
            _items.Children.Clear();
            _rows.Clear();
            foreach (var state in _service.Discover())
            {
                var check = new CheckBox { IsChecked = state.IsSelected, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 10, 0) };
                var text = new StackPanel();
                text.Children.Add(new TextBlock { Text = state.DisplayName, FontWeight = FontWeights.SemiBold });
                text.Children.Add(new TextBlock { Text = state.Id + " · Status: " + state.Status + " · Restore: " + state.RestorationCapability, Opacity = 0.75, TextWrapping = TextWrapping.Wrap });
                var rowPanel = new DockPanel { Margin = new Thickness(0, 2, 0, 2) };
                DockPanel.SetDock(check, Dock.Left); rowPanel.Children.Add(check); rowPanel.Children.Add(text);
                _items.Children.Add(rowPanel);
                _items.Children.Add(new Separator());
                _rows.Add(new Row(state, check));
            }
            if (_rows.Count == 0)
                _items.Children.Add(new TextBlock { Text = "No external PowerPoint COM add-ins were detected.", Margin = new Thickness(0, 6, 0, 4) });
        }

        private void Save(object sender, RoutedEventArgs e)
        {
            var ids = _rows.Where(row => row.CheckBox.IsChecked == true).Select(row => row.State.Id);
            if (!_service.SaveSelection(ids))
            {
                MessageBox.Show(this, "The selection could not be saved. See the local diagnostics log.", Title, MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            DialogResult = true;
        }

        private static Button Button(string text, RoutedEventHandler click)
        {
            var button = new Button { Content = text, MinWidth = 84, MinHeight = 28, Margin = new Thickness(0, 0, 6, 0) };
            button.SetResourceReference(FrameworkElement.StyleProperty, "ToolbarButton");
            button.Click += click;
            return button;
        }

        private sealed class Row
        {
            public Row(ExternalAddInState state, CheckBox checkBox) { State = state; CheckBox = checkBox; }
            public ExternalAddInState State { get; }
            public CheckBox CheckBox { get; }
        }
    }
}
