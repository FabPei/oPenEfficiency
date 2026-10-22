using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using oPenEfficiency.Models;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Services.Commands;
using oPenEfficiency.UI;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Services
{
    internal sealed class QuickContextMenuService : IDisposable
    {
        private readonly PowerPoint.Application _application;
        private Window _popup;

        public QuickContextMenuService(PowerPoint.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public void ScheduleShow(int screenX, int screenY)
        {
            Application.Current?.Dispatcher.BeginInvoke(
                new Action(() => TryShow(screenX, screenY)),
                DispatcherPriority.ContextIdle);
        }

        public void Close()
        {
            if (_popup == null) return;
            var popup = _popup;
            _popup = null;
            try { popup.Close(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "QuickContextMenuService.Close"); }
        }

        public void Dispose() => Close();

        private void TryShow(int screenX, int screenY)
        {
            try
            {
                var config = JsonService.Load<SidebarConfig>("sidebar_config.json") ?? new SidebarConfig();
                if (!config.QuickContextMenuEnabled) { Close(); return; }
                var selection = _application.ActiveWindow?.Selection;
                if (selection == null ||
                    (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                     selection.Type != PowerPoint.PpSelectionType.ppSelectionText))
                {
                    Close();
                    return;
                }

                var catalog = FeatureLibrary.Catalog
                    .Where(item => item.UserVisible && !item.IsObsolete &&
                        item.ExecutionKind == FeatureExecutionKind.Command &&
                        item.Id != "BtnQuickContextMenu")
                    .ToDictionary(item => item.Id, StringComparer.Ordinal);
                IEnumerable<string> configuredIds = config.QuickContextMenuFeatureIds?.Count > 0
                    ? (IEnumerable<string>)config.QuickContextMenuFeatureIds
                    : QuickContextMenuSettings.DefaultFeatureIds;
                var ids = QuickContextMenuSettings.Normalize(
                    configuredIds,
                    new HashSet<string>(catalog.Keys, StringComparer.Ordinal));
                if (ids.Count == 0) return;

                Close();
                var panel = new StackPanel { Margin = new Thickness(4) };
                foreach (string id in ids)
                {
                    var feature = catalog[id];
                    var button = new Button
                    {
                        Content = feature.Name,
                        HorizontalContentAlignment = HorizontalAlignment.Left,
                        MinHeight = 28,
                        Padding = new Thickness(8, 3, 12, 3),
                        ToolTip = feature.Tooltip,
                        IsEnabled = MainSidebar.IsContextCommandAvailable(feature)
                    };
                    AutomationProperties.SetName(button, feature.Name);
                    button.Click += (sender, args) =>
                    {
                        MainSidebar.ExecuteContextCommand(id);
                        Close();
                    };
                    panel.Children.Add(button);
                }

                _popup = new Window
                {
                    Title = "openEfficiency Quick Menu",
                    Content = panel,
                    SizeToContent = SizeToContent.WidthAndHeight,
                    WindowStyle = WindowStyle.ToolWindow,
                    ResizeMode = ResizeMode.NoResize,
                    ShowInTaskbar = false,
                    ShowActivated = false,
                    Topmost = true,
                    Left = PxToDipX(screenX + 12),
                    Top = PxToDipY(screenY + 12),
                    MaxHeight = 440
                };
                _popup.PreviewKeyDown += (sender, args) => { if (args.Key == Key.Escape) Close(); };
                _popup.Closed += (sender, args) => { if (ReferenceEquals(_popup, sender)) _popup = null; };
                new WindowInteropHelper(_popup).Owner = (IntPtr)(_application.ActiveWindow?.HWND ?? _application.HWND);
                _popup.Show();
                KeepWithinWorkArea(_popup);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "QuickContextMenuService.TryShow");
                Close();
            }
        }

        private static double PxToDipX(int value)
        {
            using (var graphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
                return value * 96d / graphics.DpiX;
        }

        private static double PxToDipY(int value)
        {
            using (var graphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
                return value * 96d / graphics.DpiY;
        }

        private static void KeepWithinWorkArea(Window window)
        {
            window.Dispatcher.BeginInvoke(new Action(() =>
            {
                var area = SystemParameters.WorkArea;
                window.Left = Math.Max(area.Left, Math.Min(window.Left, area.Right - window.ActualWidth));
                window.Top = Math.Max(area.Top, Math.Min(window.Top, area.Bottom - window.ActualHeight));
            }), DispatcherPriority.Loaded);
        }
    }
}
