using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using oPenEfficiency.Features;
using oPenEfficiency.Utils;
using oPenEfficiency.UI;
using oPenEfficiency.Services;

namespace oPenEfficiency
{
    public partial class ThisAddIn : oPenEfficiency.UI.Panels.ISidebarHost
    {
        private static bool _wpfDispatcherHandlerRegistered;

        PowerPoint.Application oPenEfficiency.UI.Panels.ISidebarHost.Application => this.Application;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        private oPenEfficiency.Services.SidebarPaneCoordinator _sidebarPanes;
        private CustomTaskPane _assetManagerTaskPane;
        private CustomTaskPane _themeColorTaskPane;
        private oPenEfficiency.UI.ThermometerAdorner _thermometerAdorner;
        private PowerPoint.Shape _trackedShape;
        private oPenEfficiency.Services.SnapToObjectsManager _snapManager;
        private oPenEfficiency.Services.SnapToGridManager _snapGridManager;
        private oPenEfficiency.Services.GanttInteractionCoordinator _ganttInteractionCoordinator;
        private oPenEfficiency.Utils.Win32MouseHook _sharedMouseHook;
        private oPenEfficiency.Services.QuickContextMenuService _quickContextMenu;
        private oPenEfficiency.Services.ExternalAddInPerformanceModeService _performanceMode;
        private oPenEfficiency.UI.Toolbars.GanttResizeHandlesWindow _ganttResizeHandles;

        private oPenEfficiency.Services.SelectionUpdateCoordinator _selectionUpdates;
        private System.Windows.Threading.DispatcherTimer _textEditPollingTimer;
        private System.Windows.Threading.DispatcherTimer _startupDeferredTimer;

        private float _dpiX = 96f;
        private float _dpiY = 96f;
        private bool _uiServicesInitialized;
        private oPenEfficiency.UI.Lifecycle.ModelessUiCoordinator _modelessUi;

        internal bool IsPerformanceModeEnabled => _performanceMode?.IsEnabled == true;

        internal bool SetPerformanceModeEnabled(bool enabled)
        {
            if (_performanceMode == null)
                _performanceMode = oPenEfficiency.Services.ExternalAddInPerformanceModeService.For(this.Application);
            return _performanceMode.SetEnabled(enabled);
        }

        internal void ConfigurePerformanceMode()
        {
            if (_performanceMode == null)
                _performanceMode = oPenEfficiency.Services.ExternalAddInPerformanceModeService.For(this.Application);
            var window = new oPenEfficiency.UI.Dialogs.ExternalAddInPerformanceModeWindow(_performanceMode);
            new System.Windows.Interop.WindowInteropHelper(window).Owner = (IntPtr)this.Application.HWND;
            window.ShowDialog();
        }

        void oPenEfficiency.UI.Panels.ISidebarHost.ConfigurePerformanceMode() => ConfigurePerformanceMode();


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Build marker: names the assembly file this PowerPoint session actually loaded
            // (diagnostic only, never throws; see %TEMP%\oPenEfficiency\startup_YYYYMMDD.log).
            StartupBuildLog.Record(typeof(ThisAddIn).Assembly, "oPenEfficiency");

            using (System.Drawing.Graphics g = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
            {
                _dpiX = g.DpiX;
                _dpiY = g.DpiY;
            }

            _modelessUi = new oPenEfficiency.UI.Lifecycle.ModelessUiCoordinator();
            _sidebarPanes = new oPenEfficiency.Services.SidebarPaneCoordinator(
                pane => this.CustomTaskPanes.Remove(pane));
            _selectionUpdates = new oPenEfficiency.Services.SelectionUpdateCoordinator(
                75,
                ProcessCurrentSelectionChange);

            _textEditPollingTimer = NewIdleDispatcherTimer(200, TextEditPollingTimer_Tick);

            _startupDeferredTimer = NewIdleDispatcherTimer(1200, StartupDeferredTimer_Tick);

            string logPath = Path.Combine(Path.GetTempPath(), $"oPenEfficiency_startup_{DateTime.Now:yyyyMMdd_HHmmss}.log");
            try
            {
                File.WriteAllText(logPath, $"Startup initiated at {DateTime.Now}\n");
                File.AppendAllText(logPath, $"PowerPoint Application: {this.Application?.Name}\n");
                File.AppendAllText(logPath, $"Process ID: {System.Diagnostics.Process.GetCurrentProcess().Id}\n");

                EnsureWpfApplication();
                File.AppendAllText(logPath, "WPF Application ensured\n");

                // Initialize Theme
                try
                {
                    var config = oPenEfficiency.JsonService.Load<oPenEfficiency.Models.SidebarConfig>("sidebar_config.json");
                    string theme = config?.AppTheme ?? "Dark";
                    oPenEfficiency.Services.ThemeManager.ApplyTheme(theme);
                    oPenEfficiency.Services.ThemeManager.ApplyFontFamily(config?.AppFontFamily);
                    File.AppendAllText(logPath, $"Theme {theme} applied\n");
                }
                catch (Exception ex)
                {
                    File.AppendAllText(logPath, $"Error initializing theme: {ex.Message}\n");
                }

                // Register window events BEFORE creating task panes
                this.Application.WindowSelectionChange += Application_WindowSelectionChange;
                this.Application.PresentationClose += Application_PresentationClose;
                this.Application.WindowActivate += Application_WindowActivate;
                this.Application.WindowBeforeRightClick += Application_WindowBeforeRightClick;
                _startupDeferredTimer.Start();
                File.AppendAllText(logPath, "Deferred UI initialization scheduled\n");

                File.AppendAllText(logPath, "Startup completed successfully.\n");
            }
            catch (Exception ex)
            {
                File.AppendAllText(logPath, $"FATAL ERROR during startup: {ex}\n");
                File.AppendAllText(logPath, $"Stack trace: {ex.StackTrace}\n");
                System.Windows.Forms.MessageBox.Show(
                    "Error starting the openEfficiency add-in. Check " + logPath + " for details.",
                    "openEfficiency Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Ensures a WPF Application singleton exists for this PowerPoint instance.
        /// VSTO add-ins do not create one automatically, but WPF Window.Show()
        /// requires it. ShutdownMode.OnExplicitShutdown prevents it from
        /// exiting when the last window closes.
        /// </summary>
        private static void EnsureWpfApplication()
        {
            var app = System.Windows.Application.Current;

            // Check if WPF Application already exists
            if (app == null)
            {
                // Create a new WPF Application instance for this PowerPoint instance
                app = new System.Windows.Application();
                app.ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown;

                // Load base resources
                try
                {
                    var controlsResource = new System.Windows.ResourceDictionary
                    {
                        Source = new Uri("pack://application:,,,/oPenEfficiency;component/UI/Styles/Controls.xaml", UriKind.RelativeOrAbsolute)
                    };
                    app.Resources.MergedDictionaries.Add(controlsResource);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error loading global resources: {ex.Message}");
                }
            }

            if (!_wpfDispatcherHandlerRegistered)
            {
                app.DispatcherUnhandledException += WpfApplication_DispatcherUnhandledException;
                _wpfDispatcherHandlerRegistered = true;
            }
        }

        // Set at the start of ThisAddIn_Shutdown. Static because the WPF dispatcher handler
        // below is static; one add-in instance per host process, so one flag is correct.
        private static volatile bool _shuttingDown;

        private static void WpfApplication_DispatcherUnhandledException(
            object sender,
            System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            e.Handled = true;
            // Shutdown-time ThreadAbortException is PowerPoint tearing the add-in down while
            // deferred UI initialization is still running (QA-P0 batch C finding): it is not a
            // feature failure and never user-facing, so it must reach neither the error log
            // nor the message box.
            if (_shuttingDown && e.Exception is System.Threading.ThreadAbortException) return;
            ExceptionLogger.Log(e.Exception, "ThisAddIn.WpfDispatcherUnhandledException");
            try
            {
                System.Windows.MessageBox.Show(
                    Services.Commands.CommandRegistry.UserVisibleFailureMessage,
                    "openEfficiency",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.WpfDispatcherUnhandledException.ShowError");
            }
        }

        /// <summary>
        /// QA-P0 diagnostic switch. When the isolated lifecycle harness sets
        /// OPE_QA_SKIP_WINDOW_ACTIVATE_SIDEBAR=1, the synchronous pane creation inside the
        /// WindowActivate callback is skipped so the harness can measure whether that path
        /// causes the native PowerPoint exit. It is not a product setting: without the
        /// variable the production behavior is unchanged.
        /// </summary>
        private static readonly bool SkipSidebarOnWindowActivate = string.Equals(
            Environment.GetEnvironmentVariable("OPE_QA_SKIP_WINDOW_ACTIVATE_SIDEBAR"),
            "1",
            StringComparison.Ordinal);

        private void Application_WindowActivate(PowerPoint.Presentation Pres, PowerPoint.DocumentWindow Wn)
        {
            try
            {
                CloseGanttResizeHandles();

                if (!_uiServicesInitialized)
                {
                    return;
                }

                PruneClosedSidebarPanes();
                if (SkipSidebarOnWindowActivate)
                {
                    return;
                }

                // QA-P0 note: a deferred variant of this call (BeginInvoke at ApplicationIdle
                // and at Normal priority) was measured in batches 16 and 17. It did not reduce
                // the native crash rate against the synchronous call (0/10 and 3/10 against
                // 4/10 and 6/10 fully green) and it starved the pane creation, so the
                // synchronous call is kept until there is evidence for a deferral.
                CreateSidebarForActiveWindow();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.Application_WindowActivate");
            }
        }

        /// <summary>
        /// Creates a timer that fires from the WPF dispatcher at idle priority. A WinForms
        /// WM_TIMER callback is delivered by the Win32 message pump, which also runs inside
        /// PowerPoint's native and externally driven COM calls; the callback then re-enters
        /// Office's automation layer while that call is still on the stack and faulted
        /// PowerPoint natively (ppcore.dll, 0xc0000005).
        /// </summary>
        private static System.Windows.Threading.DispatcherTimer NewIdleDispatcherTimer(
            int intervalMilliseconds,
            EventHandler tickHandler)
        {
            var timer = new System.Windows.Threading.DispatcherTimer(
                System.Windows.Threading.DispatcherPriority.ApplicationIdle)
            {
                Interval = TimeSpan.FromMilliseconds(intervalMilliseconds)
            };
            timer.Tick += tickHandler;
            return timer;
        }

        private void StartupDeferredTimer_Tick(object sender, EventArgs e)
        {
            _startupDeferredTimer?.Stop();
            // Cancel the deferred init when shutdown has already begun so the abort cannot
            // land inside EnsureUiServicesInitialized at all (QA-P0 batch C finding).
            if (_shuttingDown) return;
            EnsureUiServicesInitialized();
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            // Never retain PowerPoint's event Selection RCW. During native chart creation the
            // object can be transient and invalid by the time the debounce timer fires.
            _selectionUpdates?.Schedule();
        }

        private int GetActiveWindowHWnd()
        {
            if (ExceptionPolicy.TryComProbe(
                () => this.Application.ActiveWindow?.HWND ?? 0,
                out int windowHandle))
            {
                return windowHandle;
            }

            return 0;
        }

        private CustomTaskPane GetSidebarForActiveWindow()
        {
            int hWnd = GetActiveWindowHWnd();
            if (hWnd == 0) return null;

            if (_sidebarPanes != null && _sidebarPanes.TryGet(hWnd, out var taskPane))
            {
                return taskPane;
            }
            return null;
        }

        private void CreateSidebarForActiveWindow()
        {
            int hWnd = GetActiveWindowHWnd();
            if (hWnd == 0) return;

            // Check if we already have a task pane for this window
            if (_sidebarPanes != null && _sidebarPanes.TryGet(hWnd, out var existingPane))
            {
                return;
            }

            // Create new task pane for this window
            var ppManager = new PowerPointManager(this.Application);
            var host = new SidebarTaskPaneHost();
            var taskPane = CustomTaskPanes.Add(host, "openEfficiency Sidebar");
            taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            taskPane.Width = 300;
            taskPane.Visible = true;

            _sidebarPanes.Add(hWnd, taskPane);

            if (host.Sidebar != null)
            {
                host.Sidebar.SetManager(ppManager);
                host.Sidebar.SetHost(this);
            }
        }

        private void EnsureUiServicesInitialized()
        {
            if (_uiServicesInitialized)
            {
                return;
            }

            try
            {
                CreateSidebarForActiveWindow();
                ShortcutManager.StartHook(IsEligibleShortcutWindow);

                // One shared WH_MOUSE thread hook for all mouse-driven services (PQ-016);
                // created here and disposed last during shutdown.
                if (_sharedMouseHook == null)
                {
                    _sharedMouseHook = new oPenEfficiency.Utils.Win32MouseHook();
                }

                if (_quickContextMenu == null)
                {
                    _quickContextMenu = new oPenEfficiency.Services.QuickContextMenuService(this.Application);
                    _sharedMouseHook.LeftButtonDown += SharedMouseHook_CloseQuickContextMenu;
                }

                if (_snapManager == null)
                {
                    _snapManager = new oPenEfficiency.Services.SnapToObjectsManager(
                        new PowerPointManager(this.Application), _sharedMouseHook);
                }

                if (_snapGridManager == null)
                {
                    _snapGridManager = new oPenEfficiency.Services.SnapToGridManager(
                        new PowerPointManager(this.Application), _sharedMouseHook);
                }

                if (_ganttInteractionCoordinator == null)
                {
                    _ganttInteractionCoordinator = new oPenEfficiency.Services.GanttInteractionCoordinator(
                        new PowerPointManager(this.Application),
                        HandleGanttLeftMouseDown,
                        GlobalMouseHook_RightButtonUp,
                        _sharedMouseHook);
                }

                _uiServicesInitialized = true;
            }
            catch (Exception ex)
            {
                // A ThreadAbortException that arrives while the host is shutting down is
                // teardown, not a feature failure: the abort rethrows after the catch anyway,
                // so suppressing the log entry loses nothing (QA-P0 batch C finding).
                if (!(_shuttingDown && ex is System.Threading.ThreadAbortException))
                {
                    ExceptionLogger.Log(ex, "ThisAddIn.EnsureUiServicesInitialized");
                }
            }
        }

        private bool IsEligibleShortcutWindow()
        {
            if (!_uiServicesInitialized || this.Application == null)
            {
                return false;
            }

            try
            {
                return this.Application.ActivePresentation != null &&
                       this.Application.ActiveWindow != null &&
                       this.Application.ActiveWindow.HWND != 0 &&
                       GetForegroundWindow() ==
                           (IntPtr)this.Application.ActiveWindow.HWND;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.IsEligibleShortcutWindow");
                return false;
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private void Application_PresentationClose(PowerPoint.Presentation Pres)
        {
            try
            {
                HashSet<int> closingWindowHandles = GetPresentationWindowHandles(Pres);
                if (closingWindowHandles.Contains(GetActiveWindowHWnd())) CloseGanttResizeHandles();
                foreach (int windowHandle in closingWindowHandles)
                    _sidebarPanes?.Remove(windowHandle);
                PruneClosedSidebarPanes();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.Application_PresentationClose");
            }
        }

        private HashSet<int> GetPresentationWindowHandles(PowerPoint.Presentation presentation)
        {
            var handles = new HashSet<int>();
            PowerPoint.DocumentWindows windows = null;
            try
            {
                windows = presentation?.Windows;
                if (windows == null) return handles;
                for (int index = 1; index <= windows.Count; index++)
                {
                    PowerPoint.DocumentWindow window = null;
                    try
                    {
                        window = windows[index];
                        if (window.HWND != 0) handles.Add(window.HWND);
                    }
                    catch (COMException)
                    {
                        // A closing document window can disappear during enumeration.
                    }
                    catch (InvalidComObjectException)
                    {
                        // A closing document window can disappear during enumeration.
                    }
                    finally
                    {
                        if (window != null)
                            ExceptionPolicy.RunBestEffortComCleanup(
                                () => Marshal.ReleaseComObject(window),
                                "ThisAddIn.GetPresentationWindowHandles.ReleaseWindow");
                    }
                }
            }
            catch (COMException)
            {
                // Presentation.Windows may already be disconnected when the close event arrives.
            }
            catch (InvalidComObjectException)
            {
                // Presentation.Windows may already be disconnected when the close event arrives.
            }
            finally
            {
                if (windows != null)
                    ExceptionPolicy.RunBestEffortComCleanup(
                        () => Marshal.ReleaseComObject(windows),
                        "ThisAddIn.GetPresentationWindowHandles.ReleaseWindows");
            }
            return handles;
        }

        private void PruneClosedSidebarPanes()
        {
            if (_sidebarPanes == null || this.Application == null) return;
            var liveHandles = new HashSet<int>();
            PowerPoint.DocumentWindows windows = null;
            try
            {
                windows = this.Application.Windows;
                for (int index = 1; index <= windows.Count; index++)
                {
                    PowerPoint.DocumentWindow window = null;
                    try
                    {
                        window = windows[index];
                        if (window.HWND != 0) liveHandles.Add(window.HWND);
                    }
                    catch (COMException)
                    {
                        // PowerPoint can remove a window while activation events are dispatched.
                    }
                    catch (InvalidComObjectException)
                    {
                        // PowerPoint can remove a window while activation events are dispatched.
                    }
                    finally
                    {
                        if (window != null)
                            ExceptionPolicy.RunBestEffortComCleanup(
                                () => Marshal.ReleaseComObject(window),
                                "ThisAddIn.PruneClosedSidebarPanes.ReleaseWindow");
                    }
                }
                _sidebarPanes.Prune(liveHandles);
            }
            catch (COMException)
            {
                // Keep existing panes when the live-window probe is temporarily unavailable.
            }
            catch (InvalidComObjectException)
            {
                // Keep existing panes when the live-window probe is temporarily unavailable.
            }
            finally
            {
                if (windows != null)
                    ExceptionPolicy.RunBestEffortComCleanup(
                        () => Marshal.ReleaseComObject(windows),
                        "ThisAddIn.PruneClosedSidebarPanes.ReleaseWindows");
            }
        }

        private void TextEditPollingTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                var sel = this.Application?.ActiveWindow?.Selection;
                if (sel == null)
                {
                    _textEditPollingTimer?.Stop();
                    return;
                }

                if (sel.Type != PowerPoint.PpSelectionType.ppSelectionText)
                {
                    _textEditPollingTimer?.Stop();
                    return;
                }

                // Protect against COMException when accessing ShapeRange during text selection
                PowerPoint.ShapeRange shapeRange = null;
                try { shapeRange = sel.ShapeRange; }
                catch (COMException)
                {
                    // Text selections do not always expose ShapeRange.
                }

                if (shapeRange != null && shapeRange.Count > 0)
                {
                    var parentShape = shapeRange[1];
                    var taskPane = GetSidebarForActiveWindow();
                    var sidebar = (taskPane?.Control as SidebarTaskPaneHost)?.Sidebar;
                    if (sidebar != null)
                    {
                        sidebar.UpdateSizePanel(parentShape);
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.TextEditPollingTimer_Tick");
            }
        }

        private void ProcessCurrentSelectionChange()
        {
            PowerPoint.DocumentWindow activeWindow = null;
            PowerPoint.Selection selection = null;
            try
            {
                activeWindow = this.Application?.ActiveWindow;
                selection = activeWindow?.Selection;
                if (selection != null)
                {
                    ProcessSelectionChange(selection);
                }
            }
            catch (COMException ex)
            {
                // PowerPoint can still be finalizing a native object (notably a chart) when the
                // deferred refresh runs. The next selection event retries with a fresh RCW.
                ExceptionLogger.Log(ex, "ThisAddIn.ProcessCurrentSelectionChange.TransientSelection");
            }
            catch (InvalidComObjectException ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ProcessCurrentSelectionChange.InvalidSelection");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ProcessCurrentSelectionChange");
            }
        }

        private void ProcessSelectionChange(PowerPoint.Selection Sel)
        {
            try
            {
                // Auto-close GanttToolbar if selection changes away from Gantt
                if (_ganttToolbar != null)
                {
                    bool stillOnGantt = false;
                    try
                    {
                        if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && Sel.ShapeRange.Count == 1)
                        {
                            stillOnGantt = PowerPointShapeManager.IsGanttShape(Sel.ShapeRange[1]);
                        }
                    }
                    catch (COMException)
                    {
                        // Selection can become stale while PowerPoint dispatches the event.
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "ThisAddIn.ProcessSelectionChange.GanttToolbarProbe");
                    }

                    if (!stillOnGantt)
                    {
                        try { _ganttToolbar.Close(); }
                        catch (InvalidOperationException)
                        {
                            // The modeless toolbar is already closing or closed.
                        }
                        catch (Exception ex)
                        {
                            ExceptionLogger.Log(ex, "ThisAddIn.ProcessSelectionChange.CloseGanttToolbar");
                        }
                        _ganttToolbar = null;
                    }
                }

                if (_ganttResizeHandles != null)
                {
                    bool stillOnGantt = false;
                    try
                    {
                        if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && Sel.ShapeRange.Count == 1)
                        {
                            stillOnGantt = PowerPointShapeManager.IsGanttShape(Sel.ShapeRange[1]);
                        }
                    }
                    catch (COMException)
                    {
                        // Selection can become stale while PowerPoint dispatches the event.
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "ThisAddIn.ProcessSelectionChange.GanttResizeProbe");
                    }

                    if (!stillOnGantt)
                    {
                        CloseGanttResizeHandles();
                    }
                }

                // Forward to sidebar for ShapeSizePanel live update
                var taskPane = GetSidebarForActiveWindow();
                var sidebar = (taskPane?.Control as SidebarTaskPaneHost)?.Sidebar;
                if (sidebar == null) return;

                sidebar?.OnSelectionChange(Sel);

                // Notify Asset Manager if open
                if (_assetManagerTaskPane != null && _assetManagerTaskPane.Visible)
                {
                    (_assetManagerTaskPane.Control as AssetManagerHost)?.OnSelectionChange(Sel);
                }

                // Start polling timer when text is selected inside a shape
                if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    _textEditPollingTimer?.Start();

                    try
                    {
                        if (Sel.ShapeRange != null && Sel.ShapeRange.Count == 1)
                        {
                            var textShape = Sel.ShapeRange[1];
                            var inlineLinkAtSelection = oPenEfficiency.Features.Utilities.InsertExcelValueFeature.GetInlineAtSelection(Sel);
                            if (inlineLinkAtSelection != null &&
                                !string.IsNullOrEmpty(GetTagValue(textShape, oPenEfficiency.Features.Utilities.InsertExcelValueFeature.InlineTag)))
                            {
                                ShowInlineExcelValueToolbar(textShape);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "ThisAddIn.ProcessSelectionChange.TextInlineToolbar");
                    }
                }
                else
                {
                    _textEditPollingTimer?.Stop();
                }

                if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && Sel.ShapeRange.Count == 1)
                {
                    var shape = Sel.ShapeRange[1];
                    try
                    {
                        if (Sel.HasChildShapeRange && Sel.ChildShapeRange.Count == 1)
                        {
                            shape = Sel.ChildShapeRange[1];
                        }
                    }
                    catch (COMException)
                    {
                        // ChildShapeRange is optional and can disappear during selection changes.
                    }

                    PowerPoint.Shape targetShape = shape;
                    string altText = targetShape.AlternativeText ?? "";
                    string trafficLightState = TrafficLightFeature.GetPersistedState(targetShape);

                    if (altText.StartsWith("oPE_HarveyBall")) ShowHarveyBallMenu();
                    else if (oPenEfficiency.Core.Services.CheckboxStatePolicy.IsCheckboxTag(altText)) ShowCheckboxMenu(altText);
                    else if (!string.IsNullOrEmpty(trafficLightState)) ShowTrafficLightMenu(trafficLightState);
                    else if (altText.StartsWith(StarRatingFeature.StarRatingTag)) ShowStarRatingToolbar(altText);
                    else if (altText.StartsWith(ThermometerFeature.ThermometerTag)) ShowThermometerToolbar(altText);
                    else if (altText.StartsWith(ProgressSeriesFeature.SeriesTag)) ShowProgressSeriesToolbar(targetShape);
                    else if (altText.StartsWith(NumerationFeature.NumerationTag)) ShowNumerationToolbar(targetShape);
                    else if (altText.StartsWith(oPenEfficiency.Features.Utilities.TableLinkFeature.Tag)) ShowTableLinkToolbar(targetShape);
                    else if (altText.StartsWith(oPenEfficiency.Features.Utilities.InsertExcelValueFeature.Tag)) ShowExcelValueToolbar(targetShape);

                    // Independent of AlternativeText: text boxes can hold inline linked Excel values.
                    if (!string.IsNullOrEmpty(GetTagValue(targetShape, oPenEfficiency.Features.Utilities.InsertExcelValueFeature.InlineTag)))
                    {
                        ShowInlineExcelValueToolbar(targetShape);
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ProcessSelectionChange");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // From here on, shutdown-time ThreadAbortException noise is expected teardown
            // behavior and must not be reported as a feature failure (see the logger guards).
            _shuttingDown = true;
            CloseGanttResizeHandles();

            if (_modelessUi != null)
            {
                try { _modelessUi.Dispose(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.Shutdown.ModelessUi"); }
                _modelessUi = null;
            }

            // Unregister event handlers first
            try
            {
                this.Application.WindowSelectionChange -= Application_WindowSelectionChange;
                this.Application.PresentationClose -= Application_PresentationClose;
                this.Application.WindowActivate -= Application_WindowActivate;
                this.Application.WindowBeforeRightClick -= Application_WindowBeforeRightClick;
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.Shutdown.UnregisterEvents"); }

            try { ShortcutManager.StopHook(); } catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.Shutdown.ShortcutManager"); }

            // Dispose SnapToObjectsManager (releases mouse hook)
            if (_snapManager != null)
            {
                try { _snapManager.Dispose(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.Shutdown.SnapManager"); }
                _snapManager = null;
            }

            // Dispose SnapToGridManager (releases mouse hook)
            if (_snapGridManager != null)
            {
                try { _snapGridManager.Dispose(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.Shutdown.SnapGridManager"); }
                _snapGridManager = null;
            }

            if (_ganttInteractionCoordinator != null)
            {
                try { _ganttInteractionCoordinator.Dispose(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.Shutdown.GanttInteractionCoordinator"); }
                _ganttInteractionCoordinator = null;
            }

            if (_performanceMode != null)
            {
                try { _performanceMode.Dispose(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.Shutdown.PerformanceMode"); }
                _performanceMode = null;
            }

            if (_quickContextMenu != null)
            {
                try
                {
                    if (_sharedMouseHook != null)
                        _sharedMouseHook.LeftButtonDown -= SharedMouseHook_CloseQuickContextMenu;
                    _quickContextMenu.Dispose();
                }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.Shutdown.QuickContextMenu"); }
                _quickContextMenu = null;
            }

            // The shared mouse hook outlives its subscribers and is disposed last (PQ-016).
            if (_sharedMouseHook != null)
            {
                try { _sharedMouseHook.Dispose(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.Shutdown.SharedMouseHook"); }
                _sharedMouseHook = null;
            }

            if (_selectionUpdates != null)
            {
                try { _selectionUpdates.Dispose(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.Shutdown.SelectionUpdates"); }
                _selectionUpdates = null;
            }

            if (_textEditPollingTimer != null)
            {
                try
                {
                    // DispatcherTimer detaches with Stop() and an unsubscribed Tick handler;
                    // it is not IDisposable.
                    _textEditPollingTimer.Stop();
                    _textEditPollingTimer.Tick -= TextEditPollingTimer_Tick;
                }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.Shutdown.TextEditPollingTimer"); }
                _textEditPollingTimer = null;
            }

            if (_startupDeferredTimer != null)
            {
                try
                {
                    // DispatcherTimer detaches with Stop() and an unsubscribed Tick handler;
                    // it is not IDisposable. Every timer is unsubscribed symmetrically with
                    // its subscription (QA-P0 timer-teardown rule).
                    _startupDeferredTimer.Stop();
                    _startupDeferredTimer.Tick -= StartupDeferredTimer_Tick;
                }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.Shutdown.StartupDeferredTimer"); }
                _startupDeferredTimer = null;
            }

            // Remove all window-specific task panes
            try
            {
                _sidebarPanes?.Dispose();
                _sidebarPanes = null;

                if (_assetManagerTaskPane != null)
                {
                    this.CustomTaskPanes.Remove(_assetManagerTaskPane);
                    _assetManagerTaskPane = null;
                }
                if (_themeColorTaskPane != null)
                {
                    this.CustomTaskPanes.Remove(_themeColorTaskPane);
                    _themeColorTaskPane = null;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.Shutdown.CustomTaskPanes");
            }
        }

        public void ToggleThemeColorPane()
        {
            try
            {
                if (_themeColorTaskPane != null)
                {
                    _themeColorTaskPane.Visible = !_themeColorTaskPane.Visible;
                }
                else
                {
                    var manager = new PowerPointManager(this.Application);
                    var host = new ThemeColorPaneHost(manager);
                    _themeColorTaskPane = CustomTaskPanes.Add(host, "Theme Palette");
                    _themeColorTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                    _themeColorTaskPane.Width = 180;
                    _themeColorTaskPane.Visible = true;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ToggleThemeColorPane");
                System.Windows.Forms.MessageBox.Show("Error docking palette: " + ex.Message);
            }
        }

        /// <summary>
        /// Gets the SnapToObjectsManager instance for snap-to-objects functionality.
        /// </summary>
        public oPenEfficiency.Services.SnapToObjectsManager SnapManager
        {
            get
            {
                EnsureUiServicesInitialized();
                return _snapManager;
            }
        }

        /// <summary>
        /// Gets the SnapToGridManager instance for snap-to-grid functionality.
        /// </summary>
        public oPenEfficiency.Services.SnapToGridManager SnapGridManager
        {
            get
            {
                EnsureUiServicesInitialized();
                return _snapGridManager;
            }
        }

        public void OpenFeatureExplorer(string initialCategory = null, bool focusSearch = false)
        {
            try
            {
                EnsureUiServicesInitialized();
                var wnd = new oPenEfficiency.UI.Dialogs.FeatureExplorerWindow(initialCategory, focusSearch);
                new System.Windows.Interop.WindowInteropHelper(wnd).Owner = (IntPtr)Application.HWND;
                wnd.ShowDialog();
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "OpenFeatureExplorer");
                throw;
            }
        }

        public void OpenIconBrowser(
            oPenEfficiency.UI.Dialogs.IconBrowserStartView startView =
                oPenEfficiency.UI.Dialogs.IconBrowserStartView.Default)
        {
            try
            {
                EnsureUiServicesInitialized();
                var wnd = new oPenEfficiency.UI.Dialogs.IconBrowserWindow(
                    new PowerPointManager(this.Application),
                    startView);
                wnd.ShowDialog();
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "OpenIconBrowser");
                throw;
            }
        }

        public void OpenIllustrationLibrary(PowerPointManager manager = null)
        {
            try
            {
                EnsureUiServicesInitialized();
                var wnd = new oPenEfficiency.UI.Dialogs.IllustrationLibraryWindow(manager ?? new PowerPointManager(this.Application));
                wnd.ShowDialog();
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "OpenIllustrationLibrary");
                throw;
            }
        }

        public void ToggleSidebar()
        {
            try
            {
                EnsureUiServicesInitialized();
                int hWnd = GetActiveWindowHWnd();
                if (hWnd == 0) return;

                CustomTaskPane taskPane = null;
                _sidebarPanes?.TryGet(hWnd, out taskPane);

                if (taskPane == null)
                {
                    // Create new task pane for this window
                    CreateSidebarForActiveWindow();
                }
                else
                {
                    // Toggle visibility
                    taskPane.Visible = !taskPane.Visible;

                    // Refresh selection display when showing
                    if (taskPane.Visible)
                    {
                        var sidebar = (taskPane.Control as SidebarTaskPaneHost)?.Sidebar;
                        var selection = this.Application.ActiveWindow?.Selection;
                        if (sidebar != null && selection != null)
                        {
                            sidebar.OnSelectionChange(selection);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ToggleSidebar");
                System.Windows.Forms.MessageBox.Show($"Error toggling Sidebar: {ex.Message}", "openEfficiency Error",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        public bool IsAssetManagerVisible => _assetManagerTaskPane?.Visible == true;

        public void OpenAssetManager(string section = null, bool saveCurrentSelection = false)
        {
            if (_assetManagerTaskPane == null || !_assetManagerTaskPane.Visible)
                ToggleAssetManager();

            var host = _assetManagerTaskPane?.Control as AssetManagerHost;
            if (!string.IsNullOrWhiteSpace(section)) host?.SelectSection(section);
            if (saveCurrentSelection) host?.SaveCurrentSelection();
        }

        public void ToggleAssetManager()
        {
            try
            {
                if (_assetManagerTaskPane != null)
                {
                    _assetManagerTaskPane.Visible = !_assetManagerTaskPane.Visible;

                    if (_assetManagerTaskPane.Visible)
                    {
                        var selection = this.Application.ActiveWindow?.Selection;
                        if (selection != null)
                        {
                            (_assetManagerTaskPane.Control as AssetManagerHost)?.OnSelectionChange(selection);
                        }
                    }
                }
                else
                {
                    var host = new oPenEfficiency.UI.AssetManagerHost(
                        new PowerPointManager(this.Application));
                    _assetManagerTaskPane = CustomTaskPanes.Add(host, "Asset Library");
                    // Compact docked default (docs/product/ui-space-audit-asset-library-2026-10-11.md AM-P-04):
                    // the former 1000x1000 floating pane outgrows common work areas and occludes the slide.
                    _assetManagerTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                    _assetManagerTaskPane.Width = 480;
                    _assetManagerTaskPane.VisibleChanged += (s, e) =>
                    {
                        try
                        {
                            var sidebar = (GetSidebarForActiveWindow()?.Control as SidebarTaskPaneHost)?.Sidebar;
                            sidebar?.RefreshButtonStates();
                        }
                        catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.AssetManager.VisibleChanged"); }
                    };
                    _assetManagerTaskPane.Visible = true;

                    // Initial selection check for new pane
                    var selection = this.Application.ActiveWindow?.Selection;
                    if (selection != null)
                    {
                        host.OnSelectionChange(selection);
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ToggleAssetManager");
                throw;
            }
        }


        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new oPenEfficiency.UI.MainRibbon();
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}


