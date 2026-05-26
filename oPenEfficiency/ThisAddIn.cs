using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using oPenEfficiency.Features;
using oPenEfficiency.Utils;
using oPenEfficiency.UI;

namespace oPenEfficiency
{
    public partial class ThisAddIn
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        // Dictionary to track task panes per window (hWnd -> TaskPane)
        private Dictionary<int, CustomTaskPane> _windowTaskPanes = new Dictionary<int, CustomTaskPane>();
        private CustomTaskPane _assetManagerTaskPane;
        private CustomTaskPane _themeColorTaskPane;
        private oPenEfficiency.UI.ThermometerAdorner _thermometerAdorner;
        private PowerPoint.Shape _trackedShape;
        private oPenEfficiency.Services.SnapToObjectsManager _snapManager;
        private oPenEfficiency.Services.SnapToGridManager _snapGridManager;

        private System.Windows.Forms.Timer _selectionDebounceTimer;
        private PowerPoint.Selection _pendingSelection;
        private System.Windows.Forms.Timer _textEditPollingTimer;

        private float _dpiX = 96f;
        private float _dpiY = 96f;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Register global exception handlers for unhandled crashes
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            
            using (System.Drawing.Graphics g = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
            {
                _dpiX = g.DpiX;
                _dpiY = g.DpiY;
            }

            _selectionDebounceTimer = new System.Windows.Forms.Timer();
            _selectionDebounceTimer.Interval = 75;
            _selectionDebounceTimer.Tick += SelectionDebounceTimer_Tick;

            _textEditPollingTimer = new System.Windows.Forms.Timer();
            _textEditPollingTimer.Interval = 200;
            _textEditPollingTimer.Tick += TextEditPollingTimer_Tick;

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

                // Create sidebar for the first window
                CreateSidebarForActiveWindow();

                ShortcutManager.StartHook();
                File.AppendAllText(logPath, "Shortcut hook started\n");

                File.AppendAllText(logPath, "Event handlers and managers initialized\n");

                _snapManager = new oPenEfficiency.Services.SnapToObjectsManager(new PowerPointManager(this.Application));
                File.AppendAllText(logPath, "SnapToObjectsManager initialized\n");

                _snapGridManager = new oPenEfficiency.Services.SnapToGridManager(new PowerPointManager(this.Application));
                File.AppendAllText(logPath, "SnapToGridManager initialized\n");

                File.AppendAllText(logPath, "Startup completed successfully.\n");
            }
            catch (Exception ex)
            {
                File.AppendAllText(logPath, $"FATAL ERROR during startup: {ex}\n");
                File.AppendAllText(logPath, $"Stack trace: {ex.StackTrace}\n");
                System.Windows.Forms.MessageBox.Show(
                    "Error starting oPenEfficiency Add-in. Check " + logPath + " for details.",
                    "oPenEfficiency Error",
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
            // Check if WPF Application already exists
            if (System.Windows.Application.Current == null)
            {
                // Create a new WPF Application instance for this PowerPoint instance
                var app = new System.Windows.Application();
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
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
            {
                ExceptionLogger.Log(ex, "Global.UnhandledException", showErrorUI: true);
            }
        }

        private void Wpf_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            ExceptionLogger.Log(e.Exception, "WPF.DispatcherCrash", showErrorUI: true);
            e.Handled = true; // Prevent the app from crashing entirely if possible
        }
        
        private void Application_WindowActivate(PowerPoint.Presentation Pres, PowerPoint.DocumentWindow Wn)
        {
            try
            {
                CreateSidebarForActiveWindow();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.Application_WindowActivate");
            }
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            _pendingSelection = Sel;
            _selectionDebounceTimer?.Stop();
            _selectionDebounceTimer?.Start();
        }

        private int GetActiveWindowHWnd()
        {
            try
            {
                if (this.Application.ActiveWindow != null)
                {
                    return this.Application.ActiveWindow.HWND;
                }
            }
            catch { }
            return 0;
        }

        private CustomTaskPane GetSidebarForActiveWindow()
        {
            int hWnd = GetActiveWindowHWnd();
            if (hWnd == 0) return null;

            if (_windowTaskPanes.TryGetValue(hWnd, out var taskPane))
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
            if (_windowTaskPanes.TryGetValue(hWnd, out var existingPane))
            {
                return;
            }

            // Create new task pane for this window
            var ppManager = new PowerPointManager(this.Application);
            var host = new SidebarTaskPaneHost();
            var taskPane = CustomTaskPanes.Add(host, "oPen Efficiency");
            taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            taskPane.Width = 300;
            taskPane.Visible = true;

            _windowTaskPanes[hWnd] = taskPane;

            if (host.Sidebar != null)
                host.Sidebar.SetManager(ppManager);
        }

        private void Application_PresentationClose(PowerPoint.Presentation Pres)
        {
            try
            {
                // Prevent deleting the active window's task pane if a background presentation is closing
                PowerPoint.Presentation activePres = null;
                try { activePres = this.Application.ActivePresentation; } catch { }

                if (activePres != null)
                {
                    try
                    {
                        // If the presentation closing is not the currently active one, ignore.
                        if (Pres.Name != activePres.Name) return;
                    }
                    catch { }
                }

                int hWnd = GetActiveWindowHWnd();
                if (hWnd != 0 && _windowTaskPanes.TryGetValue(hWnd, out var taskPane))
                {
                    try
                    {
                        this.CustomTaskPanes.Remove(taskPane);
                    }
                    catch { }
                    _windowTaskPanes.Remove(hWnd);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.Application_PresentationClose");
            }
        }

        private void SelectionDebounceTimer_Tick(object sender, EventArgs e)
        {
            _selectionDebounceTimer?.Stop();
            if (_pendingSelection != null)
            {
                ProcessSelectionChange(_pendingSelection);
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

                // When text is selected, the ShapeRange contains the parent shape
                // This is more reliable than trying to get the parent from TextRange
                if (sel.ShapeRange != null && sel.ShapeRange.Count > 0)
                {
                    var parentShape = sel.ShapeRange[1];
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

        private void ProcessSelectionChange(PowerPoint.Selection Sel)
        {
            try
            {
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
                        if (Sel.HasChildShapeRange)
                        {
                            if (Sel.ChildShapeRange.Count == 1)
                            {
                                shape = Sel.ChildShapeRange[1];
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "ThisAddIn.ProcessSelectionChange.HasChildShapeRange");
                    }

                    PowerPoint.Shape targetShape = shape;
                    string altText = targetShape.AlternativeText ?? "";
                    
                    try 
                    {
                        if (!altText.StartsWith(ThermometerFeature.ThermometerTag) && targetShape.Child == Microsoft.Office.Core.MsoTriState.msoTrue && targetShape.ParentGroup != null)
                        {
                            targetShape = targetShape.ParentGroup;
                            altText = targetShape.AlternativeText ?? "";
                        }
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "ThisAddIn.ProcessSelectionChange.ParentGroup");
                    }

                    if (altText.StartsWith("oPE_HarveyBall")) ShowHarveyBallMenu();
                    else if (altText.StartsWith("oPE_Checkbox")) ShowCheckboxMenu(altText);
                    else if (altText.StartsWith(TrafficLightFeature.TrafficLightTag)) ShowTrafficLightMenu(altText);
                    else if (altText.StartsWith(StarRatingFeature.StarRatingTag)) ShowStarRatingToolbar(altText);
                    else if (altText.StartsWith(ThermometerFeature.ThermometerTag)) ShowThermometerToolbar(altText);
                    else if (altText.StartsWith(ProgressSeriesFeature.SeriesTag)) 
                {
                    System.Diagnostics.Debug.WriteLine("ProgressSeries detected - showing toolbar");
                    ShowProgressSeriesToolbar(targetShape);
                }
                    else if (altText.StartsWith(NumerationFeature.NumerationTag)) ShowNumerationToolbar(targetShape);
                }
            }
            catch (Exception)
            {
                // Silent in production, but let's show for debugging
                // System.Windows.Forms.MessageBox.Show("Selection Error: " + ex.ToString());
            }
        }

        private string GetTagValue(PowerPoint.Shape shape, string tagName)
        {
            try
            {
                if (shape.Tags.Count == 0) return "";
                for (int i = 1; i <= shape.Tags.Count; i++)
                {
                    if (shape.Tags.Name(i).Equals(tagName, StringComparison.OrdinalIgnoreCase))
                    {
                        return shape.Tags.Value(i);
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.GetTagValue");
            }
            return "";
        }

        private oPenEfficiency.UI.RatingToolbar _ratingToolbar;

        private void ShowStarRatingToolbar(string altText)
        {
            try
            {
                // Close existing toolbar if open
                if (_ratingToolbar != null)
                {
                    try { _ratingToolbar.Close(); } catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.ShowStarRatingToolbar.CloseOld"); }
                    _ratingToolbar = null;
                }

                var manager = new PowerPointManager(this.Application);
                _ratingToolbar = new oPenEfficiency.UI.RatingToolbar(manager, altText);
                _ratingToolbar.Closed += (s, args) => { _ratingToolbar = null; };

                // Use simple mouse positioning like traffic light menu
                var mousePos = System.Windows.Forms.Cursor.Position;
                _ratingToolbar.Left = mousePos.X + 10;
                _ratingToolbar.Top = mousePos.Y + 10;

                _ratingToolbar.Show();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ShowStarRatingToolbar");
            }
        }

        private oPenEfficiency.UI.ProgressSeriesToolbar _progressSeriesToolbar;
        private DateTime _progressSeriesOpenedTime;

        private void ShowProgressSeriesToolbar(PowerPoint.Shape shape)
        {
            try
            {
                if (_progressSeriesToolbar != null)
                {
                    // If already open, just update the target without closing
                    _progressSeriesToolbar.UpdateTarget(shape);
                    return;
                }

                var manager = new PowerPointManager(this.Application);
                _progressSeriesToolbar = new oPenEfficiency.UI.ProgressSeriesToolbar(manager, shape);
                _progressSeriesToolbar.Closed += (s, args) => { _progressSeriesToolbar = null; };
                _progressSeriesOpenedTime = DateTime.Now;

                // Position like other features: use shape position
                var window = Application.ActiveWindow;
                if (window != null)
                {
                    int pixelX = window.PointsToScreenPixelsX(shape.Left + shape.Width);
                    int pixelY = window.PointsToScreenPixelsY(shape.Top);

                    _progressSeriesToolbar.Left = (pixelX * 96.0 / _dpiX) + 10;
                    _progressSeriesToolbar.Top = (pixelY * 96.0 / _dpiY);
                }
                else
                {
                    // Fallback to mouse position if no window
                    var mousePos = System.Windows.Forms.Cursor.Position;
                    _progressSeriesToolbar.Left = mousePos.X + 10;
                    _progressSeriesToolbar.Top = mousePos.Y + 10;
                }

                System.Diagnostics.Debug.WriteLine("Showing ProgressSeriesToolbar");
                _progressSeriesToolbar.Show();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ShowProgressSeriesToolbar");
                System.Diagnostics.Debug.WriteLine($"ShowProgressSeriesToolbar error: {ex.Message}");
                System.Windows.MessageBox.Show($"Error showing Progress Series toolbar: {ex.Message}");
            }
        }

        private oPenEfficiency.UI.NumerationToolbar _numerationToolbar;

        private void ShowNumerationToolbar(PowerPoint.Shape shape)
        {
            try
            {
                if (_numerationToolbar != null)
                {
                    try { _numerationToolbar.Close(); } catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.ShowNumerationToolbar.CloseOld"); }
                    _numerationToolbar = null;
                }

                var manager = new PowerPointManager(this.Application);
                _numerationToolbar = new oPenEfficiency.UI.NumerationToolbar(manager, shape);

                var mousePos = System.Windows.Forms.Cursor.Position;
                _numerationToolbar.Left = (mousePos.X * 96.0 / _dpiX) - 210; // Center 420px width
                _numerationToolbar.Top = (mousePos.Y * 96.0 / _dpiY) + 20;

                _numerationToolbar.Show();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ShowNumerationToolbar");
            }
        }

        private void ShowThermometerToolbar(string altText)
        {
            try
            {
                var manager = new PowerPointManager(this.Application);
                var toolbar = new oPenEfficiency.UI.FloatingToolbar();
                
                // Parse state
                float percentage = 0.5f;
                int colorRgb = 0x0000FF; // Default
                
                var parts = altText.Split('|');
                if (parts.Length >= 2) float.TryParse(parts[1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out percentage);
                if (parts.Length >= 3) int.TryParse(parts[2], out colorRgb);

                toolbar.AddLabel("🌡️");
                
                // 1. Percentage Slider
                toolbar.AddSlider(0, 100, (int)(percentage * 100), (s, e) => 
                {
                    var track = s as System.Windows.Forms.TrackBar;
                    if (track != null)
                    {
                        float p = track.Value / 100f;
                        ThermometerFeature.Execute(manager, p, colorRgb);
                    }
                });

                toolbar.AddSeparator();

                // 2. Color Dropdown
                toolbar.AddDropDownButton("Color", null, (dropDown) => 
                {
                    int[] colors = new[] { 0x0000FF, 0x00FF00, 0xFF0000, 0x008CFF, 0x00FFFF };
                    string[] names = new[] { "Red", "Green", "Blue", "Orange", "Cyan" };
                    
                    for (int i = 0; i < colors.Length; i++)
                    {
                        int c = colors[i];
                        var item = dropDown.DropDownItems.Add(names[i]);
                        item.Click += (s, e) => 
                        {
                            ThermometerFeature.Execute(manager, percentage, c);
                            toolbar.Close();
                        };
                    }
                    
                    dropDown.DropDownItems.Add(new System.Windows.Forms.ToolStripSeparator());
                    var custom = dropDown.DropDownItems.Add("Custom...");
                    custom.Click += (s, e) => 
                    {
                        var dialog = new System.Windows.Forms.ColorDialog();
                        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            var color = dialog.Color;
                            int rgb = color.R | (color.G << 8) | (color.B << 16);
                            ThermometerFeature.Execute(manager, percentage, rgb);
                            toolbar.Close();
                        }
                    };
                });

                var mousePos = System.Windows.Forms.Cursor.Position;
                toolbar.Location = new System.Drawing.Point(mousePos.X + 10, mousePos.Y + 10);
                toolbar.Show();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Toolbar Error: " + ex.Message);
            }
        }

        private void UpdateAdornerPosition()
        {
            if (_thermometerAdorner == null || _trackedShape == null) return;
            try
            {
                var activeWin = this.Application.ActiveWindow;
                // Get screen coordinates of the shape
                float shapeLeft = _trackedShape.Left;
                float shapeTop = _trackedShape.Top;
                float shapeWidth = _trackedShape.Width;

                int screenX = activeWin.PointsToScreenPixelsX(shapeLeft + (shapeWidth / 2));
                int screenY = activeWin.PointsToScreenPixelsY(shapeTop);

                double adornerWidth = _thermometerAdorner.Width > 0 ? _thermometerAdorner.Width : 300;
                double adornerHeight = _thermometerAdorner.Height > 0 ? _thermometerAdorner.Height : 48;

                _thermometerAdorner.Left = screenX - (adornerWidth / 2);
                _thermometerAdorner.Top = screenY - adornerHeight - 20;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.UpdateAdornerPosition");
                CloseThermometerAdorner();
            }
        }

        private void CloseThermometerAdorner()
        {
            if (_thermometerAdorner != null)
            {
                try { _thermometerAdorner.Close(); } catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.CloseThermometerAdorner"); }
                _thermometerAdorner = null;
            }
            _trackedShape = null;

        }

        private void ShowTrafficLightMenu(string altText)
        {
            try
            {
                var manager = new PowerPointManager(this.Application);
                var trafficMenu = new oPenEfficiency.UI.TrafficLightMenu(manager, altText);
                
                var mousePos = System.Windows.Forms.Cursor.Position;
                // Add some DPI scaling roughly
                trafficMenu.Left = (mousePos.X * 96.0 / _dpiX) + 10;
                trafficMenu.Top = (mousePos.Y * 96.0 / _dpiY) + 10;
                
                trafficMenu.Show();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("TrafficLight menu error: " + ex.Message);
            }
        }

        private oPenEfficiency.UI.CheckboxToolbar _checkboxToolbar;

        private void ShowCheckboxMenu(string altText)
        {
            try
            {
                if (_checkboxToolbar != null)
                {
                    try { _checkboxToolbar.Close(); } catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.ShowCheckboxMenu.CloseOld"); }
                    _checkboxToolbar = null;
                }

                var manager = new PowerPointManager(this.Application);
                _checkboxToolbar = new oPenEfficiency.UI.CheckboxToolbar(manager, altText);
                _checkboxToolbar.Closed += (s, args) => { _checkboxToolbar = null; };

                _checkboxToolbar.Show();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ShowCheckboxMenu");
            }
        }

        private void ShowHarveyBallMenu()
        {
            var menu = new System.Windows.Forms.ContextMenuStrip();
            var levels = new[] 
            { 
                new { Name = "0%", Value = 0f },
                new { Name = "10%", Value = 0.1f },
                new { Name = "25%", Value = 0.25f },
                new { Name = "50%", Value = 0.5f },
                new { Name = "75%", Value = 0.75f },
                new { Name = "100%", Value = 1f }
            };

            foreach (var level in levels)
            {
                var item = menu.Items.Add(level.Name);
                float val = level.Value;
                item.Click += (s, args) => 
                {
                    var manager = new PowerPointManager(this.Application);
                    oPenEfficiency.Features.HarveyBallFeature.Execute(manager, val);
                };
            }
            
            menu.Items.Add(new System.Windows.Forms.ToolStripSeparator());
            var cycleItem = menu.Items.Add("Cycle (+25%)");
            cycleItem.Click += (s, args) => 
            {
                // Re-use logic from MainSidebar if possible, or duplicate for now
                var manager = new PowerPointManager(this.Application);
                CycleHarveyBall(manager);
            };

            var mousePos = System.Windows.Forms.Cursor.Position;
            menu.Show(mousePos);
        }

        private void CycleHarveyBall(PowerPointManager manager)
        {
             var selection = manager.GetSelectedShapes();
             float nextPercentage = 0.25f; 

             if (selection != null && selection.Count == 1)
             {
                 var selected = selection[1];
                 string alt = selected.AlternativeText;
                 if (alt.StartsWith("oPE_HarveyBall"))
                 {
                     var parts = alt.Split('|');
                     if (parts.Length == 2 && float.TryParse(parts[1], out float current))
                     {
                         nextPercentage = current + 0.25f;
                         if (nextPercentage > 1.01f) nextPercentage = 0f;
                     }
                 }
             }
             oPenEfficiency.Features.HarveyBallFeature.Execute(manager, nextPercentage);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Unregister event handlers first
            try
            {
                this.Application.WindowSelectionChange -= Application_WindowSelectionChange;
                this.Application.PresentationClose -= Application_PresentationClose;
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

            // Stop and dispose selection debounce timer
            if (_selectionDebounceTimer != null)
            {
                try
                {
                    _selectionDebounceTimer.Stop();
                    _selectionDebounceTimer.Dispose();
                }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.Shutdown.DebounceTimer"); }
                _selectionDebounceTimer = null;
            }

            // Remove all window-specific task panes
            try
            {
                var panesToRemove = new List<CustomTaskPane>(_windowTaskPanes.Values);
                foreach (var pane in panesToRemove)
                {
                    try
                    {
                        this.CustomTaskPanes.Remove(pane);
                    }
                    catch { }
                }
                _windowTaskPanes.Clear();

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
                System.Windows.Forms.MessageBox.Show("Error docking palette: " + ex.Message);
            }
        }

        /// <summary>
        /// Gets the SnapToObjectsManager instance for snap-to-objects functionality.
        /// </summary>
        public oPenEfficiency.Services.SnapToObjectsManager SnapManager => _snapManager;

        /// <summary>
        /// Gets the SnapToGridManager instance for snap-to-grid functionality.
        /// </summary>
        public oPenEfficiency.Services.SnapToGridManager SnapGridManager => _snapGridManager;

        public void OpenFeatureExplorer()
        {
            try
            {
                var wnd = new oPenEfficiency.UI.Dialogs.FeatureExplorerWindow();
                wnd.ShowDialog();
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "OpenFeatureExplorer");
            }
        }

        public void ToggleSidebar()
        {
            try
            {
                int hWnd = GetActiveWindowHWnd();
                if (hWnd == 0) return;

                CustomTaskPane taskPane = null;
                _windowTaskPanes.TryGetValue(hWnd, out taskPane);

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
                System.Windows.Forms.MessageBox.Show($"Error toggling sidebar: {ex.Message}", "oPenEfficiency Error",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
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
                    var host = new oPenEfficiency.UI.AssetManagerHost();
                    _assetManagerTaskPane = CustomTaskPanes.Add(host, "Asset Library");
                    _assetManagerTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating;
                    _assetManagerTaskPane.Width = 1000;
                    _assetManagerTaskPane.Height = 1000;
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
                System.Windows.Forms.MessageBox.Show("Error toggling Asset Manager: " + ex.Message);
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


