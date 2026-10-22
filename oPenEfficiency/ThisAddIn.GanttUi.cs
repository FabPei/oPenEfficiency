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
    public partial class ThisAddIn
    {
        private void LogDebug(string message)
        {
            try
            {
                string logPath = Path.Combine(Path.GetTempPath(), "oPenEfficiency_Gantt_Debug.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss.fff}] {message}\n");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Gantt debug logging failed: " + ex.Message);
            }
        }

        public static void LogDebugStatic(string message)
        {
            try
            {
                string logPath = Path.Combine(Path.GetTempPath(), "oPenEfficiency_Gantt_Debug.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss.fff}] {message}\n");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Gantt debug logging failed: " + ex.Message);
            }
        }

        private DateTime _lastRightClickTime = DateTime.MinValue;

        private void GlobalMouseHook_RightButtonUp(object sender, Win32MouseHook.MouseHookEventArgs e)
        {
            // DEBOUNCE: If BeforeRightClick already handled this, skip.
            if ((DateTime.Now - _lastRightClickTime).TotalMilliseconds < 300) return;

            LogDebug($"Global Hook: Right Click Up at ({e.ScreenX}, {e.ScreenY})");
            bool? dummy = false;
            HandleRightClick(e.ScreenX, e.ScreenY, ref dummy);
        }

        private void Application_WindowBeforeRightClick(PowerPoint.Selection Sel, ref bool Cancel)
        {
            LogDebug("--- Native BeforeRightClick Event ---");
            POINT p;
            if (!GetCursorPos(out p)) return;

            bool? cancelWrapper = Cancel;
            bool handled = HandleRightClick(p.X, p.Y, ref cancelWrapper);
            if (cancelWrapper.HasValue) Cancel = cancelWrapper.Value;

            if (handled)
            {
                _lastRightClickTime = DateTime.Now;
            }
            else
            {
                _quickContextMenu?.ScheduleShow(p.X, p.Y);
            }
        }

        private void SharedMouseHook_CloseQuickContextMenu(object sender, Win32MouseHook.MouseHookEventArgs e)
        {
            System.Windows.Application.Current?.Dispatcher.BeginInvoke(
                new Action(() => _quickContextMenu?.Close()));
        }

        private bool HandleRightClick(int screenX, int screenY, ref bool? Cancel)
        {
            try
            {
                var window = Application.ActiveWindow;
                if (window == null) return false;

                if (!TryScreenToSlidePoint(screenX, screenY, out float slidePtX, out float slidePtY))
                {
                    LogDebug("ABORT: Could not convert screen to slide coordinates.");
                    return false;
                }

                LogDebug($"Final Slide Points: {slidePtX:F1}, {slidePtY:F1}");

                PowerPoint.Shape mainGanttGroup = null;
                PowerPoint.Shape clickedShape = null;

                try
                {
                    try
                    {
                        object obj = window.RangeFromPoint(screenX, screenY);
                        if (obj is PowerPoint.Shape s) clickedShape = s;
                        else if (obj is PowerPoint.ShapeRange sr && sr.Count > 0) clickedShape = sr[1];
                    }
                    catch (Exception ex)
                    {
                        LogDebug($"RangeFromPoint failed: {ex.Message}");
                    }

                    // FALLBACK 1: Selection
                    if (clickedShape == null)
                    {
                        try
                        {
                            if (window.Selection.Type != PowerPoint.PpSelectionType.ppSelectionNone)
                            {
                                clickedShape = window.Selection.ShapeRange[1];
                                LogDebug($"Selection Fallback: {clickedShape.Name}");
                            }
                        }
                        catch (COMException)
                        {
                            // The current selection is only a fallback and may be unavailable.
                        }
                    }

                    if (clickedShape != null)
                    {
                        mainGanttGroup = FindGanttRootGroupRecursive(clickedShape);
                    }

                    // FALLBACK 2: Coordinate boundaries
                    if (mainGanttGroup == null)
                    {
                        PowerPoint.Slide slide = null;
                        try { slide = (PowerPoint.Slide)window.View.Slide; }
                        catch (COMException)
                        {
                            // No editable slide is active for this window.
                        }
                        if (slide != null)
                        {
                            foreach (PowerPoint.Shape shape in slide.Shapes)
                            {
                                if (IsGanttRootGroup(shape))
                                {
                                    // 20pt tolerance for background areas
                                    if (slidePtX >= shape.Left - 20 && slidePtX <= (shape.Left + shape.Width + 20) &&
                                        slidePtY >= shape.Top - 20 && slidePtY <= (shape.Top + shape.Height + 20))
                                    {
                                        mainGanttGroup = shape;
                                        LogDebug($"Coordinate Fallback found Root: {mainGanttGroup.Name}");
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogDebug($"Error in hit testing: {ex.Message}");
                }

                if (mainGanttGroup != null)
                {
                    LogDebug($"Gantt Chart DETECTED. Root: {mainGanttGroup.Name}. Showing custom menu.");

                    // SUPPRESSION
                    if (Cancel.HasValue) Cancel = true;
                    try
                    {
                        System.Windows.Forms.SendKeys.Send("{ESC}");

                        System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(async () =>
                        {
                            await System.Threading.Tasks.Task.Delay(100);
                            System.Windows.Forms.SendKeys.Send("{ESC}");
                        }), System.Windows.Threading.DispatcherPriority.Background);
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "ThisAddIn.HandleRightClick.SuppressNativeMenu");
                    }

                    PowerPoint.Shape taskShape = IsRealTaskShape(clickedShape) ? clickedShape : FindGanttChildAtPoint(mainGanttGroup, slidePtX, slidePtY);
                    POINT mousePos = new POINT { X = screenX, Y = screenY };

                    if (taskShape != null)
                    {
                        ShowGanttResizeHandles(taskShape);
                        ShowGanttStylingToolbar(taskShape);

                        // ALSO show the context menu for styling (like shapes)
                        string taskIdStr = PowerPointShapeManager.GetGanttId(taskShape);
                        if (Guid.TryParse(taskIdStr, out Guid taskId))
                        {
                            var vm = GetActiveGanttViewModel();
                            var task = vm?.Project.Tasks.FirstOrDefault(t => t.Id == taskId);
                            bool isTaskLabel = PowerPointShapeManager.GetGanttType(taskShape) == PowerPointShapeManager.GanttShapeType.RowAnchor.ToString();
                            ShowGanttContextMenu(mainGanttGroup, slidePtX, slidePtY, mousePos, task, isTaskLabel);
                        }
                    }
                    else
                    {
                        ShowGanttContextMenu(mainGanttGroup, slidePtX, slidePtY, mousePos, null);
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogDebug($"FATAL in HandleRightClick: {ex}");
            }
            return false;
        }

        private bool TryScreenToSlidePoint(int screenX, int screenY, out float pointX, out float pointY)
        {
            pointX = 0; pointY = 0;
            try
            {
                var window = Application.ActiveWindow;
                var pres = Application.ActivePresentation;
                if (window == null || pres == null) return false;

                float slideWidth = pres.PageSetup.SlideWidth;
                float slideHeight = pres.PageSetup.SlideHeight;

                int px0 = window.PointsToScreenPixelsX(0);
                int px1 = window.PointsToScreenPixelsX(slideWidth);
                int py0 = window.PointsToScreenPixelsY(0);
                int py1 = window.PointsToScreenPixelsY(slideHeight);

                float scaleX = (float)(slideWidth / Math.Max(1, (px1 - px0)));
                float scaleY = (float)(slideHeight / Math.Max(1, (py1 - py0)));

                pointX = (screenX - px0) * scaleX;
                pointY = (screenY - py0) * scaleY;
                return true;
            }
            catch (COMException)
            {
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.TryScreenToSlidePoint");
                return false;
            }
        }

        private bool IsGanttRootGroup(PowerPoint.Shape s)
        {
            if (s == null) return false;
            try
            {
                if (s.Name == GanttRenderer.GANTT_GROUP_NAME) return true;
                return PowerPointShapeManager.HasGanttProjectTag(s);
            }
            catch (COMException)
            {
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.IsGanttRootGroup");
                return false;
            }
        }

        private PowerPoint.Shape FindGanttRootGroupRecursive(object obj)
        {
            if (obj == null) return null;
            try
            {
                if (obj is PowerPoint.Shape s)
                {
                    if (IsGanttRootGroup(s)) return s;

                    if (TryGetParentGroup(s, out PowerPoint.Shape parentGroup) && parentGroup != null)
                    {
                        return FindGanttRootGroupRecursive(parentGroup);
                    }

                    return FindGanttRootGroupRecursive(s.Parent);
                }
                if (obj is PowerPoint.GroupShapes gs) return FindGanttRootGroupRecursive(gs.Parent);
            }
            catch (COMException)
            {
                // A stale group hierarchy is a negative probe result.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.FindGanttRootGroupRecursive");
            }
            return null;
        }

        private static bool TryGetParentGroup(PowerPoint.Shape shape, out PowerPoint.Shape parentGroup)
        {
            try
            {
                parentGroup = shape.ParentGroup;
                return true;
            }
            catch (Exception ex) when (ex is COMException || ex is UnauthorizedAccessException)
            {
                // PowerPoint reports an ungrouped shape as either COMException or E_ACCESSDENIED,
                // depending on the Office build.
                parentGroup = null;
                return false;
            }
        }

        private bool IsRealTaskShape(PowerPoint.Shape s)
        {
            if (s == null) return false;
            try {
                string id = PowerPointShapeManager.GetGanttId(s);
                return !string.IsNullOrEmpty(id) && id != Guid.Empty.ToString();
            }
            catch (COMException)
            {
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.IsRealTaskShape");
                return false;
            }
        }



        private oPenEfficiency.UI.Menus.GanttContextMenu _ganttContextMenu;

        private void ShowGanttContextMenu(PowerPoint.Shape group, double slidePtX, double slidePtY, POINT mousePos, oPenEfficiency.Models.GanttTask targetTask = null, bool isTaskLabel = false)
        {
            LogDebug("Attempting to show Gantt Context Menu...");
            var vm = GetActiveGanttViewModel();
            if (vm == null)
            {
                LogDebug("ABORT: No active Gantt ViewModel found.");
                return;
            }

            if (PowerPointShapeManager.GetChartId(group) != vm.Project.ChartId.ToString())
            {
                LogDebug("ABORT: Clicked chart does not belong to the active Gantt editor.");
                return;
            }

            if (vm.Project.IsLinkedToExcel)
            {
                LogDebug("ABORT: Project is linked to Excel (Read-Only).");
                return;
            }

            if (_ganttContextMenu != null)
            {
                try { _ganttContextMenu.Close(); }
                catch (InvalidOperationException)
                {
                    // The modeless menu is already closing or closed.
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "ThisAddIn.ShowGanttContextMenu.ClosePrevious");
                }
                _ganttContextMenu = null;
            }

            var window = Application.ActiveWindow;
            double screenGroupLeft = window.PointsToScreenPixelsX(group.Left);
            double screenGroupTop = window.PointsToScreenPixelsY(group.Top);
            double screenGroupRight = window.PointsToScreenPixelsX(group.Left + group.Width);
            double screenGroupBottom = window.PointsToScreenPixelsY(group.Top + group.Height);

            double relativePtX = slidePtX - group.Left;
            double relativePtY = slidePtY - group.Top;

            // 1. Calculate Insertion Index & Shade Range
            double headerHeight = GanttRenderer.GetHeaderHeight(vm.Project);
            DateTime? shadeStart = null;
            DateTime? shadeEnd = null;
            int index = -1;

            LogDebug($"Header Check: relativePtY={relativePtY:F1}, headerHeight={headerHeight:F1} | BoundingBox: {group.Left},{group.Top} {group.Width}x{group.Height}");

            if (relativePtY < headerHeight + 5) // 5pt tolerance
            {
                LogDebug("DETECTION: Header Area Clicked.");
                int levelRow = (int)Math.Floor(relativePtY / 20.0);
                if (levelRow < 0) levelRow = 0;

                var enabledLevels = new List<string>();
                if (vm.Project.ShowYears) enabledLevels.Add("Years");
                if (vm.Project.ShowMonths) enabledLevels.Add("Months");
                if (vm.Project.ShowWeeks) enabledLevels.Add("Weeks");
                if (vm.Project.ShowDays) enabledLevels.Add("Days");

                LogDebug($"Header Levels: Row={levelRow}, EnabledCount={enabledLevels.Count}, Levels=[{string.Join(",", enabledLevels)}]");

                if (levelRow < enabledLevels.Count)
                {
                    string clickedLevel = enabledLevels[levelRow];
                    LogDebug($"Level Identified: {clickedLevel}");

                    double tableWidth = GanttRenderer.GetTableWidth(vm.Project);
                    double pointsPerDay = GanttRenderer.GetPointsPerDay(vm.Project, group.Width);
                    double totalTimelineWidth = Math.Max(1, Math.Ceiling((vm.Project.ProjectEnd - vm.Project.ProjectStart).TotalDays)) * pointsPerDay;
                    double timelineLeftAbsolute = group.Left + tableWidth;

                    DateTime clickedDate = GanttMathEngine.CoordinateToDate(slidePtX, timelineLeftAbsolute, totalTimelineWidth, vm.Project.ProjectStart, vm.Project.ProjectEnd).Date;
                    LogDebug($"CoordinateToDate result: {clickedDate:yyyy-MM-dd} (X={slidePtX:F1}, TimelineLeft={timelineLeftAbsolute:F1})");

                    switch (clickedLevel)
                    {
                        case "Years":
                            shadeStart = new DateTime(clickedDate.Year, 1, 1);
                            shadeEnd = shadeStart.Value.AddYears(1);
                            break;
                        case "Months":
                            shadeStart = new DateTime(clickedDate.Year, clickedDate.Month, 1);
                            shadeEnd = shadeStart.Value.AddMonths(1);
                            break;
                        case "Weeks":
                            int diff = (7 + (clickedDate.DayOfWeek - DayOfWeek.Monday)) % 7;
                            shadeStart = clickedDate.AddDays(-1 * diff).Date;
                            shadeEnd = shadeStart.Value.AddDays(7);
                            break;
                        case "Days":
                            shadeStart = clickedDate.Date;
                            shadeEnd = clickedDate.Date.AddDays(1);
                            break;
                    }
                    LogDebug($"Shade Range Calculated: {shadeStart:yyyy-MM-dd} to {shadeEnd:yyyy-MM-dd}");
                }
            }
            else
            {
                LogDebug("DETECTION: Task Area Clicked.");
                index = GanttRenderer.GetTaskIndexAtVerticalOffset(vm.Project, relativePtY - headerHeight);
            }

            double tableWidthForDate = GanttRenderer.GetTableWidth(vm.Project);
            double pointsPerDayForDate = GanttRenderer.GetPointsPerDay(vm.Project, group.Width);
            double totalTimelineWidthForDate = Math.Max(1, Math.Ceiling((vm.Project.ProjectEnd - vm.Project.ProjectStart).TotalDays)) * pointsPerDayForDate;
            double timelineLeftAbsoluteForDate = group.Left + tableWidthForDate;

            DateTime date = GanttMathEngine.CoordinateToDate(slidePtX, timelineLeftAbsoluteForDate, totalTimelineWidthForDate, vm.Project.ProjectStart, vm.Project.ProjectEnd).Date;

            if (date < vm.Project.ProjectStart) date = vm.Project.ProjectStart;
            if (date > vm.Project.ProjectEnd) date = vm.Project.ProjectEnd;

            bool isFlexible = !vm.Project.FollowGanttRules;

            // Derive targetTask from row index when not found via shape hit-test
            // (e.g. click landed in the label column but not on the text shape itself)
            if (targetTask == null && index >= 0 && index < vm.Project.Tasks.Count)
                targetTask = vm.Project.Tasks[index];

            try
            {
                _ganttContextMenu = new oPenEfficiency.UI.Menus.GanttContextMenu(vm, index, date, isFlexible, shadeStart, shadeEnd, targetTask, isTaskLabel);
                _modelessUi.TrackWindow("GanttContextMenu", _ganttContextMenu);
                new System.Windows.Interop.WindowInteropHelper(_ganttContextMenu).Owner = (IntPtr)Application.HWND;

                // RAW COORDINATES: Log and set directly to cursor pos
                double targetX = (mousePos.X * 96.0 / _dpiX) - 120; // Offset to left
                double targetY = (mousePos.Y * 96.0 / _dpiY) - 50;  // Offset slightly up
                LogDebug($"Placing ContextMenu at Screen Pixels ({mousePos.X}, {mousePos.Y}) -> WPF Points ({targetX:F1}, {targetY:F1}) using DPI ({_dpiX}, {_dpiY})");

                _ganttContextMenu.Left = targetX;
                _ganttContextMenu.Top = targetY;

                System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(async () =>
                {
                    try
                    {
                        // Wait for native menu to spawn
                        await System.Threading.Tasks.Task.Delay(100);
                        System.Windows.Forms.SendKeys.Send("{ESC}");
                        await System.Threading.Tasks.Task.Delay(50);

                        if (_ganttContextMenu != null)
                        {
                            _ganttContextMenu.Show();
                            _ganttContextMenu.Activate();
                            _ganttContextMenu.Focus();
                            LogDebug($"Context Menu VISIBLE={_ganttContextMenu.IsVisible}, Opacity={_ganttContextMenu.Opacity}, Size={_ganttContextMenu.ActualWidth}x{_ganttContextMenu.ActualHeight}");
                        }
                    }
                    catch (Exception ex) { LogDebug($"Delayed show error: {ex}"); }
                }), System.Windows.Threading.DispatcherPriority.Normal);
            }
            catch (Exception ex)
            {
                LogDebug($"ERROR creating context menu: {ex}");
            }
        }

        private PowerPoint.Shape FindGanttChildAtPoint(PowerPoint.Shape group, double slidePtX, double slidePtY)
        {
            try
            {
                if (group.Type != Office.MsoShapeType.msoGroup) return group;

                foreach (PowerPoint.Shape child in group.GroupItems)
                {
                    string id = PowerPointShapeManager.GetGanttId(child);
                    if (string.IsNullOrEmpty(id) || id == Guid.Empty.ToString()) continue;

                    try
                    {
                        double tolerance = 5.0; // 5 points tolerance
                        double l = child.Left - tolerance;
                        double t = child.Top - tolerance;
                        double r = child.Left + child.Width + tolerance;
                        double b = child.Top + child.Height + tolerance;

                        if (slidePtX >= l && slidePtX <= r && slidePtY >= t && slidePtY <= b)
                        {
                            return child;
                        }
                    }
                    catch (COMException)
                    {
                        // Child bounds can disappear while a grouped shape is edited.
                    }
                }
            }
            catch (COMException)
            {
                // GroupItems can become stale during hit testing.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.FindGanttChildAtPoint");
            }
            return null;
        }

        private void HandleGanttLeftMouseDown(int screenX, int screenY)
        {
            try
            {
                if (_ganttResizeHandles != null && _ganttResizeHandles.IsPointerOverHandle(screenX, screenY))
                {
                    return;
                }

                var viewModel = GetActiveGanttViewModel();
                if (viewModel == null || viewModel.Project.IsLinkedToExcel)
                {
                    CloseGanttResizeHandles();
                    return;
                }

                if (!TryScreenToSlidePoint(screenX, screenY, out float slideX, out float slideY))
                {
                    CloseGanttResizeHandles();
                    return;
                }

                PowerPoint.Shape ganttGroup = null;
                PowerPoint.Slide slide = null;
                try { slide = (PowerPoint.Slide)Application.ActiveWindow.View.Slide; }
                catch (COMException)
                {
                    // No editable slide is active.
                }
                if (slide == null)
                {
                    CloseGanttResizeHandles();
                    return;
                }

                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (IsGanttRootGroup(shape) &&
                        slideX >= shape.Left && slideX <= shape.Left + shape.Width &&
                        slideY >= shape.Top && slideY <= shape.Top + shape.Height)
                    {
                        ganttGroup = shape;
                        break;
                    }
                }

                if (ganttGroup == null)
                {
                    CloseGanttResizeHandles();
                    return;
                }

                var hitShape = FindGanttChildAtPoint(ganttGroup, slideX, slideY);
                if (hitShape == null)
                {
                    CloseGanttResizeHandles();
                    return;
                }

                ShowGanttResizeHandles(hitShape);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.HandleGanttLeftMouseDown");
                CloseGanttResizeHandles();
            }
        }

        private void ShowGanttResizeHandles(PowerPoint.Shape hitShape)
        {
            try
            {
                string taskIdText = PowerPointShapeManager.GetGanttId(hitShape);
                if (!Guid.TryParse(taskIdText, out Guid taskId) || taskId == Guid.Empty)
                {
                    CloseGanttResizeHandles();
                    return;
                }

                var viewModel = GetActiveGanttViewModel();
                if (viewModel == null || PowerPointShapeManager.GetChartId(hitShape) != viewModel.Project.ChartId.ToString())
                {
                    CloseGanttResizeHandles();
                    return;
                }
                var task = viewModel?.Project.Tasks.FirstOrDefault(item => item.Id == taskId);
                if (task == null || task.IsMilestone || task.IsSummary || task.IsEventLine || viewModel.Project.IsLinkedToExcel)
                {
                    CloseGanttResizeHandles();
                    return;
                }

                PowerPoint.Slide slide = null;
                try { slide = (PowerPoint.Slide)Application.ActiveWindow.View.Slide; }
                catch (COMException)
                {
                    // No editable slide is active.
                }
                var taskBar = PowerPointShapeManager.GetShapeByTypeAndId(
                    slide,
                    PowerPointShapeManager.GanttShapeType.TaskBar,
                    taskId,
                    viewModel.Project.ChartId);
                if (taskBar == null)
                {
                    CloseGanttResizeHandles();
                    return;
                }

                if (_ganttResizeHandles != null && _ganttResizeHandles.IsLoaded && _ganttResizeHandles.MatchesTask(taskId))
                {
                    _ganttResizeHandles.RefreshPosition();
                    return;
                }

                CloseGanttResizeHandles();
                _ganttResizeHandles = new oPenEfficiency.UI.Toolbars.GanttResizeHandlesWindow(
                    new PowerPointManager(Application), taskBar, viewModel, taskId);
                _modelessUi.TrackWindow("GanttResizeHandles", _ganttResizeHandles);
                _ganttResizeHandles.Closed += (s, e) => _ganttResizeHandles = null;
                _ganttResizeHandles.Show();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ShowGanttResizeHandles");
                CloseGanttResizeHandles();
            }
        }

        private void CloseGanttResizeHandles()
        {
            if (_ganttResizeHandles == null) return;
            try { _ganttResizeHandles.Close(); }
            catch (InvalidOperationException)
            {
                // The modeless window is already closing or closed.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.CloseGanttResizeHandles");
            }
            _ganttResizeHandles = null;
        }

        private oPenEfficiency.UI.Toolbars.GanttToolbar _ganttToolbar;

        private void ShowGanttStylingToolbar(PowerPoint.Shape shape)
        {
            try
            {
                var vm = GetActiveGanttViewModel();
                if (vm == null)
                {
                    // If no active UI found, let's at least try to find the Task ID to see if we should show a message
                    string taskIdStr = PowerPointShapeManager.GetGanttId(shape);
                    System.Diagnostics.Debug.WriteLine($"Gantt Right-Click: No active ViewModel found for Task {taskIdStr}");
                    return;
                }

                if (PowerPointShapeManager.GetChartId(shape) != vm.Project.ChartId.ToString())
                {
                    return;
                }

                if (_ganttToolbar != null)
                {
                    try { _ganttToolbar.Close(); }
                    catch (InvalidOperationException)
                    {
                        // The modeless toolbar is already closing or closed.
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "ThisAddIn.ShowGanttStylingToolbar.ClosePrevious");
                    }
                    _ganttToolbar = null;
                }

                _ganttToolbar = new oPenEfficiency.UI.Toolbars.GanttToolbar(shape, vm);
                _modelessUi.TrackWindow("GanttToolbar", _ganttToolbar);
                new System.Windows.Interop.WindowInteropHelper(_ganttToolbar).Owner = (IntPtr)Application.HWND;
                _ganttToolbar.Closed += (s, args) => { _ganttToolbar = null; };

                var dispatcher = System.Windows.Application.Current?.Dispatcher;
                if (dispatcher == null) return;

                dispatcher.BeginInvoke(new Action(async () =>
                {
                    try
                    {
                        await System.Threading.Tasks.Task.Delay(100);
                        if (_ganttToolbar == null) return;

                        System.Windows.Forms.SendKeys.Send("{ESC}");
                        await System.Threading.Tasks.Task.Delay(50);
                        if (_ganttToolbar == null) return;

                        var window = Application.ActiveWindow;
                        if (window != null)
                        {
                            oPenEfficiency.UI.Lifecycle.ModelessUiCoordinator.PositionAtCursor(
                                _ganttToolbar, -130, -22, _dpiX, _dpiY);
                        }

                        _ganttToolbar.Show();
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "ThisAddIn.ShowGanttStylingToolbar.Dispatcher");
                    }
                }), System.Windows.Threading.DispatcherPriority.Normal);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ShowGanttStylingToolbar");
            }
        }

        public oPenEfficiency.UI.Panels.GanttPanelViewModel GetActiveGanttViewModel()
        {
            // 1. Check Sidebar TaskPanes
            var taskPane = GetSidebarForActiveWindow();
            var sidebar = (taskPane?.Control as SidebarTaskPaneHost)?.Sidebar;
            if (sidebar != null && sidebar.DynamicContentPanel != null)
            {
                foreach (var child in sidebar.DynamicContentPanel.Children)
                {
                    if (child is oPenEfficiency.UI.Panels.GanttChartPanel ganttPanel && ganttPanel.DataContext is oPenEfficiency.UI.Panels.GanttPanelViewModel vm)
                    {
                        return vm;
                    }
                }
            }

            // 2. Check Floating Windows
            try
            {
                foreach (System.Windows.Window win in System.Windows.Application.Current.Windows)
                {
                    if (win is oPenEfficiency.UI.Dialogs.GanttChartWindow gw)
                    {
                        foreach (var child in gw.MainGrid.Children)
                        {
                            if (child is oPenEfficiency.UI.Panels.GanttChartPanel gp && gp.DataContext is oPenEfficiency.UI.Panels.GanttPanelViewModel vm)
                            {
                                return vm;
                            }
                        }
                    }
                }
            }
            catch (InvalidOperationException)
            {
                // The WPF window collection can change while modeless windows close.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.GetActiveGanttViewModel.Windows");
            }

            return null;
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

    }
}
