using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using oPenEfficiency.Utils;
using PP = global::Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Excel.Services
{
    /// <summary>
    /// Attaches to a running PowerPoint instance.
    ///
    /// Why not just Marshal.GetActiveObject? PowerPoint only registers itself in the
    /// Running Object Table once a presentation has been *saved* (given a file moniker),
    /// so GetActiveObject fails for a brand-new unsaved "Presentation1". To be robust we
    /// find the PowerPoint window and pull the native object model out of it via
    /// AccessibleObjectFromWindow(OBJID_NATIVEOM), which works regardless of save state.
    /// </summary>
    public static class PowerPointConnector
    {
        private const uint OBJID_NATIVEOM = 0xFFFFFFF0;
        private static readonly Guid IID_IDispatch =
            new Guid("00020400-0000-0000-C000-000000000046");

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(
            IntPtr hwnd, uint id, ref Guid iid,
            [MarshalAs(UnmanagedType.IUnknown)] ref object ppvObject);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        /// <summary>
        /// Returns a running PowerPoint Application, or null if none is running.
        /// Tries the window-attach route first (works for unsaved presentations),
        /// then falls back to the Running Object Table.
        /// </summary>
        public static PP.Application GetRunning()
        {
            bool ambiguous;
            return GetRunning(out ambiguous);
        }

        public static PP.Application GetRunning(out bool ambiguous)
        {
            ambiguous = false;
            List<uint> processIds = EnumeratePowerPointProcessIds()
                .Where(processId => processId != 0)
                .Distinct()
                .ToList();
            PowerPointInstanceSelection processSelection = PowerPointInstancePolicy.SelectProcessIds(processIds);
            List<PowerPointWindowCandidate> candidates =
                processSelection == PowerPointInstanceSelection.Single
                    ? EnumerateWindowCandidates()
                    : new List<PowerPointWindowCandidate>();
            PowerPointInstanceSelection selection = PowerPointInstancePolicy.Select(candidates);

            PowerPointAttachRoute route = PowerPointInstancePolicy.SelectRoute(processSelection, selection);
            if (route == PowerPointAttachRoute.Ambiguous)
            {
                ambiguous = true;
                return null;
            }
            if (route == PowerPointAttachRoute.None) return null;

            if (route == PowerPointAttachRoute.VisibleWindow)
            {
                var app = GetFromWindow(candidates[0].WindowHandle);
                if (app != null) return app;
                return GetFromRunningObjectTable(candidates[0].ProcessId);
            }

            // Exactly one live POWERPNT process, but no enumerable visible PPTFrameClass window
            // (hidden or off-screen frame). The single process is still the unambiguous target, so
            // the PID-checked Running Object Table route is attempted before failing closed:
            // GetFromRunningObjectTable re-verifies that the automation object belongs to that very
            // process, so this can never attach to a foreign deck. Before DL-CONNECTOR-ROT-001 this
            // state returned a silent null even though an attachable ROT entry for that process
            // existed, which made the connector depend on an accident of window state.
            PP.Application fromRunningObjectTable = GetFromRunningObjectTable(processIds[0]);
            if (fromRunningObjectTable != null) return fromRunningObjectTable;

            ExceptionLogger.Log(
                "PowerPointConnector.GetRunning: exactly one POWERPNT process is running (pid " +
                processIds[0] + "), but no visible PPTFrameClass window could be enumerated and the " +
                "PID-checked Running Object Table entry did not resolve to that process. No PowerPoint " +
                "instance was attached (fail closed).",
                "PowerPointConnector.GetRunning.NoVisibleFrame");
            return null;
        }

        private static IEnumerable<uint> EnumeratePowerPointProcessIds()
        {
            foreach (Process process in Process.GetProcessesByName("POWERPNT"))
            {
                using (process)
                {
                    bool exited;
                    try { exited = process.HasExited; }
                    catch (InvalidOperationException) { continue; }
                    catch (System.ComponentModel.Win32Exception) { continue; }
                    if (!exited) yield return unchecked((uint)process.Id);
                }
            }
        }

        private static PP.Application GetFromWindow(IntPtr frame)
        {
            try
            {
                // The slide document window carries the native object model.
                IntPtr pane = FindChildWindow(frame, "paneClassDC");
                if (pane == IntPtr.Zero) return null;

                Guid iid = IID_IDispatch;
                object obj = null;
                int hr = AccessibleObjectFromWindow(pane, OBJID_NATIVEOM, ref iid, ref obj);
                if (hr != 0 || obj == null) return null;

                dynamic native = obj;
                // The returned object is typically a DocumentWindow; get its Application.
                try
                {
                    PP.Application application = (PP.Application)native.Application;
                    if (!ReferenceEquals(application, obj) && Marshal.IsComObject(obj))
                        Marshal.ReleaseComObject(obj);
                    return application;
                }
                catch (COMException)
                {
                    // Some NativeOM variants expose the application object directly.
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "PowerPointConnector.GetFromWindow.NativeApplication");
                }

                try { return (PP.Application)native; }
                catch (COMException)
                {
                    // The window did not expose a usable PowerPoint application.
                }
                catch (InvalidCastException)
                {
                    // NativeOM returned another Office automation object.
                }
                if (Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj);
                return null;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PowerPointConnector.GetFromWindow");
                return null;
            }
        }

        private static List<PowerPointWindowCandidate> EnumerateWindowCandidates()
        {
            var candidates = new List<PowerPointWindowCandidate>();
            EnumWindows((hWnd, lParam) =>
            {
                if (IsWindowVisible(hWnd) && GetClass(hWnd) == "PPTFrameClass")
                {
                    uint processId;
                    GetWindowThreadProcessId(hWnd, out processId);
                    if (processId != 0)
                        candidates.Add(new PowerPointWindowCandidate(hWnd, processId));
                }
                return true;
            }, IntPtr.Zero);
            return candidates
                .GroupBy(candidate => candidate.ProcessId)
                .Select(group => group.OrderBy(candidate => candidate.WindowHandle.ToInt64()).First())
                .OrderBy(candidate => candidate.ProcessId)
                .ToList();
        }

        private static IntPtr FindChildWindow(IntPtr parent, string className)
        {
            IntPtr found = IntPtr.Zero;
            EnumChildWindows(parent, (hWnd, lParam) =>
            {
                if (GetClass(hWnd) == className)
                {
                    found = hWnd;
                    return false; // stop
                }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        private static string GetClass(IntPtr hWnd)
        {
            var sb = new StringBuilder(256);
            GetClassName(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private static PP.Application GetFromRunningObjectTable(uint expectedProcessId)
        {
            PP.Application application = null;
            try
            {
                application = (PP.Application)Marshal.GetActiveObject("PowerPoint.Application");
                uint processId;
                GetWindowThreadProcessId(new IntPtr(application.HWND), out processId);
                if (processId == expectedProcessId) return application;
                Marshal.ReleaseComObject(application);
                return null;
            }
            catch (COMException)
            {
                if (application != null) Marshal.ReleaseComObject(application);
                return null;
            }
            catch (Exception ex)
            {
                if (application != null && Marshal.IsComObject(application)) Marshal.ReleaseComObject(application);
                ExceptionLogger.Log(ex, "PowerPointConnector.GetRunning.ROT");
                return null;
            }
        }
    }

    internal sealed class PowerPointWindowCandidate
    {
        internal PowerPointWindowCandidate(IntPtr windowHandle, uint processId)
        {
            WindowHandle = windowHandle;
            ProcessId = processId;
        }

        internal IntPtr WindowHandle { get; }
        internal uint ProcessId { get; }
    }

    internal enum PowerPointInstanceSelection
    {
        None,
        Single,
        Ambiguous
    }

    /// <summary>How the connector reached (or refused) a PowerPoint instance.</summary>
    internal enum PowerPointAttachRoute
    {
        None,

        Ambiguous,

        VisibleWindow,

        RunningObjectTableFallback
    }
    internal static class PowerPointInstancePolicy
    {
        internal static PowerPointInstanceSelection SelectProcessIds(IEnumerable<uint> processIds)
        {
            int count = (processIds ?? Enumerable.Empty<uint>())
                .Where(processId => processId != 0)
                .Distinct()
                .Take(2)
                .Count();
            if (count == 0) return PowerPointInstanceSelection.None;
            return count == 1 ? PowerPointInstanceSelection.Single : PowerPointInstanceSelection.Ambiguous;
        }

        internal static PowerPointInstanceSelection Select(IEnumerable<PowerPointWindowCandidate> candidates)
        {
            return SelectProcessIds((candidates ?? Enumerable.Empty<PowerPointWindowCandidate>())
                .Select(candidate => candidate.ProcessId));
        }
        /// <summary>
        /// Decides how to attach, given what the process list and the visible window candidates
        /// showed. A single live process without an enumerable visible frame still routes to the
        /// PID-checked Running Object Table fallback instead of the silent null that
        /// DL-CONNECTOR-ROT-001 reported as a defect.
        /// </summary>
        internal static PowerPointAttachRoute SelectRoute(
            PowerPointInstanceSelection processSelection,
            PowerPointInstanceSelection windowSelection)
        {
            if (processSelection == PowerPointInstanceSelection.Ambiguous ||
                (processSelection == PowerPointInstanceSelection.Single &&
                 windowSelection == PowerPointInstanceSelection.Ambiguous))
            {
                return PowerPointAttachRoute.Ambiguous;
            }
            if (processSelection == PowerPointInstanceSelection.None) return PowerPointAttachRoute.None;
            return windowSelection == PowerPointInstanceSelection.Single
                ? PowerPointAttachRoute.VisibleWindow
                : PowerPointAttachRoute.RunningObjectTableFallback;
        }
    }
}
