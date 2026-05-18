using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Utils
{
    /// <summary>
    /// Global mouse hook for capturing mouse events outside the application window.
    /// Used for drag operations in Gantt chart and other interactive features.
    /// </summary>
    /// <remarks>
    /// <para><strong>Important:</strong> This global hook intercepts ALL mouse messages on the current thread,
    /// which has a side effect: <c>SendKeys.SendWait()</c> calls are intercepted and may not reach
    /// PowerPoint. This is why features like "Format Shape" cannot use keyboard shortcuts
    /// (e.g., <c>SendKeys.SendWait("^1")</c>) as a fallback - the hook catches them first.</para>
    /// <para>Features requiring native PowerPoint dialogs must use <c>CommandBars.ExecuteMso()</c>
    /// or direct COM manipulation instead of SendKeys.</para>
    /// </remarks>
    public class Win32MouseHook : IDisposable
    {
        // -------------------------------------------------------------
        // Win32 API Definitions
        // -------------------------------------------------------------
        private const int WH_MOUSE = 7;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_MOUSEMOVE = 0x0200;
        private const int WM_LBUTTONUP = 0x0202;
        private const int WM_RBUTTONUP = 0x0205;

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEHOOKSTRUCT
        {
            public POINT pt;
            public IntPtr hwnd;
            public uint wHitTestCode;
            public IntPtr dwExtraInfo;
        }

        public class MouseHookEventArgs : EventArgs
        {
            public int ScreenX { get; }
            public int ScreenY { get; }

            public MouseHookEventArgs(int screenX, int screenY)
            {
                ScreenX = screenX;
                ScreenY = screenY;
            }
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        // Delegate for the hook procedure
        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        // -------------------------------------------------------------
        // Hook State
        // -------------------------------------------------------------
        private IntPtr _hookId = IntPtr.Zero;
        private HookProc _proc;
        private bool _disposed = false;

        // Event fired when the left mouse button is released
        public event EventHandler LeftButtonUp;
        public event EventHandler RightButtonUp;
        public event EventHandler<MouseHookEventArgs> LeftButtonDown;

        // Callback for filtering mouse move events
        public Func<int, int, bool> MouseMoveFilter { get; set; }

        public Win32MouseHook()
        {
            // Maintain a reference to the delegate so the GC doesn't collect it
            _proc = HookCallback;
            if (!InstallHook())
            {
                throw new InvalidOperationException("Failed to install mouse hook. See debug output for details.");
            }
        }

        private bool InstallHook()
        {
            if (_hookId != IntPtr.Zero) return true; // Already installed

            uint threadId = GetCurrentThreadId();
            _hookId = SetWindowsHookEx(WH_MOUSE, _proc, IntPtr.Zero, threadId);

            if (_hookId == IntPtr.Zero)
            {
                var errorCode = Marshal.GetLastWin32Error();
                var message = $"Win32MouseHook failed to install hook. Error code: {errorCode}";
                Debug.WriteLine(message);
                ExceptionLogger.Log(message, "Win32MouseHook.InstallHook");
                return false;
            }
            else
            {
                Debug.WriteLine($"Win32MouseHook successfully installed on thread {threadId}.");
                return true;
            }
        }

        private void UninstallHook()
        {
            if (_hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
                Debug.WriteLine("Win32MouseHook uninstalled.");
            }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            // Prevent use after dispose
            if (_disposed)
            {
                return CallNextHookEx(_hookId, nCode, wParam, lParam);
            }

            // nCode < 0 means we must pass the message on without processing
            if (nCode >= 0)
            {
                int msg = wParam.ToInt32();

                if (msg == WM_LBUTTONDOWN)
                {
                    try
                    {
                        var hookStruct = (MOUSEHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MOUSEHOOKSTRUCT));
                        LeftButtonDown?.Invoke(this, new MouseHookEventArgs(hookStruct.pt.X, hookStruct.pt.Y));
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "Win32MouseHook.WM_LBUTTONDOWN");
                    }
                    // Always pass through WM_LBUTTONDOWN to PowerPoint for selection
                }
                else if (msg == WM_MOUSEMOVE)
                {
                    // Check if there's a filter callback
                    if (MouseMoveFilter != null)
                    {
                        try
                        {
                            var hookStruct = (MOUSEHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MOUSEHOOKSTRUCT));
                            bool suppress = MouseMoveFilter(hookStruct.pt.X, hookStruct.pt.Y);
                            if (suppress)
                            {
                                // Suppress the message by returning non-zero
                                return (IntPtr)1;
                            }
                        }
                        catch (Exception ex)
                        {
                            ExceptionLogger.Log(ex, "Win32MouseHook.WM_MOUSEMOVE");
                        }
                    }
                }
                else if (msg == WM_LBUTTONUP)
                {
                    LeftButtonUp?.Invoke(this, EventArgs.Empty);
                }
                else if (msg == WM_RBUTTONUP)
                {
                    RightButtonUp?.Invoke(this, EventArgs.Empty);
                }
            }

            // Always call the next hook in the chain unless we suppressed above
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        /// <summary>
        /// Disposes the mouse hook and unregisters it from the system.
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;

            _disposed = true;
            UninstallHook();

            // Clear event handlers to prevent memory leaks
            LeftButtonUp = null;
            RightButtonUp = null;
            LeftButtonDown = null;
            MouseMoveFilter = null;

            GC.SuppressFinalize(this);
        }

        ~Win32MouseHook()
        {
            if (!_disposed)
            {
                Debug.WriteLine("Win32MouseHook finalized without being disposed. Ensure Dispose() is called.");
            }
            UninstallHook();
        }
    }
}
