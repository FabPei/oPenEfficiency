using System;
using System.Linq;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Routes global mouse events to the Gantt interaction machinery. Hook callbacks must stay
    /// COM-free: synchronous PowerPoint COM inside a WH_MOUSE callback runs while
    /// WM_LBUTTONDOWN is still in dispatch, and STA re-entrancy there can reorder the pending
    /// WM_LBUTTONUP and silently swallow clicks (PQ-015). All COM work is deferred to the
    /// dispatcher and gated on an open Gantt window.
    /// </summary>
    internal sealed class GanttInteractionCoordinator : IDisposable
    {
        private readonly GanttInteractionManager _interactionManager;
        private readonly Win32MouseHook _mouseHook;
        private readonly Action<int, int> _leftButtonDown;
        private readonly EventHandler<Win32MouseHook.MouseHookEventArgs> _rightButtonUp;
        private bool _disposed;

        public GanttInteractionCoordinator(
            PowerPointManager manager,
            Action<int, int> leftButtonDown,
            EventHandler<Win32MouseHook.MouseHookEventArgs> rightButtonUp,
            Win32MouseHook sharedMouseHook)
        {
            _leftButtonDown = leftButtonDown ?? throw new ArgumentNullException(nameof(leftButtonDown));
            _rightButtonUp = rightButtonUp ?? throw new ArgumentNullException(nameof(rightButtonUp));
            _interactionManager = new GanttInteractionManager(
                manager ?? throw new ArgumentNullException(nameof(manager)));
            // One shared thread hook per process, owned by ThisAddIn (PQ-016).
            _mouseHook = sharedMouseHook
                ?? throw new ArgumentNullException(nameof(sharedMouseHook));
            _mouseHook.LeftButtonDown += OnLeftButtonDown;
            _mouseHook.LeftButtonUp += OnLeftButtonUp;
            _mouseHook.RightButtonUp += OnHookRightButtonUp;
            _mouseHook.MouseMoveFilter += _interactionManager.OnMouseMove;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _mouseHook.LeftButtonDown -= OnLeftButtonDown;
            _mouseHook.LeftButtonUp -= OnLeftButtonUp;
            _mouseHook.RightButtonUp -= OnHookRightButtonUp;
            _mouseHook.MouseMoveFilter -= _interactionManager.OnMouseMove;
            _interactionManager.Dispose();
        }

        /// <summary>Cheap, COM-free check: only probe the mouse while a Gantt window is live.</summary>
        internal static bool AnyGanttWindowOpen()
        {
            var windows = System.Windows.Application.Current?.Windows;
            if (windows == null) return false;
            return windows.OfType<UI.Dialogs.GanttChartWindow>().Any(w => w.IsLoaded);
        }

        private static void DeferToUi(Action work)
        {
            var dispatcher = System.Windows.Application.Current?.Dispatcher;
            if (dispatcher == null) return;
            dispatcher.BeginInvoke(work, System.Windows.Threading.DispatcherPriority.Input);
        }

        private void OnLeftButtonDown(object sender, Win32MouseHook.MouseHookEventArgs args)
        {
            if (_disposed || !AnyGanttWindowOpen()) return;

            int screenX = args.ScreenX;
            int screenY = args.ScreenY;
            DeferToUi(() =>
            {
                if (_disposed || !Win32MouseHook.IsLeftMouseButtonDown()) return;
                _interactionManager.OnMouseDown(screenX, screenY);
                _leftButtonDown(screenX, screenY);
            });
        }

        private void OnLeftButtonUp(object sender, EventArgs args)
        {
            // State-only reset; no COM in the hook callback.
            _interactionManager.OnMouseUp();
        }

        private void OnHookRightButtonUp(object sender, Win32MouseHook.MouseHookEventArgs args)
        {
            if (_disposed || !AnyGanttWindowOpen()) return;

            var handler = _rightButtonUp;
            if (handler == null) return;
            DeferToUi(() => handler(this, args));
        }
    }
}
