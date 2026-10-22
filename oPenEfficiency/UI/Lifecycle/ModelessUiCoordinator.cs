using System;
using System.Collections.Generic;
using System.Windows;
using Forms = System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI.Lifecycle
{
    internal sealed class ModelessUiCoordinator : IDisposable
    {
        private readonly Dictionary<string, Window> _windows =
            new Dictionary<string, Window>(StringComparer.Ordinal);
        private readonly Dictionary<string, Forms.Form> _forms =
            new Dictionary<string, Forms.Form>(StringComparer.Ordinal);
        private bool _disposed;

        internal int TrackedWindowCount => _windows.Count;

        public T ReplaceWindow<T>(string key, Func<T> factory, Action<T> position)
            where T : Window
        {
            ThrowIfDisposed();
            Close(key);
            return CreateWindow(key, factory, position);
        }

        public T GetOrCreateWindow<T>(
            string key,
            Func<T> factory,
            Action<T> update,
            Action<T> position)
            where T : Window
        {
            ThrowIfDisposed();
            if (_windows.TryGetValue(key, out var existing) && existing is T typed && typed.IsLoaded)
            {
                update?.Invoke(typed);
                typed.Activate();
                return typed;
            }

            return CreateWindow(key, factory, position);
        }

        public T TrackForm<T>(string key, T form) where T : Forms.Form
        {
            ThrowIfDisposed();
            if (_forms.TryGetValue(key, out var existing) && !ReferenceEquals(existing, form))
            {
                try { existing.Close(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ModelessUiCoordinator.TrackForm.ClosePrevious"); }
            }

            _forms[key] = form;
            form.FormClosed += (sender, args) =>
            {
                if (_forms.TryGetValue(key, out var tracked) && ReferenceEquals(tracked, form))
                    _forms.Remove(key);
            };
            return form;
        }

        public T TrackWindow<T>(string key, T window) where T : Window
        {
            ThrowIfDisposed();
            Close(key);
            _windows[key] = window;
            window.Closed += (sender, args) =>
            {
                if (_windows.TryGetValue(key, out var tracked) && ReferenceEquals(tracked, window))
                    _windows.Remove(key);
            };
            return window;
        }

        public T GetWindow<T>(string key) where T : Window
        {
            return _windows.TryGetValue(key, out var window) ? window as T : null;
        }

        public void Close(string key)
        {
            if (_windows.TryGetValue(key, out var window))
            {
                _windows.Remove(key);
                try { window.Close(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, $"ModelessUiCoordinator.CloseWindow.{key}"); }
            }

            if (_forms.TryGetValue(key, out var form))
            {
                _forms.Remove(key);
                try { form.Close(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, $"ModelessUiCoordinator.CloseForm.{key}"); }
            }
        }

        public static void PositionAtCursor(Window window, double offsetX, double offsetY, float dpiX, float dpiY)
        {
            var cursor = Forms.Cursor.Position;
            var screen = Forms.Screen.FromPoint(cursor);
            double scaleX = dpiX > 0 ? dpiX / 96.0 : 1.0;
            double scaleY = dpiY > 0 ? dpiY / 96.0 : 1.0;

            window.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            double width = !double.IsNaN(window.Width) ? window.Width : window.DesiredSize.Width;
            double height = !double.IsNaN(window.Height) ? window.Height : window.DesiredSize.Height;

            double workLeft = screen.WorkingArea.Left / scaleX;
            double workTop = screen.WorkingArea.Top / scaleY;
            double workRight = screen.WorkingArea.Right / scaleX;
            double workBottom = screen.WorkingArea.Bottom / scaleY;
            double requestedLeft = (cursor.X / scaleX) + offsetX;
            double requestedTop = (cursor.Y / scaleY) + offsetY;

            window.Left = Math.Max(workLeft, Math.Min(requestedLeft, workRight - width));
            window.Top = Math.Max(workTop, Math.Min(requestedTop, workBottom - height));
        }

        public static void PositionAtShape(
            Window toolbar,
            PowerPoint.Application application,
            PowerPoint.Shape shape,
            double offsetX,
            double offsetY,
            float dpiX,
            float dpiY)
        {
            var window = application?.ActiveWindow;
            if (window == null)
            {
                PositionAtCursor(toolbar, offsetX, offsetY, dpiX, dpiY);
                return;
            }

            int pixelX = window.PointsToScreenPixelsX(shape.Left + shape.Width);
            int pixelY = window.PointsToScreenPixelsY(shape.Top);
            toolbar.Left = (pixelX * 96.0 / dpiX) + offsetX;
            toolbar.Top = (pixelY * 96.0 / dpiY) + offsetY;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            foreach (var key in new List<string>(_windows.Keys)) Close(key);
            foreach (var key in new List<string>(_forms.Keys)) Close(key);
        }

        private T CreateWindow<T>(string key, Func<T> factory, Action<T> position)
            where T : Window
        {
            var window = factory();
            _windows[key] = window;
            window.Closed += (sender, args) =>
            {
                if (_windows.TryGetValue(key, out var tracked) && ReferenceEquals(tracked, window))
                    _windows.Remove(key);
            };
            position?.Invoke(window);
            window.Show();
            return window;
        }

        private void ThrowIfDisposed()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(ModelessUiCoordinator));
        }
    }
}
