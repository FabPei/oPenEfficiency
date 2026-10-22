using System;
using System.Collections.Generic;
using Microsoft.Office.Tools;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    internal sealed class SidebarPaneCoordinator : IDisposable
    {
        private readonly Dictionary<int, CustomTaskPane> _panes =
            new Dictionary<int, CustomTaskPane>();
        private readonly Action<CustomTaskPane> _removePane;
        private bool _disposed;

        public SidebarPaneCoordinator(Action<CustomTaskPane> removePane)
        {
            _removePane = removePane ?? throw new ArgumentNullException(nameof(removePane));
        }

        public int Count => _panes.Count;

        public bool TryGet(int windowHandle, out CustomTaskPane pane)
        {
            return _panes.TryGetValue(windowHandle, out pane);
        }

        public void Add(int windowHandle, CustomTaskPane pane)
        {
            if (_disposed) throw new ObjectDisposedException(nameof(SidebarPaneCoordinator));
            if (windowHandle == 0) throw new ArgumentOutOfRangeException(nameof(windowHandle));
            _panes[windowHandle] = pane ?? throw new ArgumentNullException(nameof(pane));
        }

        public void Remove(int windowHandle)
        {
            if (!_panes.TryGetValue(windowHandle, out var pane)) return;
            _panes.Remove(windowHandle);
            RemovePane(pane, "SidebarPaneCoordinator.Remove");
        }

        public void Prune(ISet<int> liveWindowHandles)
        {
            if (liveWindowHandles == null) throw new ArgumentNullException(nameof(liveWindowHandles));
            foreach (int staleHandle in new List<int>(_panes.Keys))
            {
                if (!liveWindowHandles.Contains(staleHandle)) Remove(staleHandle);
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            foreach (var pane in new List<CustomTaskPane>(_panes.Values))
                RemovePane(pane, "SidebarPaneCoordinator.Dispose");
            _panes.Clear();
        }

        private void RemovePane(CustomTaskPane pane, string context)
        {
            // Best-effort cleanup: during shutdown the VSTO CustomTaskPanes collection can be disposed
            // before this coordinator removes its panes (measured: ObjectDisposedException from
            // CustomTaskPaneCollectionImpl.Remove). The cleanup target is then already unavailable, which
            // is an expected cleanup miss; every other exception is still logged with its context.
            ExceptionPolicy.RunBestEffort(
                () => _removePane(pane),
                context,
                ex => ExceptionPolicy.IsExpectedComFailure(ex) || ex is ObjectDisposedException,
                ExceptionLogger.Log);
        }
    }
}
