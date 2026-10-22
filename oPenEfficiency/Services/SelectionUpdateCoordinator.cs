using System;
using System.Windows.Threading;

namespace oPenEfficiency.Services
{
    internal sealed class SelectionUpdateCoordinator : IDisposable
    {
        private readonly DispatcherTimer _timer;
        private readonly Action _processSelection;
        private bool _pending;

        public SelectionUpdateCoordinator(
            int debounceMilliseconds,
            Action processSelection)
        {
            _processSelection = processSelection ?? throw new ArgumentNullException(nameof(processSelection));
            _timer = new DispatcherTimer(DispatcherPriority.ApplicationIdle)
            {
                Interval = TimeSpan.FromMilliseconds(debounceMilliseconds)
            };
            _timer.Tick += OnTick;
        }

        public void Schedule()
        {
            _pending = true;
            _timer.Stop();
            _timer.Start();
        }

        public void Dispose()
        {
            _timer.Stop();
            _timer.Tick -= OnTick;
            _pending = false;
        }

        private void OnTick(object sender, EventArgs args)
        {
            _timer.Stop();
            if (!_pending) return;
            _pending = false;
            _processSelection();
        }
    }
}
