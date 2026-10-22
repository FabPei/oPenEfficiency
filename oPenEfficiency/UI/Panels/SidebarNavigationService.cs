using System;
using System.Windows;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI.Panels
{
    internal sealed class SidebarNavigationService : IDisposable
    {
        private readonly FrameworkElement _owner;
        private readonly Func<PowerPointManager> _managerProvider;
        private readonly Action _settingsApplied;
        private UI.SpecialCharactersWindow _specialCharactersWindow;

        public SidebarNavigationService(
            FrameworkElement owner,
            Func<PowerPointManager> managerProvider,
            Action settingsApplied)
        {
            _owner = owner ?? throw new ArgumentNullException(nameof(owner));
            _managerProvider = managerProvider ?? throw new ArgumentNullException(nameof(managerProvider));
            _settingsApplied = settingsApplied ?? throw new ArgumentNullException(nameof(settingsApplied));
        }

        public void OpenApplyTextEditor()
        {
            new UI.ApplyTextWindow(_managerProvider()).Show();
        }

        public void ApplyTextFromMaster(bool useLastSelected)
        {
            var window = new UI.ApplyTextWindow(_managerProvider());
            try { window.ApplyMasterText(useLastSelected); }
            finally { window.Close(); }
        }

        public void OpenSettings(int sectionIndex)
        {
            var manager = _managerProvider();
            var window = new UI.SettingsWindow(manager);
            window.SelectSection(sectionIndex);
            window.SettingsApplied += (sender, args) => _settingsApplied();
            SetPowerPointOwner(window, manager);
            window.ShowDialog();
        }

        public void OpenTemplateManager()
        {
            var manager = _managerProvider();
            if (manager == null)
                throw new InvalidOperationException(
                    "PowerPoint is unavailable for Template Manager.");
            var window = new UI.Dialogs.TemplateManagerWindow(manager);
            SetPowerPointOwner(window, manager);
            window.ShowDialog();
        }

        public void OpenSpecialCharacters(string selection)
        {
            try
            {
                var ownerWindow = Window.GetWindow(_owner);
                if (_specialCharactersWindow == null)
                {
                    var manager = _managerProvider();
                    var window = new UI.SpecialCharactersWindow(manager);
                    _specialCharactersWindow = window;
                    window.Closed += (sender, args) =>
                    {
                        if (ReferenceEquals(_specialCharactersWindow, window))
                            _specialCharactersWindow = null;
                    };
                    SetPowerPointOwner(window, manager);
                }

                if (_specialCharactersWindow.WindowState == WindowState.Minimized)
                    _specialCharactersWindow.WindowState = WindowState.Normal;
                _specialCharactersWindow.SelectLanguage(selection);

                if (ownerWindow != null)
                {
                    double targetLeft = ownerWindow.Left - _specialCharactersWindow.Width - 10;
                    if (targetLeft < 0) targetLeft = ownerWindow.Left + ownerWindow.Width + 10;
                    _specialCharactersWindow.Left = targetLeft;
                    _specialCharactersWindow.Top = ownerWindow.Top;
                }
                else
                {
                    _specialCharactersWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                }

                _specialCharactersWindow.Show();
                _specialCharactersWindow.Activate();
                _specialCharactersWindow.Focus();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SidebarNavigationService.OpenSpecialCharacters");
                var failedWindow = _specialCharactersWindow;
                _specialCharactersWindow = null;
                try { failedWindow?.Close(); }
                catch (Exception closeEx)
                {
                    ExceptionLogger.Log(closeEx, "SidebarNavigationService.OpenSpecialCharacters.FailureCleanup");
                }
                throw;
            }
        }

        public void Dispose()
        {
            if (_specialCharactersWindow == null) return;
            try { _specialCharactersWindow.Close(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "SidebarNavigationService.Dispose"); }
            _specialCharactersWindow = null;
        }

        private static void SetPowerPointOwner(Window window, PowerPointManager manager)
        {
            var application = manager?.GetApplication();
            if (application != null)
                new System.Windows.Interop.WindowInteropHelper(window).Owner = (IntPtr)application.HWND;
        }
    }
}
