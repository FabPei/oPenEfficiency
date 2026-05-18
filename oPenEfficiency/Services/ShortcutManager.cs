using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Input;
using oPenEfficiency.Utils;

namespace oPenEfficiency
{
    public static class ShortcutManager
    {
        private static void Log(string msg)
        {
            try
            {
                var tempPath = Path.Combine(Path.GetTempPath(), "oPenEfficiency");
                if (!Directory.Exists(tempPath)) Directory.CreateDirectory(tempPath);
                var path = Path.Combine(tempPath, "ShortcutDebug_oPenEfficiency.txt");
                File.AppendAllText(path, $"[{DateTime.Now:HH:mm:ss.fff}] {msg}\r\n");
            }
            catch
            {
                // Silently ignore logging failures
            }
        }

        private static Dictionary<int, Action> _hotkeyActions = new Dictionary<int, Action>();
        private static int _currentHotkeyId = 0;
        private static HiddenMessageWindow _messageWindow;
        private static SynchronizationContext _syncContext;
        private static bool _hotkeyFailureShown = false;

        public static void RegisterHandler(string featureId, Action action)
        {
            // We store the generic featureId mapping temporarily.
            // When LoadShortcuts is called, we will bind the configured keys to these actions.
            _featureActions[featureId] = action;
        }
        
        private static Dictionary<string, Action> _featureActions = new Dictionary<string, Action>();

        public static void ClearShortcut(string featureId)
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var configPath = Path.Combine(appData, "oPenEfficiency", "shortcuts.cfg");
            
            if (!File.Exists(configPath)) return;

            var lines = File.ReadAllLines(configPath).ToList();
            bool changed = false;
            for (int i = 0; i < lines.Count; i++)
            {
                if (lines[i].StartsWith(featureId + "="))
                {
                    lines.RemoveAt(i);
                    changed = true;
                    break;
                }
            }

            if (changed)
            {
                File.WriteAllLines(configPath, lines);
                Reload();
            }
        }

        public static Dictionary<string, string> GetShortcuts()
        {
            var results = new Dictionary<string, string>();
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var configPath = Path.Combine(appData, "oPenEfficiency", "shortcuts.cfg");
            
            if (File.Exists(configPath))
            {
                var lines = File.ReadAllLines(configPath);
                foreach (var line in lines)
                {
                    var parts = line.Split('=');
                    if (parts.Length == 2)
                    {
                        var keysList = parts[1].Split(',').Where(k => !string.IsNullOrWhiteSpace(k) && k != "None").ToList();
                        if (keysList.Count > 0)
                        {
                            results[parts[0]] = string.Join(" + ", keysList);
                        }
                    }
                }
            }
            return results;
        }

        public static void LoadShortcuts()
        {
            // Unregister existing hotkeys before loading new ones
            if (_messageWindow != null)
            {
                foreach (int id in _hotkeyActions.Keys)
                {
                    UnregisterHotKey(_messageWindow.Handle, id);
                }
            }
            _hotkeyActions.Clear();

            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var configPath = Path.Combine(appData, "oPenEfficiency", "shortcuts.cfg");

            if (!File.Exists(configPath)) return;

            // Re-initialize context if available
            if (System.Windows.Application.Current != null)
            {
                _syncContext = new System.Windows.Threading.DispatcherSynchronizationContext(System.Windows.Application.Current.Dispatcher);
            }

            var lines = File.ReadAllLines(configPath);
            foreach (var line in lines)
            {
                var parts = line.Split('=');
                if (parts.Length == 2)
                {
                    var featureId = parts[0];
                    if (!_featureActions.ContainsKey(featureId)) continue;
                    
                    var keysStr = parts[1].Split(',');
                    uint fsModifiers = 0;
                    uint vk = 0;

                    foreach (var k in keysStr)
                    {
                        if (Enum.TryParse(k, out Key key) && key != Key.None)
                        {
                            if (key == Key.LeftCtrl || key == Key.RightCtrl) fsModifiers |= MOD_CONTROL;
                            else if (key == Key.LeftShift || key == Key.RightShift) fsModifiers |= MOD_SHIFT;
                            else if (key == Key.LeftAlt || key == Key.RightAlt || key == Key.System) fsModifiers |= MOD_ALT;
                            else vk = (uint)KeyInterop.VirtualKeyFromKey(key);
                        }
                    }

                    if (vk != 0 && _messageWindow != null)
                    {
                        int id = Interlocked.Increment(ref _currentHotkeyId);

                        // Register Hotkey NO_REPEAT to avoid spamming
                        if (RegisterHotKey(_messageWindow.Handle, id, fsModifiers | MOD_NOREPEAT, vk))
                        {
                            _hotkeyActions[id] = _featureActions[featureId];
                            Log($"Registered Hotkey: {featureId} -> Modifiers: {fsModifiers}, VK: {vk}");
                        }
                        else
                        {
                            var errorCode = Marshal.GetLastWin32Error();
                            var errorMsg = $"Failed to register hotkey for '{featureId}'. Another application may be using this shortcut.";
                            Log(errorMsg + $" Error code: {errorCode}");
                            ExceptionLogger.Log(errorMsg, "ShortcutManager.RegisterHotKey");

                            // Show user notification on first failure
                            if (!_hotkeyFailureShown)
                            {
                                _hotkeyFailureShown = true;
                                System.Windows.Forms.MessageBox.Show(
                                    errorMsg + "\n\nPlease check the shortcut configuration or close conflicting applications.",
                                    "oPenEfficiency - Hotkey Registration Failed",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                            }
                        }
                    }
                }
            }
        }

        public static void StartHook()
        {
            Log("Starting ShortcutManager (RegisterHotKey)...");
            _syncContext = SynchronizationContext.Current;
            
            if (_messageWindow == null)
            {
                _messageWindow = new HiddenMessageWindow();
                _messageWindow.HotkeyPressed += MessageWindow_HotkeyPressed;
            }
            
            // Reload shortcuts now that window exists
            LoadShortcuts();
        }

        public static void StopHook()
        {
            Log("Stopping ShortcutManager...");
            if (_messageWindow != null)
            {
                foreach (int id in _hotkeyActions.Keys)
                {
                    UnregisterHotKey(_messageWindow.Handle, id);
                }
                _hotkeyActions.Clear();
                _messageWindow.DestroyHandle();
                _messageWindow = null;
            }
        }

        public static void Reload()
        {
            LoadShortcuts();
        }

        private static void MessageWindow_HotkeyPressed(int hotkeyId)
        {
            if (_hotkeyActions.ContainsKey(hotkeyId))
            {
                var handler = _hotkeyActions[hotkeyId];
                Log($"Hotkey {hotkeyId} triggered. Executing handler.");

                if (_syncContext != null)
                {
                    _syncContext.Post(_ => { try { handler?.Invoke(); } catch (Exception ex) { Log($"SyncContext Error: {ex}"); } }, null);
                }
                else if (System.Windows.Application.Current != null)
                {
                    System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() => { try { handler?.Invoke(); } catch (Exception ex) { Log($"Dispatcher Error: {ex}"); } }));
                }
                else
                {
                    System.Threading.Tasks.Task.Run(() => { try { handler?.Invoke(); } catch (Exception ex) { Log($"Task Error: {ex}"); } });
                }
            }
        }

        // native P/Invoke
        private const int WM_HOTKEY = 0x0312;
        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;
        private const uint MOD_NOREPEAT = 0x4000;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private class HiddenMessageWindow : NativeWindow
        {
            public event Action<int> HotkeyPressed;

            public HiddenMessageWindow()
            {
                CreateHandle(new CreateParams());
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_HOTKEY)
                {
                    int id = m.WParam.ToInt32();
                    HotkeyPressed?.Invoke(id);
                }
                base.WndProc(ref m);
            }
        }
    }
}

