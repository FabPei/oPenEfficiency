using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Input;
using oPenEfficiency.Services.Shortcuts;
using oPenEfficiency.Utils;

namespace oPenEfficiency
{
    public static class ShortcutManager
    {
        private const int WhKeyboardLl = 13;
        private const int WmKeyDown = 0x0100;
        private const int WmKeyUp = 0x0101;
        private const int WmSysKeyDown = 0x0104;
        private const int WmSysKeyUp = 0x0105;

        private const uint VkLShift = 0xA0;
        private const uint VkRShift = 0xA1;
        private const uint VkLControl = 0xA2;
        private const uint VkRControl = 0xA3;
        private const uint VkLMenu = 0xA4;
        private const uint VkRMenu = 0xA5;
        private const uint VkLWin = 0x5B;
        private const uint VkRWin = 0x5C;

        private static readonly object SyncRoot = new object();
        private static readonly Dictionary<string, Action> FeatureActions =
            new Dictionary<string, Action>(StringComparer.Ordinal);
        private static readonly Dictionary<string, ShortcutAssignmentState> Assignments =
            new Dictionary<string, ShortcutAssignmentState>(StringComparer.Ordinal);
        private static readonly HashSet<uint> PhysicalKeysDown = new HashSet<uint>();
        private static readonly Dictionary<string, int> NormalizedKeyCounts =
            new Dictionary<string, int>(StringComparer.Ordinal);
        private static readonly HashSet<string> TriggeredSignatures =
            new HashSet<string>(StringComparer.Ordinal);

        private static LowLevelKeyboardProc _hookProc;
        private static IntPtr _hookHandle;
        private static SynchronizationContext _syncContext;
        private static Func<bool> _eligiblePresentationWindow;
        private static bool _started;
        private static int _modalCaptureCount;

        internal static bool IsHookInstalledForDiagnostics
        {
            get
            {
                lock (SyncRoot) return _started && _hookHandle != IntPtr.Zero;
            }
        }

        public static void RegisterHandler(string featureId, Action action)
        {
            if (string.IsNullOrWhiteSpace(featureId) || action == null) return;
            lock (SyncRoot)
                FeatureActions[featureId] = action;
        }

        public static bool StartHook(Func<bool> eligiblePresentationWindow = null)
        {
            lock (SyncRoot)
            {
                _eligiblePresentationWindow = eligiblePresentationWindow;
                _syncContext = SynchronizationContext.Current;
                _started = true;
            }

            bool installed = EnsureHookInstalled();
            LoadShortcuts();
            Log(installed
                ? "Started application-scoped shortcut observer."
                : "Shortcut observer failed to start.");
            return installed;
        }

        public static void StopHook()
        {
            IntPtr handle;
            lock (SyncRoot)
            {
                _started = false;
                handle = _hookHandle;
                _hookHandle = IntPtr.Zero;
                Assignments.Clear();
                PhysicalKeysDown.Clear();
                NormalizedKeyCounts.Clear();
                TriggeredSignatures.Clear();
            }

            if (handle != IntPtr.Zero && !UnhookWindowsHookEx(handle))
            {
                ExceptionLogger.Log(
                    "Failed to remove the shortcut keyboard observer. Error code: " +
                    Marshal.GetLastWin32Error(),
                    "ShortcutManager.StopHook");
            }
        }

        public static IDisposable SuspendActivationForModalCapture()
        {
            Interlocked.Increment(ref _modalCaptureCount);
            return new ModalCaptureScope();
        }

        public static void LoadShortcuts()
        {
            var classified = ShortcutPolicy.ClassifyLegacyAssignments(ReadRawAssignments());
            lock (SyncRoot)
            {
                Assignments.Clear();
                foreach (var assignment in classified)
                {
                    if (!FeatureActions.ContainsKey(assignment.FeatureId))
                    {
                        assignment.IsActive = false;
                        if (string.IsNullOrWhiteSpace(assignment.AttentionReason))
                        {
                            assignment.AttentionReason =
                                "The command handler is not available in this PowerPoint session.";
                        }
                    }

                    Assignments[assignment.FeatureId] = assignment;
                }
            }
        }

        public static void Reload() => LoadShortcuts();

        public static IReadOnlyList<ShortcutAssignmentState> GetAssignments()
        {
            lock (SyncRoot)
            {
                return Assignments.Values
                    .Select(Clone)
                    .OrderBy(assignment => assignment.FeatureId, StringComparer.Ordinal)
                    .ToArray();
            }
        }

        public static ShortcutAssignmentState GetAssignment(string featureId)
        {
            lock (SyncRoot)
            {
                return featureId != null && Assignments.TryGetValue(featureId, out var assignment)
                    ? Clone(assignment)
                    : null;
            }
        }

        public static Dictionary<string, string> GetShortcuts()
        {
            lock (SyncRoot)
            {
                return Assignments.Values
                    .Where(assignment => assignment.IsActive && assignment.Chord != null)
                    .ToDictionary(
                        assignment => assignment.FeatureId,
                        assignment => assignment.Chord.DisplayText,
                        StringComparer.Ordinal);
            }
        }

        public static ShortcutValidationResult ValidateShortcut(
            string featureId,
            IEnumerable<string> rawKeys,
            bool riskConfirmed = false)
        {
            ShortcutAssignmentState[] existing;
            lock (SyncRoot)
                existing = Assignments.Values.Select(Clone).ToArray();

            return ShortcutPolicy.Validate(
                rawKeys,
                featureId,
                existing,
                riskConfirmed);
        }

        public static bool TryReplaceShortcut(
            string featureId,
            IEnumerable<string> rawKeys,
            bool riskConfirmed,
            out ShortcutValidationResult validation,
            out string error)
        {
            validation = ValidateShortcut(featureId, rawKeys, riskConfirmed);
            error = validation.Error;
            if (!validation.IsValid) return false;

            var previous = GetAssignment(featureId);
            var replacement = new ShortcutAssignmentState
            {
                FeatureId = featureId,
                RawValue = validation.Chord.StorageText,
                Chord = validation.Chord,
                IsActive = true
            };

            bool replaced = ShortcutReplacementTransaction.TryReplace(
                previous,
                replacement,
                candidate => IsRuntimeReadyFor(candidate.FeatureId) && EnsureHookInstalled(),
                candidate => WriteAssignment(candidate.FeatureId, candidate.Chord.StorageText),
                prior => LoadShortcuts(),
                out error);

            if (replaced)
                LoadShortcuts();
            return replaced;
        }

        public static bool ClearShortcut(string featureId)
        {
            try
            {
                WriteAssignment(featureId, null);
                LoadShortcuts();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ShortcutManager.ClearShortcut");
                return false;
            }
        }

        private static bool IsRuntimeReadyFor(string featureId)
        {
            lock (SyncRoot)
            {
                return _started &&
                       !string.IsNullOrWhiteSpace(featureId) &&
                       FeatureActions.ContainsKey(featureId);
            }
        }

        private static bool EnsureHookInstalled()
        {
            lock (SyncRoot)
            {
                if (!_started) return false;
                if (_hookHandle != IntPtr.Zero) return true;

                _hookProc = HookCallback;
                using (var process = Process.GetCurrentProcess())
                using (var module = process.MainModule)
                {
                    IntPtr moduleHandle = module == null
                        ? IntPtr.Zero
                        : GetModuleHandle(module.ModuleName);
                    _hookHandle = SetWindowsHookEx(
                        WhKeyboardLl,
                        _hookProc,
                        moduleHandle,
                        0);
                }

                if (_hookHandle == IntPtr.Zero)
                {
                    ExceptionLogger.Log(
                        "Failed to install the shortcut keyboard observer. Error code: " +
                        Marshal.GetLastWin32Error(),
                        "ShortcutManager.StartHook");
                    return false;
                }

                return true;
            }
        }

        private static IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (code >= 0)
                {
                    int message = wParam.ToInt32();
                    bool isDown = message == WmKeyDown || message == WmSysKeyDown;
                    bool isUp = message == WmKeyUp || message == WmSysKeyUp;
                    if (isDown || isUp)
                    {
                        var data = Marshal.PtrToStructure<KeyboardHookData>(lParam);
                        ProcessPhysicalKey(data.VirtualKey, isDown);
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ShortcutManager.HookCallback");
            }

            // Observation only: never suppress Windows or PowerPoint keystrokes.
            return CallNextHookEx(_hookHandle, code, wParam, lParam);
        }

        private static void ProcessPhysicalKey(uint virtualKey, bool isDown)
        {
            string normalizedKey = NormalizeVirtualKey(virtualKey) ?? "VK_" + virtualKey;
            List<KeyValuePair<string, Action>> matches = null;

            lock (SyncRoot)
            {
                if (isDown)
                {
                    if (!PhysicalKeysDown.Add(virtualKey)) return;
                    NormalizedKeyCounts[normalizedKey] =
                        NormalizedKeyCounts.TryGetValue(normalizedKey, out int count)
                            ? count + 1
                            : 1;

                    var pressed = new HashSet<string>(
                        NormalizedKeyCounts.Where(pair => pair.Value > 0).Select(pair => pair.Key),
                        StringComparer.Ordinal);

                    matches = Assignments.Values
                        .Where(assignment =>
                            assignment.IsActive &&
                            assignment.Chord != null &&
                            !TriggeredSignatures.Contains(assignment.Chord.Signature) &&
                            pressed.SetEquals(assignment.Chord.Keys) &&
                            FeatureActions.ContainsKey(assignment.FeatureId))
                        .Select(assignment =>
                            new KeyValuePair<string, Action>(
                                assignment.FeatureId,
                                FeatureActions[assignment.FeatureId]))
                        .ToList();

                    foreach (var match in matches)
                    {
                        var assignment = Assignments[match.Key];
                        TriggeredSignatures.Add(assignment.Chord.Signature);
                    }
                }
                else
                {
                    if (!PhysicalKeysDown.Remove(virtualKey)) return;
                    if (NormalizedKeyCounts.TryGetValue(normalizedKey, out int count))
                    {
                        if (count <= 1) NormalizedKeyCounts.Remove(normalizedKey);
                        else NormalizedKeyCounts[normalizedKey] = count - 1;
                    }
                    TriggeredSignatures.Clear();
                }
            }

            if (matches == null || matches.Count == 0 || !IsPowerPointForeground())
                return;

            foreach (var match in matches)
                ScheduleEligibleExecution(match.Key, match.Value);
        }

        private static void ScheduleEligibleExecution(string featureId, Action handler)
        {
            Action execute = () =>
            {
                if (!CanExecuteNow(featureId, handler)) return;
                try
                {
                    handler();
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "ShortcutManager.Execute." + featureId);
                }
            };

            var context = _syncContext;
            if (context != null)
            {
                context.Post(_ => execute(), null);
                return;
            }

            var application = System.Windows.Application.Current;
            if (application != null && application.Dispatcher != null)
                application.Dispatcher.BeginInvoke(execute);
        }

        private static bool CanExecuteNow(string featureId, Action handler)
        {
            Func<bool> eligible;
            bool initialized;
            bool handlerInitialized;
            lock (SyncRoot)
            {
                initialized = _started;
                handlerInitialized =
                    handler != null &&
                    FeatureActions.TryGetValue(featureId, out var registered) &&
                    ReferenceEquals(registered, handler);
                eligible = _eligiblePresentationWindow;
            }

            bool eligibleWindow = false;
            try
            {
                eligibleWindow = eligible != null && eligible();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ShortcutManager.CheckEligibility");
            }

            return ShortcutActivationPolicy.CanExecute(
                IsPowerPointForeground(),
                eligibleWindow,
                initialized,
                handlerInitialized,
                Volatile.Read(ref _modalCaptureCount) > 0);
        }

        private static bool IsPowerPointForeground()
        {
            IntPtr foreground = GetForegroundWindow();
            if (foreground == IntPtr.Zero) return false;
            GetWindowThreadProcessId(foreground, out uint processId);
            return processId == (uint)Process.GetCurrentProcess().Id;
        }

        private static string NormalizeVirtualKey(uint virtualKey)
        {
            switch (virtualKey)
            {
                case VkLControl: return "Ctrl";
                case VkRControl: return "Ctrl";
                case VkLMenu: return "Alt";
                case VkRMenu: return "Alt";
                case VkLShift: return "Shift";
                case VkRShift: return "Shift";
                case VkLWin: return "Win";
                case VkRWin: return "Win";
                default:
                    var key = KeyInterop.KeyFromVirtualKey((int)virtualKey);
                    return ShortcutPolicy.NormalizeKey(key.ToString());
            }
        }

        private static IEnumerable<KeyValuePair<string, string>> ReadRawAssignments()
        {
            string path = GetConfigPath();
            if (!File.Exists(path)) return Enumerable.Empty<KeyValuePair<string, string>>();

            var assignments = new List<KeyValuePair<string, string>>();
            foreach (string line in File.ReadAllLines(path))
            {
                int separator = line.IndexOf('=');
                if (separator <= 0) continue;
                string featureId = line.Substring(0, separator).Trim();
                string rawValue = line.Substring(separator + 1);
                if (featureId.Length > 0)
                    assignments.Add(new KeyValuePair<string, string>(featureId, rawValue));
            }
            return assignments;
        }

        private static void WriteAssignment(string featureId, string storageValue)
        {
            if (string.IsNullOrWhiteSpace(featureId))
                throw new ArgumentException("A feature ID is required.", nameof(featureId));

            string path = GetConfigPath();
            string folder = Path.GetDirectoryName(path);
            Directory.CreateDirectory(folder);

            var source = File.Exists(path)
                ? File.ReadAllLines(path).ToList()
                : new List<string>();
            var output = new List<string>();
            bool written = false;

            foreach (string line in source)
            {
                int separator = line.IndexOf('=');
                bool sameFeature = separator > 0 &&
                    string.Equals(
                        line.Substring(0, separator).Trim(),
                        featureId,
                        StringComparison.Ordinal);
                if (!sameFeature)
                {
                    output.Add(line);
                    continue;
                }

                if (!written && storageValue != null)
                {
                    output.Add(featureId + "=" + storageValue);
                    written = true;
                }
            }

            if (!written && storageValue != null)
                output.Add(featureId + "=" + storageValue);

            string temporaryPath = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            File.WriteAllLines(temporaryPath, output);
            try
            {
                if (File.Exists(path))
                {
                    string backupPath = temporaryPath + ".bak";
                    File.Replace(temporaryPath, path, backupPath, true);
                    try { File.Delete(backupPath); }
                    catch { /* A stale backup is harmless after the replacement file was saved. */ }
                }
                else
                {
                    File.Move(temporaryPath, path);
                }
            }
            finally
            {
                try
                {
                    if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
                }
                catch { /* Preserve the primary shortcut-save error if restoring the backup fails. */ }
            }
        }

        private static string GetConfigPath()
        {
            string appData = Environment.GetFolderPath(
                Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appData, "oPenEfficiency", "shortcuts.cfg");
        }

        private static ShortcutAssignmentState Clone(ShortcutAssignmentState assignment) =>
            new ShortcutAssignmentState
            {
                FeatureId = assignment.FeatureId,
                RawValue = assignment.RawValue,
                Chord = assignment.Chord == null
                    ? null
                    : new ShortcutChord(assignment.Chord.Keys),
                IsActive = assignment.IsActive,
                AttentionReason = assignment.AttentionReason
            };

        private static void Log(string message)
        {
            try
            {
                string folder = Path.Combine(Path.GetTempPath(), "oPenEfficiency");
                Directory.CreateDirectory(folder);
                File.AppendAllText(
                    Path.Combine(folder, "ShortcutDebug_oPenEfficiency.txt"),
                    $"[{DateTime.Now:HH:mm:ss.fff}] {message}{Environment.NewLine}");
            }
            catch { /* Shortcut cleanup is best-effort during settings migration. */ }
        }

        private sealed class ModalCaptureScope : IDisposable
        {
            private int _disposed;

            public void Dispose()
            {
                if (Interlocked.Exchange(ref _disposed, 1) == 0)
                    Interlocked.Decrement(ref _modalCaptureCount);
            }
        }

        private delegate IntPtr LowLevelKeyboardProc(
            int code,
            IntPtr wParam,
            IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct KeyboardHookData
        {
            public uint VirtualKey;
            public uint ScanCode;
            public uint Flags;
            public uint Time;
            public UIntPtr ExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(
            int hookId,
            LowLevelKeyboardProc callback,
            IntPtr module,
            uint threadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hook);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(
            IntPtr hook,
            int code,
            IntPtr wParam,
            IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string moduleName);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(
            IntPtr window,
            out uint processId);
    }
}
