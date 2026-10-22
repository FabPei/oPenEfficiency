using System;
using System.Collections.Generic;
using System.Linq;

namespace oPenEfficiency.Services.Shortcuts
{
    public sealed class ShortcutChord
    {
        public ShortcutChord(IEnumerable<string> keys)
        {
            Keys = (keys ?? Enumerable.Empty<string>()).ToArray();
        }

        public IReadOnlyList<string> Keys { get; }
        public string Signature => string.Join("|", Keys);
        public string StorageText => string.Join(",", Keys);
        public string DisplayText => string.Join(" + ", Keys.Select(ShortcutPolicy.GetDisplayName));
    }

    public sealed class ShortcutValidationResult
    {
        public bool IsValid { get; set; }
        public bool RequiresRiskConfirmation { get; set; }
        public string Error { get; set; }
        public string Warning { get; set; }
        public string ConflictingFeatureId { get; set; }
        public ShortcutChord Chord { get; set; }
    }

    public sealed class ShortcutAssignmentState
    {
        public string FeatureId { get; set; }
        public string RawValue { get; set; }
        public ShortcutChord Chord { get; set; }
        public bool IsActive { get; set; }
        public string AttentionReason { get; set; }

        public bool NeedsAttention => !string.IsNullOrWhiteSpace(AttentionReason);
        public string ShortcutText =>
            Chord?.DisplayText ??
            ShortcutPolicy.FormatLegacyValue(RawValue);
    }

    public static class ShortcutPolicy
    {
        private static readonly Dictionary<string, string> Aliases =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["Control"] = "Ctrl",
                ["LeftCtrl"] = "Ctrl",
                ["RightCtrl"] = "Ctrl",
                ["Ctrl"] = "Ctrl",
                ["LeftAlt"] = "Alt",
                ["RightAlt"] = "Alt",
                ["Alt"] = "Alt",
                ["LeftShift"] = "Shift",
                ["RightShift"] = "Shift",
                ["Shift"] = "Shift",
                ["LWin"] = "Win",
                ["RWin"] = "Win",
                ["Win"] = "Win",
                ["Enter"] = "Return",
                ["Esc"] = "Escape"
            };

        private static readonly HashSet<string> NamedSupportedKeys =
            new HashSet<string>(new[]
            {
                "Ctrl", "Alt", "Shift", "Win",
                "Back", "Tab", "Clear", "Return", "Pause", "CapsLock", "Escape", "Space",
                "PageUp", "PageDown", "End", "Home", "Left", "Up", "Right", "Down",
                "Select", "Print", "Execute", "Snapshot", "Insert", "Delete", "Help",
                "Apps", "Sleep", "Multiply", "Add", "Separator", "Subtract", "Decimal",
                "Divide", "NumLock", "Scroll",
                "BrowserBack", "BrowserForward", "BrowserRefresh", "BrowserStop",
                "BrowserSearch", "BrowserFavorites", "BrowserHome",
                "VolumeMute", "VolumeDown", "VolumeUp", "MediaNextTrack",
                "MediaPreviousTrack", "MediaStop", "MediaPlayPause", "LaunchMail",
                "SelectMedia", "LaunchApplication1", "LaunchApplication2",
                "Oem1", "OemPlus", "OemComma", "OemMinus", "OemPeriod", "Oem2",
                "Oem3", "Oem4", "Oem5", "Oem6", "Oem7", "Oem8", "Oem102"
            }, StringComparer.Ordinal);

        private static readonly HashSet<string> TextKeys =
            new HashSet<string>(StringComparer.Ordinal);

        static ShortcutPolicy()
        {
            for (char letter = 'A'; letter <= 'Z'; letter++)
            {
                NamedSupportedKeys.Add(letter.ToString());
                TextKeys.Add(letter.ToString());
            }

            for (int digit = 0; digit <= 9; digit++)
            {
                NamedSupportedKeys.Add("D" + digit);
                NamedSupportedKeys.Add("NumPad" + digit);
                TextKeys.Add("D" + digit);
                TextKeys.Add("NumPad" + digit);
            }

            for (int function = 1; function <= 24; function++)
                NamedSupportedKeys.Add("F" + function);

            foreach (var key in new[]
            {
                "Space", "Oem1", "OemPlus", "OemComma", "OemMinus", "OemPeriod",
                "Oem2", "Oem3", "Oem4", "Oem5", "Oem6", "Oem7", "Oem8", "Oem102",
                "Multiply", "Add", "Subtract", "Decimal", "Divide"
            })
            {
                TextKeys.Add(key);
            }
        }

        public static ShortcutValidationResult Validate(
            IEnumerable<string> rawKeys,
            string featureId = null,
            IEnumerable<ShortcutAssignmentState> existingAssignments = null,
            bool riskConfirmed = false)
        {
            var normalized = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);

            foreach (var rawKey in rawKeys ?? Enumerable.Empty<string>())
            {
                var trimmed = (rawKey ?? string.Empty).Trim();
                if (trimmed.Length == 0 ||
                    string.Equals(trimmed, "None", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var key = NormalizeKey(trimmed);
                if (key == null || !NamedSupportedKeys.Contains(key))
                {
                    return Invalid(
                        $"'{trimmed}' cannot be observed reliably by the shortcut runtime.");
                }

                if (!seen.Add(key))
                {
                    return Invalid(
                        $"'{GetDisplayName(key)}' appears more than once after left/right key normalization.");
                }

                normalized.Add(key);
            }

            if (normalized.Count == 0)
                return Invalid("Press one or more keys, or use Clear to remove the assignment.");

            normalized = normalized
                .OrderBy(GetSortOrder)
                .ThenBy(key => key, StringComparer.Ordinal)
                .ToList();

            var chord = new ShortcutChord(normalized);

            if (normalized.Contains("Win"))
            {
                return Invalid(
                    "Windows-key chords are reserved or intercepted by Windows and cannot be assigned.",
                    chord);
            }

            if (IsCriticalSystemChord(chord))
            {
                return Invalid(
                    "This chord is reserved for critical Windows or PowerPoint behavior.",
                    chord);
            }

            var conflict = (existingAssignments ?? Enumerable.Empty<ShortcutAssignmentState>())
                .FirstOrDefault(assignment =>
                    assignment.Chord != null &&
                    !string.Equals(assignment.FeatureId, featureId, StringComparison.Ordinal) &&
                    string.Equals(assignment.Chord.Signature, chord.Signature, StringComparison.Ordinal));
            if (conflict != null)
            {
                return new ShortcutValidationResult
                {
                    IsValid = false,
                    Error = "This chord is already assigned to another openEfficiency command.",
                    ConflictingFeatureId = conflict.FeatureId,
                    Chord = chord
                };
            }

            bool modifierOnly = normalized.All(IsModifier);
            bool singleKey = normalized.Count == 1;
            bool unmodifiedText = !normalized.Any(IsModifier) &&
                                  normalized.Any(key => TextKeys.Contains(key));
            bool risky = singleKey || modifierOnly || unmodifiedText;

            return new ShortcutValidationResult
            {
                IsValid = !risky || riskConfirmed,
                RequiresRiskConfirmation = risky && !riskConfirmed,
                Error = risky && !riskConfirmed
                    ? "Confirm the activation risk before saving this chord."
                    : null,
                Warning = risky
                    ? "This chord is likely to activate during ordinary PowerPoint editing."
                    : null,
                Chord = chord
            };
        }

        public static IReadOnlyList<ShortcutAssignmentState> ClassifyLegacyAssignments(
            IEnumerable<KeyValuePair<string, string>> rawAssignments)
        {
            var states = new List<ShortcutAssignmentState>();
            foreach (var assignment in rawAssignments ?? Enumerable.Empty<KeyValuePair<string, string>>())
            {
                var validation = Validate(
                    SplitStorageValue(assignment.Value),
                    riskConfirmed: true);
                states.Add(new ShortcutAssignmentState
                {
                    FeatureId = assignment.Key,
                    RawValue = assignment.Value,
                    Chord = validation.Chord,
                    IsActive = validation.IsValid,
                    AttentionReason = validation.IsValid ? null : validation.Error
                });
            }

            foreach (var group in states
                .Where(state => state.Chord != null)
                .GroupBy(state => state.Chord.Signature, StringComparer.Ordinal)
                .Where(group => group.Count() > 1))
            {
                var featureIds = group.Select(state => state.FeatureId).ToArray();
                foreach (var state in group)
                {
                    state.IsActive = false;
                    state.AttentionReason =
                        "This chord conflicts with: " +
                        string.Join(", ", featureIds.Where(id =>
                            !string.Equals(id, state.FeatureId, StringComparison.Ordinal))) +
                        ".";
                }
            }

            return states;
        }

        public static IEnumerable<string> SplitStorageValue(string value) =>
            (value ?? string.Empty).Split(',');

        public static string FormatLegacyValue(string value)
        {
            var keys = SplitStorageValue(value)
                .Select(key => key.Trim())
                .Where(key => key.Length > 0 &&
                              !string.Equals(key, "None", StringComparison.OrdinalIgnoreCase))
                .Select(key => GetDisplayName(NormalizeKey(key) ?? key));
            var text = string.Join(" + ", keys);
            return text.Length == 0 ? "No keys" : text;
        }

        public static string NormalizeKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return null;
            var trimmed = key.Trim();
            if (Aliases.TryGetValue(trimmed, out var alias)) return alias;
            return NamedSupportedKeys.FirstOrDefault(candidate =>
                string.Equals(candidate, trimmed, StringComparison.OrdinalIgnoreCase));
        }

        public static string GetDisplayName(string key)
        {
            if (key == null) return string.Empty;
            if (key.Length == 2 && key[0] == 'D' && char.IsDigit(key[1]))
                return key.Substring(1);
            if (key.StartsWith("NumPad", StringComparison.Ordinal))
                return "Num " + key.Substring("NumPad".Length);

            switch (key)
            {
                case "Return": return "Enter";
                case "Escape": return "Esc";
                case "PageUp": return "Page Up";
                case "PageDown": return "Page Down";
                case "Snapshot": return "Print Screen";
                case "OemPlus": return "Plus";
                case "OemMinus": return "Minus";
                case "OemComma": return "Comma";
                case "OemPeriod": return "Period";
                default: return key;
            }
        }

        private static ShortcutValidationResult Invalid(
            string error,
            ShortcutChord chord = null) =>
            new ShortcutValidationResult
            {
                IsValid = false,
                Error = error,
                Chord = chord
            };

        private static bool IsModifier(string key) =>
            key == "Ctrl" || key == "Alt" || key == "Shift" || key == "Win";

        private static int GetSortOrder(string key)
        {
            switch (key)
            {
                case "Ctrl": return 0;
                case "Alt": return 1;
                case "Shift": return 2;
                case "Win": return 3;
                default: return 4;
            }
        }

        private static bool IsCriticalSystemChord(ShortcutChord chord)
        {
            var signature = chord.Signature;
            return signature == "Alt|F4" ||
                   signature == "Ctrl|Alt|Delete" ||
                   signature == "Ctrl|Shift|Escape";
        }
    }

    public static class ShortcutReplacementTransaction
    {
        public static bool TryReplace<T>(
            T previous,
            T replacement,
            Func<T, bool> tryRegister,
            Action<T> persist,
            Action<T> restorePrevious,
            out string error)
        {
            error = null;
            if (tryRegister == null || !tryRegister(replacement))
            {
                error = "The shortcut runtime could not activate the replacement.";
                return false;
            }

            try
            {
                persist?.Invoke(replacement);
                return true;
            }
            catch (Exception ex)
            {
                try { restorePrevious?.Invoke(previous); }
                catch { /* Preserve the original save error when rollback also fails. */ }
                error = "The replacement could not be saved: " + ex.Message;
                return false;
            }
        }
    }

    public static class ShortcutActivationPolicy
    {
        public static bool CanExecute(
            bool powerPointIsForeground,
            bool eligiblePresentationWindow,
            bool addInInitialized,
            bool commandHandlerInitialized,
            bool modalDialogCapturingChord) =>
            powerPointIsForeground &&
            eligiblePresentationWindow &&
            addInInitialized &&
            commandHandlerInitialized &&
            !modalDialogCapturingChord;
    }
}
