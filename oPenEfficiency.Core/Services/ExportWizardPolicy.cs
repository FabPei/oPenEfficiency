using System;
using System.IO;
using System.Linq;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free naming and path rules behind the Export Wizard (QA-S9 wave 3): how filename
    /// components are sanitized before an export is written and how two target paths are compared
    /// for the overwrite confirmation. Production implementation used by ExportWizard.
    /// </summary>
    /// <remarks>
    /// Unlike the Asset Library sanitization (which REMOVES invalid characters), the export wizard
    /// REPLACES whitespace and invalid characters with underscores - both contracts are pinned
    /// separately on purpose.
    /// </remarks>
    public static class ExportWizardPolicy
    {
        /// <summary>Replaces whitespace and invalid filename characters with underscores.</summary>
        public static string SanitizeFilenameComponent(string value)
        {
            char[] invalidCharacters = Path.GetInvalidFileNameChars();
            return new string((value ?? string.Empty)
                .Select(character => char.IsWhiteSpace(character) || invalidCharacters.Contains(character) ? '_' : character)
                .ToArray());
        }

        /// <summary>Case-insensitive full-path comparison; blank inputs are never equal.</summary>
        public static bool PathsEqual(string first, string second)
        {
            if (string.IsNullOrWhiteSpace(first) || string.IsNullOrWhiteSpace(second)) return false;
            return string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase);
        }
    }
}
