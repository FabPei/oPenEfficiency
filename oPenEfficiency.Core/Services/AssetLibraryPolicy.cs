using System;
using System.IO;
using System.Linq;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free guards behind the Asset Manager (QA-S9 wave 3): the library containment rule that
    /// gates every write/delete outside the asset root, and the folder-name sanitization.
    /// Production implementation used by LocalAssetService.
    /// </summary>
    public static class AssetLibraryPolicy
    {
        /// <summary>
        /// Whether a path lies inside (or is) the asset root. Both sides are resolved to full
        /// paths, trailing separators are trimmed and the comparison is case-insensitive; a
        /// sibling folder sharing the root's prefix (e.g. C:\Root2 against C:\Root) is outside.
        /// </summary>
        public static bool IsWithinRoot(string root, string path)
        {
            if (string.IsNullOrWhiteSpace(root) || string.IsNullOrWhiteSpace(path)) return false;
            try
            {
                string normalizedRoot = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                string candidate = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                return candidate.Equals(normalizedRoot, StringComparison.OrdinalIgnoreCase) ||
                       candidate.StartsWith(normalizedRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Sanitizes a folder name: invalid filename characters are removed, spaces become
        /// underscores after trimming, blank results fall back to "unnamed" and names are capped
        /// at 100 characters.
        /// </summary>
        public static string MakeValidFolderName(string folderName)
        {
            if (string.IsNullOrEmpty(folderName))
                return "unnamed";

            char[] invalidChars = Path.GetInvalidFileNameChars();
            string validName = string.Concat(folderName.Split(invalidChars));

            validName = validName.Trim().Replace(" ", "_");

            if (string.IsNullOrWhiteSpace(validName)) validName = "unnamed";

            return validName.Length > 100 ? validName.Substring(0, 100) : validName;
        }

        /// <summary>
        /// Resolves one new child directory beneath a selected library folder. Path-like input
        /// and Windows dot aliases are rejected; ordinary names retain the existing sanitizer.
        /// The returned canonical path is the exact value callers must pass to the filesystem.
        /// </summary>
        public static string ResolveNewFolderPath(string root, string parentPath, string folderName)
        {
            if (!IsWithinRoot(root, parentPath))
                throw new UnauthorizedAccessException("The target folder is outside the Asset Library.");

            if (!string.IsNullOrEmpty(folderName) &&
                (Path.IsPathRooted(folderName) || folderName.IndexOf(Path.DirectorySeparatorChar) >= 0 ||
                 folderName.IndexOf(Path.AltDirectorySeparatorChar) >= 0))
                throw new ArgumentException("Enter a folder name, not a path.", nameof(folderName));

            string safeName = MakeValidFolderName(folderName);
            if (safeName.EndsWith(".", StringComparison.Ordinal))
                throw new ArgumentException("Folder names cannot end with a dot.", nameof(folderName));

            string parent = Path.GetFullPath(parentPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string destination = Path.GetFullPath(Path.Combine(parent, safeName));
            if (!IsWithinRoot(root, destination) || !IsWithinRoot(parent, destination) ||
                destination.Equals(parent, StringComparison.OrdinalIgnoreCase))
                throw new UnauthorizedAccessException("The new folder must be inside the selected Asset Library folder.");

            return destination;
        }
    }
}
