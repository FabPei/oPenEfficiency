using System;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free contract for the designated Backup slide (QA-S9 dialogs/workflows wave): the
    /// openEfficiency and EE4P tag recognition and the move/copy confirmation wording.
    /// Production implementation used by MoveToBackupFeature; the Agenda features write the
    /// tag with the same literals.
    /// </summary>
    public static class BackupSlidePolicy
    {
        public const string OpeBackupSlideTag = "oPE_BackupSlide";
        public const string Ee4pBackupSlideTag = "EE4P_BACKUP_SLIDE";

        /// <summary>
        /// Whether a slide is the designated Backup slide: either marker tag carries "True",
        /// compared case-insensitively; missing or other values do not qualify.
        /// </summary>
        public static bool IsDesignatedBackupSlide(string opeBackupSlideTagValue, string ee4pBackupSlideTagValue)
        {
            return string.Equals(opeBackupSlideTagValue, "True", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(ee4pBackupSlideTagValue, "True", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>The exact Yes/No confirmation wording for a move or copy action.</summary>
        public static string BuildConfirmationText(bool isCopy, int slideCount, int backupSlideIndex)
        {
            string verb = isCopy ? "Copy" : "Move";
            return $"{verb} {slideCount} selected slide(s) to the end, behind Backup slide {backupSlideIndex}?";
        }
    }
}
