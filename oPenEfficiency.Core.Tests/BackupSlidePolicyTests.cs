using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 dialogs/workflows-wave coverage for the designated Backup slide
    /// (QA9-BtnMoveToBackup-002): the oPE/EE4P tag recognition and the move/copy confirmation
    /// wording, used by MoveToBackupFeature.
    /// </summary>
    [TestClass]
    public sealed class BackupSlidePolicyTests
    {
        [TestMethod]
        public void EitherMarkerTagWithTrueDesignatesTheBackupSlide()
        {
            Assert.IsTrue(BackupSlidePolicy.IsDesignatedBackupSlide("True", null));
            Assert.IsTrue(BackupSlidePolicy.IsDesignatedBackupSlide(null, "True"));
            Assert.IsTrue(BackupSlidePolicy.IsDesignatedBackupSlide("True", "True"));
        }

        [TestMethod]
        public void TagComparisonIsCaseInsensitive()
        {
            Assert.IsTrue(BackupSlidePolicy.IsDesignatedBackupSlide("true", null));
            Assert.IsTrue(BackupSlidePolicy.IsDesignatedBackupSlide(null, "TRUE"));
        }

        [TestMethod]
        public void MissingOrOtherValuesDoNotQualify()
        {
            Assert.IsFalse(BackupSlidePolicy.IsDesignatedBackupSlide(null, null));
            Assert.IsFalse(BackupSlidePolicy.IsDesignatedBackupSlide("", ""));
            Assert.IsFalse(BackupSlidePolicy.IsDesignatedBackupSlide("False", "False"));
            Assert.IsFalse(BackupSlidePolicy.IsDesignatedBackupSlide("1", "Yes"));
        }

        [TestMethod]
        public void ConfirmationTextUsesTheRightVerbCountsAndIndex()
        {
            Assert.AreEqual(
                "Move 3 selected slide(s) to the end, behind Backup slide 7?",
                BackupSlidePolicy.BuildConfirmationText(isCopy: false, slideCount: 3, backupSlideIndex: 7));
            Assert.AreEqual(
                "Copy 12 selected slide(s) to the end, behind Backup slide 25?",
                BackupSlidePolicy.BuildConfirmationText(isCopy: true, slideCount: 12, backupSlideIndex: 25));
        }

        [TestMethod]
        public void TagLiteralsAreStable()
        {
            Assert.AreEqual("oPE_BackupSlide", BackupSlidePolicy.OpeBackupSlideTag);
            Assert.AreEqual("EE4P_BACKUP_SLIDE", BackupSlidePolicy.Ee4pBackupSlideTag);
        }
    }
}
