using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 dialogs/workflows-wave coverage for the Presentation Cleaner confirmation policy
    /// (QA9-BtnCleaner-002): which of the 38 cleanup actions warn before running and the exact
    /// warning wording per action and scope, used by CleanerFeature.
    /// </summary>
    [TestClass]
    public sealed class CleanerActionPolicyTests
    {
        [TestMethod]
        public void DestructiveActionsRequireConfirmation()
        {
            Assert.IsTrue(CleanerActionPolicy.RequiresConfirmation(CleanerActionPolicy.CleanerAction.RemoveHiddenSlides));
            Assert.IsTrue(CleanerActionPolicy.RequiresConfirmation(CleanerActionPolicy.CleanerAction.RemoveSpeakerNotesAll));
            Assert.IsTrue(CleanerActionPolicy.RequiresConfirmation(CleanerActionPolicy.CleanerAction.RemoveCommentsSelected));
            Assert.IsTrue(CleanerActionPolicy.RequiresConfirmation(CleanerActionPolicy.CleanerAction.RemoveUnusedMasters));
            Assert.IsTrue(CleanerActionPolicy.RequiresConfirmation(CleanerActionPolicy.CleanerAction.SanitizeNumbersSelected));
            Assert.IsTrue(CleanerActionPolicy.RequiresConfirmation(CleanerActionPolicy.CleanerAction.BreakDataLinksAll));
            Assert.IsTrue(CleanerActionPolicy.RequiresConfirmation(CleanerActionPolicy.CleanerAction.RemoveNonOpeCustomXml));
        }

        [TestMethod]
        public void DetectionAnonymizeAndImageConversionAreExempt()
        {
            Assert.IsFalse(CleanerActionPolicy.RequiresConfirmation(CleanerActionPolicy.CleanerAction.DetectSuperscriptWithoutFooter));
            Assert.IsFalse(CleanerActionPolicy.RequiresConfirmation(CleanerActionPolicy.CleanerAction.AnonymizeSelected));
            Assert.IsFalse(CleanerActionPolicy.RequiresConfirmation(CleanerActionPolicy.CleanerAction.ConvertSlidesToImagesAll));
            Assert.IsFalse(CleanerActionPolicy.RequiresConfirmation(CleanerActionPolicy.CleanerAction.ConvertSlidesToImagesSelected));
        }

        [TestMethod]
        public void ConfirmationTextNamesTheOperationAndTheScope()
        {
            Assert.AreEqual(
                "Presentation Cleaner will delete all hidden slides in the entire presentation.\n\nContinue?",
                CleanerActionPolicy.GetConfirmationText(CleanerActionPolicy.CleanerAction.RemoveHiddenSlides));
            Assert.AreEqual(
                "Presentation Cleaner will remove animations in the selected slide scope.\n\nContinue?",
                CleanerActionPolicy.GetConfirmationText(CleanerActionPolicy.CleanerAction.RemoveAnimationsSelected));
            Assert.AreEqual(
                "Presentation Cleaner will delete empty or unedited footer placeholders in the entire presentation.\n\nContinue?",
                CleanerActionPolicy.GetConfirmationText(CleanerActionPolicy.CleanerAction.RemoveUneditedFootersAll));
            Assert.AreEqual(
                "Presentation Cleaner will replace visible digits with random digits while retaining text formatting in the selected slide scope.\n\nContinue?",
                CleanerActionPolicy.GetConfirmationText(CleanerActionPolicy.CleanerAction.SanitizeNumbersSelected));
        }

        [TestMethod]
        public void EveryActionFamilyCarriesItsDistinctWording()
        {
            var expected = new[]
            {
                (CleanerActionPolicy.CleanerAction.RemoveTransitionsAll, "remove transitions"),
                (CleanerActionPolicy.CleanerAction.RemoveSpeakerNotesSelected, "delete speaker notes"),
                (CleanerActionPolicy.CleanerAction.RemoveCommentsAll, "delete comments"),
                (CleanerActionPolicy.CleanerAction.RemoveDeveloperNotesSelected, "delete developer notes"),
                (CleanerActionPolicy.CleanerAction.RemoveUnusedMasters, "delete slide masters that are not used by any slide"),
                (CleanerActionPolicy.CleanerAction.DeleteStickyNotes, "delete openEfficiency sticky-note shapes"),
                (CleanerActionPolicy.CleanerAction.RemoveDoubleSpacesAll, "replace repeated spaces"),
                (CleanerActionPolicy.CleanerAction.RemoveAlternativeTextSelected, "remove title and alternative text from unprotected shapes"),
                (CleanerActionPolicy.CleanerAction.UnlockShapesAll, "unlock unprotected shapes"),
                (CleanerActionPolicy.CleanerAction.RemoveInkSelected, "delete ink objects"),
                (CleanerActionPolicy.CleanerAction.BreakDataLinksSelected, "break external shape data links and keep their current appearance"),
                (CleanerActionPolicy.CleanerAction.RemoveInvisibleOrOffSlideAll, "delete unprotected invisible and fully off-slide objects"),
                (CleanerActionPolicy.CleanerAction.RemoveDocumentInformation, "remove document properties and personal information"),
                (CleanerActionPolicy.CleanerAction.RemoveUnusedLayouts, "delete unused custom slide layouts"),
                (CleanerActionPolicy.CleanerAction.RemoveNonOpeCustomXml, "delete custom XML that is not owned by openEfficiency"),
                (CleanerActionPolicy.CleanerAction.MakeInvisibleShapesVisibleSelected, "make unprotected invisible shapes visible")
            };
            foreach (var pair in expected)
            {
                string text = CleanerActionPolicy.GetConfirmationText(pair.Item1);
                StringAssert.Contains(text, "Presentation Cleaner will " + pair.Item2 + " in ");
                StringAssert.EndsWith(text, "\n\nContinue?");
            }
        }

        [TestMethod]
        public void ExemptActionsStillResolveTheDefaultWording()
        {
            // The exempt actions never show this dialog, but the wording must stay stable.
            Assert.AreEqual(
                "Presentation Cleaner will run this cleanup action in the selected slide scope.\n\nContinue?",
                CleanerActionPolicy.GetConfirmationText(CleanerActionPolicy.CleanerAction.ConvertSlidesToImagesSelected));
        }

        [TestMethod]
        public void ScopeSelectionFollowsTheSelectedSuffix()
        {
            var all = new[] { CleanerActionPolicy.CleanerAction.RemoveAnimationsAll };
            var selected = new[] { CleanerActionPolicy.CleanerAction.RemoveAnimationsSelected };
            StringAssert.Contains(CleanerActionPolicy.GetConfirmationText(all[0]), "the entire presentation");
            StringAssert.Contains(CleanerActionPolicy.GetConfirmationText(selected[0]), "the selected slide scope");
        }
    }
}
