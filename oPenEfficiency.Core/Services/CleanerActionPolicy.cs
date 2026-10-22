using System;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free confirmation policy for the Presentation Cleaner (QA-S9 dialogs/workflows
    /// wave): which cleanup actions warn with a Yes/No confirmation before running, and the
    /// exact warning wording per action and scope. Production implementation used by
    /// CleanerFeature.
    /// </summary>
    public static class CleanerActionPolicy
    {
        public enum CleanerAction
        {
            RemoveHiddenSlides,
            RemoveAnimationsAll,
            RemoveAnimationsSelected,
            RemoveTransitionsAll,
            RemoveTransitionsSelected,
            RemoveSpeakerNotesAll,
            RemoveSpeakerNotesSelected,
            RemoveCommentsAll,
            RemoveCommentsSelected,
            RemoveUnusedMasters,
            DeleteStickyNotes,
            AnonymizeSelected,
            RemoveUneditedFootersAll,
            RemoveUneditedFootersSelected,
            DetectSuperscriptWithoutFooter,
            RemoveDoubleSpacesAll,
            RemoveDoubleSpacesSelected,
            ConvertSlidesToImagesAll,
            ConvertSlidesToImagesSelected,
            RemoveDeveloperNotesAll,
            RemoveDeveloperNotesSelected,
            RemoveAlternativeTextAll,
            RemoveAlternativeTextSelected,
            UnlockShapesAll,
            UnlockShapesSelected,
            RemoveInkAll,
            RemoveInkSelected,
            BreakDataLinksAll,
            BreakDataLinksSelected,
            RemoveInvisibleOrOffSlideAll,
            RemoveInvisibleOrOffSlideSelected,
            RemoveDocumentInformation,
            RemoveUnusedLayouts,
            RemoveNonOpeCustomXml,
            MakeInvisibleShapesVisibleAll,
            MakeInvisibleShapesVisibleSelected,
            SanitizeNumbersAll,
            SanitizeNumbersSelected
        }

        /// <summary>
        /// Whether the action asks for confirmation before running. The read-only detection,
        /// the Anonymize flow (which confirms in its own feature), and the slide-to-image
        /// conversions are exempt.
        /// </summary>
        public static bool RequiresConfirmation(CleanerAction action)
        {
            return action != CleanerAction.DetectSuperscriptWithoutFooter &&
                   action != CleanerAction.AnonymizeSelected &&
                   action != CleanerAction.ConvertSlidesToImagesAll &&
                   action != CleanerAction.ConvertSlidesToImagesSelected;
        }

        /// <summary>
        /// The exact Yes/No confirmation wording: the operation description per action family
        /// and the slide scope, ending with the shared continue question.
        /// </summary>
        public static string GetConfirmationText(CleanerAction action)
        {
            string scope = action.ToString().EndsWith("Selected", StringComparison.Ordinal) ? "the selected slide scope" : "the entire presentation";
            string operation;
            switch (action)
            {
                case CleanerAction.RemoveHiddenSlides: operation = "delete all hidden slides"; break;
                case CleanerAction.RemoveAnimationsAll:
                case CleanerAction.RemoveAnimationsSelected: operation = "remove animations"; break;
                case CleanerAction.RemoveTransitionsAll:
                case CleanerAction.RemoveTransitionsSelected: operation = "remove transitions"; break;
                case CleanerAction.RemoveSpeakerNotesAll:
                case CleanerAction.RemoveSpeakerNotesSelected: operation = "delete speaker notes"; break;
                case CleanerAction.RemoveCommentsAll:
                case CleanerAction.RemoveCommentsSelected: operation = "delete comments"; break;
                case CleanerAction.RemoveDeveloperNotesAll:
                case CleanerAction.RemoveDeveloperNotesSelected: operation = "delete developer notes"; break;
                case CleanerAction.RemoveUnusedMasters: operation = "delete slide masters that are not used by any slide"; break;
                case CleanerAction.DeleteStickyNotes: operation = "delete openEfficiency sticky-note shapes"; break;
                case CleanerAction.RemoveUneditedFootersAll:
                case CleanerAction.RemoveUneditedFootersSelected: operation = "delete empty or unedited footer placeholders"; break;
                case CleanerAction.RemoveDoubleSpacesAll:
                case CleanerAction.RemoveDoubleSpacesSelected: operation = "replace repeated spaces"; break;
                case CleanerAction.RemoveAlternativeTextAll:
                case CleanerAction.RemoveAlternativeTextSelected: operation = "remove title and alternative text from unprotected shapes"; break;
                case CleanerAction.UnlockShapesAll:
                case CleanerAction.UnlockShapesSelected: operation = "unlock unprotected shapes"; break;
                case CleanerAction.RemoveInkAll:
                case CleanerAction.RemoveInkSelected: operation = "delete ink objects"; break;
                case CleanerAction.BreakDataLinksAll:
                case CleanerAction.BreakDataLinksSelected: operation = "break external shape data links and keep their current appearance"; break;
                case CleanerAction.RemoveInvisibleOrOffSlideAll:
                case CleanerAction.RemoveInvisibleOrOffSlideSelected: operation = "delete unprotected invisible and fully off-slide objects"; break;
                case CleanerAction.RemoveDocumentInformation: operation = "remove document properties and personal information"; break;
                case CleanerAction.RemoveUnusedLayouts: operation = "delete unused custom slide layouts"; break;
                case CleanerAction.RemoveNonOpeCustomXml: operation = "delete custom XML that is not owned by openEfficiency"; break;
                case CleanerAction.MakeInvisibleShapesVisibleAll:
                case CleanerAction.MakeInvisibleShapesVisibleSelected: operation = "make unprotected invisible shapes visible"; break;
                case CleanerAction.SanitizeNumbersAll:
                case CleanerAction.SanitizeNumbersSelected: operation = "replace visible digits with random digits while retaining text formatting"; break;
                default: operation = "run this cleanup action"; break;
            }
            return $"Presentation Cleaner will {operation} in {scope}.\n\nContinue?";
        }
    }
}
