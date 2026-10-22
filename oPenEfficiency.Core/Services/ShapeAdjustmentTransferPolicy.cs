using System;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free transfer decision for Copy Shape Adjustments (QA-S9 align-and-size wave):
    /// which explicit reference inside the selection the command uses, and which selected
    /// shapes are compatible with that reference and therefore receive its adjustment
    /// values. Production implementation used by ShapeAdjustmentTransfer
    /// (AlignShapeAdjustmentsFeature).
    /// </summary>
    public static class ShapeAdjustmentTransferPolicy
    {
        /// <summary>
        /// The 1-based index of the explicitly requested reference inside the selection
        /// ("use first/last selected shape"), or null when the command falls back to the
        /// anchor shape configured in Settings.
        /// </summary>
        public static int? ResolveExplicitReferenceIndex(int selectionCount, bool? useLastSelected)
        {
            if (!useLastSelected.HasValue) return null;
            return useLastSelected.Value ? selectionCount : 1;
        }

        /// <summary>
        /// A selected shape receives the reference's adjustment values only when it is the
        /// same PowerPoint shape type, the same AutoShapeType and offers exactly the same
        /// number of adjustments; anything else is skipped untouched.
        /// </summary>
        public static bool IsTransferCompatible(
            bool sameShapeType, bool sameAutoShapeType,
            int targetAdjustmentCount, int referenceAdjustmentCount)
        {
            return sameShapeType
                && sameAutoShapeType
                && targetAdjustmentCount == referenceAdjustmentCount;
        }
    }
}
