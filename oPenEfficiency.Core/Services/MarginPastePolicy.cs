namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free paste plan for internal text margins (QA-S9 wave 5): which axes a paste action
    /// writes and which values they receive, including the reset-to-zero rule. Production
    /// implementation used by SetMarginFeature.
    /// </summary>
    public static class MarginPastePolicy
    {
        /// <summary>
        /// The resolved write plan for one paste action: the axes to write and the values they
        /// receive. Reset writes zeros for the selected axes only; a non-reset paste writes the
        /// learned values for the selected axes and never enlarges the write set.
        /// </summary>
        public sealed class MarginPlan
        {
            public bool ApplyHorizontal;
            public bool ApplyVertical;
            public float Left;
            public float Right;
            public float Top;
            public float Bottom;
        }

        /// <summary>
        /// Builds the plan for the paste combinations the context menu exposes (all margins,
        /// horizontal only, vertical only, reset). A non-reset paste requires learned session
        /// data; reset is always allowed. Returns false when the data requirement fails.
        /// </summary>
        public static bool TryBuildPlan(
            bool applyHorizontal,
            bool applyVertical,
            bool resetToZero,
            bool hasLearnedData,
            float learnedLeft,
            float learnedRight,
            float learnedTop,
            float learnedBottom,
            out MarginPlan plan)
        {
            plan = null;
            if (!resetToZero && !hasLearnedData) return false;

            plan = new MarginPlan
            {
                ApplyHorizontal = applyHorizontal,
                ApplyVertical = applyVertical,
                Left = resetToZero ? 0f : learnedLeft,
                Right = resetToZero ? 0f : learnedRight,
                Top = resetToZero ? 0f : learnedTop,
                Bottom = resetToZero ? 0f : learnedBottom
            };
            return true;
        }
    }
}
