using System.Collections.Generic;
using System.Linq;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Host-neutral stacking-order plans for the z-order commands (BtnBringForward, BtnSendBackward,
    /// BtnBringToFront, BtnSendToBack and their ranked/relative right-click variants). The COM layer
    /// executes these plans shape-by-shape; this class only decides them. All orders are bottom-to-top
    /// id sequences (the last element ends up topmost).
    /// </summary>
    public static class ZOrderRankPlanner
    {
        /// <summary>
        /// Plans the stacked order after a ranked z-order command. For the absolute commands
        /// (bring-to-front / send-to-back) the selection is appended before or after the unselected
        /// shapes; for the one-step commands the plan fills the selection's slots in the (recaptured)
        /// working order with the ranked selection. <paramref name="selectedInSelectionOrder"/> is the
        /// shape-selection order (first selected first).
        /// </summary>
        public static List<int> PlanRanked(
            IList<int> workingOrderIds,
            IList<int> selectedInSelectionOrder,
            bool firstSelectedOnTop,
            bool bringToFront,
            bool sendToBack)
        {
            var ranked = firstSelectedOnTop
                ? selectedInSelectionOrder.Reverse().ToList()
                : new List<int>(selectedInSelectionOrder);
            var selectedSet = new HashSet<int>(selectedInSelectionOrder);
            var unselected = workingOrderIds.Where(id => !selectedSet.Contains(id)).ToList();
            if (bringToFront) return unselected.Concat(ranked).ToList();
            if (sendToBack) return ranked.Concat(unselected).ToList();

            var desired = new List<int>(workingOrderIds);
            int rankedIndex = 0;
            for (int i = 0; i < desired.Count; i++)
            {
                if (selectedSet.Contains(desired[i]))
                {
                    desired[i] = ranked[rankedIndex++];
                }
            }
            return desired;
        }

        /// <summary>
        /// Plans the stacked order that places the movers directly above or below the reference shape
        /// (the last selected shape). The movers keep their internal order and unselected shapes keep
        /// their relative order. Returns null when the reference is missing or there is nothing to move.
        /// </summary>
        public static List<int> PlanRelativeToReference(
            IList<int> workingOrderIds,
            IEnumerable<int> moverIds,
            int referenceId,
            bool placeAbove)
        {
            var moverSet = new HashSet<int>(moverIds);
            var movers = workingOrderIds.Where(moverSet.Contains).ToList();
            var desired = workingOrderIds.Where(id => !moverSet.Contains(id)).ToList();
            int referenceIndex = desired.IndexOf(referenceId);
            if (referenceIndex < 0 || movers.Count == 0) return null;
            desired.InsertRange(referenceIndex + (placeAbove ? 1 : 0), movers);
            return desired;
        }
    }
}
