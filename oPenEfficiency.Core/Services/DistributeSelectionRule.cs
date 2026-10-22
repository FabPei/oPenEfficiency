namespace oPenEfficiency.Services
{
    /// <summary>
    /// The complete decision the two distribute commands ask for before they touch the PowerPoint ShapeRange:
    /// whether the selection may be distributed at all and which route distributes it. Host-neutral on purpose, so
    /// that the commands only apply the decision while the acceptance tests execute it for every shape count.
    /// Review round 6g measured that text assertions over the command's guard could be re-expressed (a helper in
    /// another file, a shadow rule class, a second assignment to the resolved route) while every gate stayed green;
    /// the plan exists so that the decision itself is executed rather than described.
    /// </summary>
    public struct DistributePlan
    {
        /// <summary>True when the selection may be distributed by any route.</summary>
        public bool Distribute;

        /// <summary>True when the distribution is measured against the slide; false for the interior gaps.</summary>
        public bool RelativeToSlide;
    }

    /// <summary>
    /// The selection contract of the two distribute commands (BtnDistributeHorizontal / BtnDistributeVertical),
    /// kept host-neutral so that the VSTO command - which owns the PowerPoint ShapeRange - and the acceptance
    /// scenarios - which cannot open PowerPoint - evaluate the same rule instead of two implementations.
    /// Owner decision a1 for REG-DISTRIBUTE-SCOPE-001: the sidebar button is enabled from MinSelection = 2 and the
    /// primary click always performs a distribution - with exactly two selected shapes it spreads them across the
    /// slide (there is no interior gap to distribute), from three shapes it distributes the interior gaps. The
    /// context-menu variants always choose their own route. Both rules are asserted by
    /// <c>DistributeSelectionGuardTests</c> (scenarios QA9-BtnDistributeHorizontal-002 / -Vertical-002) and bound to
    /// the two commands by the catalog guard tests.
    /// </summary>
    public static class DistributeSelectionRule
    {
        /// <summary>Shapes at which the primary route distributes interior gaps instead of the slide.</summary>
        public const int MinimumShapesForPrimaryRoute = 3;

        /// <summary>Shapes either route needs; two shapes can be spread over the slide.</summary>
        public const int MinimumShapesForAcrossSlideRoute = 2;

        /// <summary>
        /// True when a selection of <paramref name="selectedShapeCount"/> shapes may be distributed by the route
        /// that <paramref name="relativeToSlide"/> selects. The result is exactly the condition the two Execute
        /// methods apply before they touch the ShapeRange.
        /// </summary>
        public static bool MeetsMinimum(int selectedShapeCount, bool relativeToSlide)
        {
            return relativeToSlide
                ? selectedShapeCount >= MinimumShapesForAcrossSlideRoute
                : selectedShapeCount >= MinimumShapesForPrimaryRoute;
        }

        /// <summary>
        /// Which route a selection is distributed by. An explicit across-slide request (the context-menu variant)
        /// always keeps its route; the primary click falls back to the across-slide route at two shapes, because
        /// distributing an interior gap needs a shape between the outer two.
        /// </summary>
        public static bool UsesAcrossSlideRoute(int selectedShapeCount, bool relativeToSlide)
        {
            if (relativeToSlide)
            {
                return true;
            }

            return selectedShapeCount <= MinimumShapesForAcrossSlideRoute;
        }

        /// <summary>
        /// The whole decision in one call: which route distributes the selection and whether it may be distributed
        /// at all. The commands apply exactly this plan; the acceptance tests execute it for every shape count.
        /// </summary>
        public static DistributePlan ResolvePlan(int selectedShapeCount, bool relativeToSlide)
        {
            var acrossSlide = UsesAcrossSlideRoute(selectedShapeCount, relativeToSlide);
            var plan = new DistributePlan();
            plan.RelativeToSlide = acrossSlide;
            plan.Distribute = MeetsMinimum(selectedShapeCount, acrossSlide);
            return plan;
        }
    }
}
