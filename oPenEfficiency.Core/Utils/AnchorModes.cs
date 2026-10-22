using System;

namespace oPenEfficiency.Utils
{
    /// <summary>
    /// The single host-neutral decision for the configured anchor object mode ("First selected" vs
    /// "Last Selected"). Every production and UI comparison of the configured mode routes through this
    /// predicate so the string mapping cannot diverge between call sites.
    /// </summary>
    public static class AnchorModes
    {
        public static bool IsLastSelected(string masterObjectMode)
        {
            return string.Equals(masterObjectMode, "Last Selected", StringComparison.OrdinalIgnoreCase);
        }
    }
}
