using System;
using System.Collections.Generic;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free page-range selection for Local Slide Rebuild (QA-S9 external-data wave):
    /// the accepted page-range grammar for restricting a rebuild to specific source pages.
    /// Production implementation used by LocalSlideRebuildFeature; the document rendering
    /// and rebuild execution stay host-tier.
    /// </summary>
    public static class SlideRebuildPageRangePolicy
    {
        public const uint MinimumPage = 1;
        public const uint MaximumPage = 10000;
        public const uint MaximumRangeSpan = 999;

        /// <summary>
        /// Parses a comma-separated page selection of single pages ("3") and inclusive
        /// ranges ("4-7"). Pages are 1-based and capped at 10000; a range may span at most
        /// 1000 pages (end - start &lt;= 999); duplicates collapse into the resulting set.
        /// Any malformed token rejects the whole input.
        /// </summary>
        public static bool TryParsePageRange(string value, out HashSet<uint> pageNumbers)
        {
            pageNumbers = new HashSet<uint>();
            if (string.IsNullOrWhiteSpace(value)) return false;
            foreach (string tokenValue in value.Split(','))
            {
                string token = tokenValue.Trim();
                if (token.Length == 0) return false;
                string[] bounds = token.Split('-');
                if (bounds.Length == 1)
                {
                    if (!uint.TryParse(bounds[0], out uint page) || page < MinimumPage || page > MaximumPage) return false;
                    pageNumbers.Add(page);
                }
                else if (bounds.Length == 2)
                {
                    if (!uint.TryParse(bounds[0], out uint start) || !uint.TryParse(bounds[1], out uint end)
                        || start < MinimumPage || end > MaximumPage || end < start || end - start > MaximumRangeSpan)
                        return false;
                    for (uint page = start; page <= end; page++) pageNumbers.Add(page);
                }
                else return false;
            }
            return pageNumbers.Count > 0;
        }
    }
}
