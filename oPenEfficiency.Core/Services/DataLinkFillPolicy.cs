namespace oPenEfficiency.Services
{
    /// <summary>
    /// Fill-transfer policy shared by the Excel-side link creation and the PowerPoint-side refresh,
    /// so a linked table cell looks the same no matter which path built or last updated it
    /// (FEAT-005 DL-FORMAT-NOFILL-001).
    /// </summary>
    public static class DataLinkFillPolicy
    {
        /// <summary>
        /// BGR white. Excel displays a "No Fill" cell as the sheet background, which is white.
        /// A linked table cell therefore has to carry this explicit fill: leaving the cell's own
        /// fill invisible exposes the PowerPoint table style instead, which is the gray the owner
        /// reported as a formatting-parity defect.
        /// </summary>
        public const int NoFillBgr = 0xFFFFFF;

        /// <summary>
        /// Resolves the BGR fill a destination cell must receive for a source cell: the source
        /// color for an explicit Excel fill, otherwise the documented No-Fill white.
        /// </summary>
        public static int ResolveFillBgr(bool sourceHasFill, int sourceFillBgr)
        {
            return sourceHasFill ? sourceFillBgr : NoFillBgr;
        }
    }
}
