using System;
using System.Globalization;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free display rules behind the Excel Link Manager (QA-S9 wave 3): how a link's status
    /// text is resolved from the source's existence, and how the last-refresh tag is formatted for
    /// the list. Production implementation used by ExcelLinkManagerWindow.
    /// </summary>
    /// <remarks>
    /// The refresh formatting uses the current culture's short date/time ("g"), matching the
    /// window's historical behavior; an unparseable or missing tag renders as an em dash.
    /// </remarks>
    public static class ExcelLinkStatusPolicy
    {
        public const string LinkedStatus = "Linked";
        public const string MissingStatus = "Missing";
        public const string NoRefreshValue = "—";

        /// <summary>
        /// A link is only "Linked" when the shape carries a source path at all AND the file exists.
        /// </summary>
        public static string ResolveStatus(bool hasSourcePath, bool sourceExists)
        {
            return hasSourcePath && sourceExists ? LinkedStatus : MissingStatus;
        }

        /// <summary>Formats a last-refresh tag; blank or unparseable values render as an em dash.</summary>
        public static string FormatLastRefresh(string rawTag)
        {
            return DateTime.TryParse(rawTag, CultureInfo.CurrentCulture, DateTimeStyles.None, out var value)
                ? value.ToString("g", CultureInfo.CurrentCulture)
                : NoRefreshValue;
        }
    }
}
