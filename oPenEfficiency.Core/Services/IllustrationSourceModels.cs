using System;

namespace oPenEfficiency.Models
{
    /// <summary>
    /// How an external illustration source relates to the artwork this add-in ships.
    /// The distinction matters: "no attribution needed" is not the same as "safe to bundle",
    /// and only the second one decides whether files may travel inside the installer.
    /// </summary>
    public enum IllustrationSourceStatus
    {
        /// <summary>The artwork ships with the add-in because the license permits redistribution.</summary>
        Bundled,

        /// <summary>Redistribution is permitted by the source's own terms, but nothing is shipped with the add-in.</summary>
        BundlingPermitted,

        /// <summary>The source's terms restrict redistribution - importing is the user's own licensing decision.</summary>
        RedistributionRestricted
    }

    /// <summary>
    /// One evaluated illustration source as shown in the Illustration Library window: where to get it,
    /// which license it carries, whether it may travel inside the add-in, and what the artwork can do
    /// technically (named parts, inline colours).
    /// </summary>
    public class IllustrationSource
    {
        /// <summary>Stable id, e.g. "open-peeps".</summary>
        public string Id { get; set; }

        /// <summary>Display name, e.g. "Open Peeps".</summary>
        public string Name { get; set; }

        /// <summary>Page the user downloads from.</summary>
        public string HomePageUrl { get; set; }

        /// <summary>License as stated by the source itself, e.g. "CC0 1.0 Universal".</summary>
        public string LicenseName { get; set; }

        /// <summary>License page or terms page used to verify <see cref="LicenseName"/>.</summary>
        public string LicenseUrl { get; set; }

        public IllustrationSourceStatus Status { get; set; }

        /// <summary>True when the license obliges visible credit on every use.</summary>
        public bool AttributionRequired { get; set; }

        /// <summary>How to obtain the files; empty for bundled sources.</summary>
        public string AcquisitionNote { get; set; }

        /// <summary>What the artwork can do after import (named parts, colour mechanism, caveats).</summary>
        public string TechnicalNote { get; set; }

        /// <summary>Number of figures bundled with the add-in, 0 when nothing is shipped.</summary>
        public int BundledFigureCount { get; set; }

        /// <summary>One-line status shown as the row badge.</summary>
        public string StatusLabel
        {
            get
            {
                switch (Status)
                {
                    case IllustrationSourceStatus.Bundled:
                        return "Included in the add-in";
                    case IllustrationSourceStatus.BundlingPermitted:
                        return "Bundling allowed - not shipped";
                    default:
                        return "Check the license - redistribution restricted";
                }
            }
        }

        /// <summary>Short license summary for the details pane, e.g. "CC0 1.0 Universal - no attribution".</summary>
        public string LicenseSummary
        {
            get
            {
                string license = string.IsNullOrWhiteSpace(LicenseName) ? "Unknown license" : LicenseName;
                return AttributionRequired ? license + " - attribution required" : license + " - no attribution";
            }
        }

        /// <summary>True when this source's files already sit in the add-in's own library.</summary>
        public bool IsBundled
        {
            get { return Status == IllustrationSourceStatus.Bundled; }
        }
    }
}
