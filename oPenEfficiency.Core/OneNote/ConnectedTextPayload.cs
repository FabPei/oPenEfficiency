using System;
using System.Runtime.Serialization;
using System.Xml.Linq;

namespace oPenEfficiency.OneNote
{
    /// <summary>
    /// Host-neutral payload helpers for connected text labels (FEAT-007 ON-S0).
    /// The HTML marker format and the persisted record shape are compatibility
    /// contracts: docs/testing/feat-007-connected-onenote-compatibility.md.
    /// </summary>
    public static class ConnectedTextPayload
    {
        /// <summary>
        /// Builds the OneNote HTML fragment stored inside a connected label.
        /// The data attribute is the refresh fallback marker when the OE objectID
        /// lookup fails, so its exact spelling is part of the legacy contract.
        /// </summary>
        public static string BuildHtml(string sourceId, string text)
        {
            string encoded = System.Net.WebUtility.HtmlEncode(text ?? string.Empty)
                .Replace("\r\n", "<br />")
                .Replace("\n", "<br />")
                .Replace("\r", "<br />");
            return $"<span data-ope-source=\"{sourceId}\">{encoded}</span>";
        }

        /// <summary>The exact marker substring identifying a source inside page XML text.</summary>
        public static string SourceMarker(string sourceId)
        {
            return "data-ope-source=\"" + sourceId + "\"";
        }

        /// <summary>
        /// Reads the page-level lastModifiedTime attribute used as the
        /// expected-last-modified concurrency token for UpdatePageContent.
        /// Returns DateTime.MinValue when absent or unparsable.
        /// </summary>
        public static DateTime ReadLastModified(XDocument page)
        {
            if (page?.Root == null) return DateTime.MinValue;
            string raw = page.Root.Attribute("lastModifiedTime")?.Value;
            if (string.IsNullOrEmpty(raw)) return DateTime.MinValue;
            return DateTime.TryParse(raw, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.RoundtripKind, out DateTime value)
                ? value
                : DateTime.MinValue;
        }
    }

    /// <summary>
    /// Persisted connected-target record (connected_text_labels.json).
    /// Property names are the legacy JSON contract and must not change.
    /// </summary>
    [DataContract]
    public sealed class ConnectedTextTargetRecord
    {
        [DataMember]
        public string SourceId { get; set; }

        [DataMember]
        public string Host { get; set; }

        [DataMember]
        public string DocumentPath { get; set; }

        [DataMember]
        public string TargetId { get; set; }
    }
}
