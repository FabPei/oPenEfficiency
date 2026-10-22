using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace oPenEfficiency.OneNote
{
    /// <summary>Builds the minimal, typed OneNote page updates used by both safe COM paths.</summary>
    public static class OneNotePageMutations
    {
        public static bool TryCreateAppend(
            XDocument page,
            string pageId,
            string sourceId,
            string text,
            out XDocument changes,
            out string error)
        {
            changes = null;
            if (!TryValidatePage(page, pageId, out error)) return false;
            if (string.IsNullOrWhiteSpace(sourceId))
            {
                error = "The connected-text source identifier is missing.";
                return false;
            }

            XNamespace ns = page.Root.Name.Namespace;
            var outline = new XElement(ns + "Outline",
                new XElement(ns + "Position",
                    new XAttribute("x", "36"),
                    new XAttribute("y", "36"),
                    new XAttribute("z", "0")),
                new XElement(ns + "OEChildren",
                    new XElement(ns + "OE",
                        new XElement(ns + "T", new XCData(ConnectedTextPayload.BuildHtml(sourceId, text))))));

            changes = new XDocument(
                new XElement(page.Root.Name,
                    new XAttribute("ID", pageId),
                    outline));
            error = null;
            return true;
        }

        public static bool TryCreateRefresh(
            XDocument page,
            string pageId,
            IReadOnlyList<OneNoteBridgeTarget> targets,
            out XDocument changes,
            out int replacedCount,
            out string error)
        {
            changes = null;
            replacedCount = 0;
            if (!TryValidatePage(page, pageId, out error)) return false;
            if (targets == null)
            {
                error = "The connected-text refresh request is missing its targets.";
                return false;
            }

            var changedOutlines = new HashSet<XElement>();
            foreach (OneNoteBridgeTarget target in targets)
            {
                if (target == null || string.IsNullOrWhiteSpace(target.SourceId)) continue;
                XElement element = null;
                if (!string.IsNullOrWhiteSpace(target.TargetId))
                {
                    element = page.Descendants().FirstOrDefault(node =>
                        node.Name.LocalName == "OE" &&
                        string.Equals(node.Attribute("objectID")?.Value, target.TargetId, StringComparison.OrdinalIgnoreCase));
                }

                if (element == null)
                {
                    string marker = ConnectedTextPayload.SourceMarker(target.SourceId);
                    element = page.Descendants().FirstOrDefault(node =>
                        node.Name.LocalName == "OE" &&
                        node.Value.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0);
                }

                XElement textNode = element?.Descendants().FirstOrDefault(node => node.Name.LocalName == "T");
                XElement outline = element?.Ancestors().FirstOrDefault(node => node.Name.LocalName == "Outline");
                if (textNode == null || outline == null) continue;

                textNode.ReplaceNodes(new XCData(ConnectedTextPayload.BuildHtml(target.SourceId, target.Text)));
                changedOutlines.Add(outline);
                replacedCount++;
            }

            if (replacedCount == 0)
            {
                error = null;
                return true;
            }

            changes = new XDocument(
                new XElement(page.Root.Name,
                    new XAttribute("ID", pageId),
                    changedOutlines.Select(outline => new XElement(outline))));
            error = null;
            return true;
        }

        private static bool TryValidatePage(XDocument page, string pageId, out string error)
        {
            if (page?.Root == null || string.IsNullOrWhiteSpace(pageId))
            {
                error = "The OneNote page update is incomplete.";
                return false;
            }

            string returnedPageId = page.Root.Attribute("ID")?.Value;
            if (string.IsNullOrWhiteSpace(returnedPageId) ||
                !string.Equals(returnedPageId, pageId, StringComparison.OrdinalIgnoreCase))
            {
                error = "OneNote returned content for an unexpected page. The update was cancelled.";
                return false;
            }

            if (ConnectedTextPayload.ReadLastModified(page) == DateTime.MinValue)
            {
                error = "OneNote did not provide a valid page version. The update was cancelled to avoid overwriting newer changes.";
                return false;
            }

            error = null;
            return true;
        }
    }
}
