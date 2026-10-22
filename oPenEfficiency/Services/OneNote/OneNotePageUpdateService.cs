using System;
using System.Collections.Generic;
using System.Xml.Linq;
using oPenEfficiency.OneNote;

namespace oPenEfficiency.Services.OneNote
{
    /// <summary>
    /// Isolated page-XML mutation flows for connected text labels (FEAT-007 ON-S0).
    /// Every write captures the page's lastModifiedTime first and passes it to
    /// UpdatePageContent with force=false, so a concurrently edited page rejects the
    /// update instead of being overwritten.
    /// </summary>
    public static class OneNotePageUpdateService
    {
        /// <summary>
        /// Appends a new connected-text outline to the page. XML shape is the legacy
        /// contract (Position 36/36/0 → OEChildren → OE → T CDATA payload).
        /// </summary>
        public static bool TryAppendConnectedText(
            OneNoteDesktopClient client,
            string pageId,
            string sourceId,
            string text,
            out string error)
        {
            error = null;
            if (!client.TryGetPageXml(pageId, out XDocument page, out error)) return false;
            DateTime expectedLastModified = ConnectedTextPayload.ReadLastModified(page);

            if (!OneNotePageMutations.TryCreateAppend(
                    page, pageId, sourceId, text, out XDocument changes, out error)) return false;
            return client.TryUpdatePage(changes, expectedLastModified, out error);
        }

        /// <summary>
        /// Replaces the text of every connected label on one page. The page is read
        /// once and written once (guarded); labels are found by OE objectID first and
        /// by the legacy data-ope-source marker as fallback. Returns how many labels
        /// were replaced; 0 with no error simply means nothing matched.
        /// </summary>
        public static bool TryReplacePageTexts(
            OneNoteDesktopClient client,
            string pageId,
            IReadOnlyList<ConnectedTextTargetRecord> pageTargets,
            IReadOnlyDictionary<string, string> textsBySourceId,
            out int replacedCount,
            out string error)
        {
            replacedCount = 0;
            error = null;
            if (!client.TryGetPageXml(pageId, out XDocument page, out error)) return false;
            DateTime expectedLastModified = ConnectedTextPayload.ReadLastModified(page);

            var targets = new List<OneNoteBridgeTarget>();
            foreach (ConnectedTextTargetRecord target in pageTargets)
            {
                if (!textsBySourceId.TryGetValue(target.SourceId, out string text)) continue;
                targets.Add(new OneNoteBridgeTarget
                {
                    SourceId = target.SourceId,
                    TargetId = target.TargetId,
                    Text = text
                });
            }
            if (!OneNotePageMutations.TryCreateRefresh(
                    page, pageId, targets, out XDocument changes, out replacedCount, out error)) return false;
            if (changes == null) return true;
            return client.TryUpdatePage(changes, expectedLastModified, out error);
        }
    }
}
