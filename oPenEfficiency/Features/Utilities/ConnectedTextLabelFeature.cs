using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.OneNote;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Services.OneNote;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnConnectedTextLabel",
        Name = "Insert Linked Text Label",
        Tooltip = "Send selected PowerPoint text to Word or OneNote and keep it refreshable",
        SearchAliases = new[] { "Connected Text Labels", "Linked Text", "Word Label", "OneNote Label", "Connected Label" },
        IconData = "M3,5A2,2 0 0,1 5,3H11V5H5V19H11V21H5A2,2 0 0,1 3,19V5M13,7H21V9H18V17H21V19H13V17H16V9H13V7Z",
        Color = "#2563EB",
        Description = "Publishes selected PowerPoint shape text as a refreshable label in an open Word document or OneNote page.",
        DetailedHelpText = "### Insert Linked Text Label\nSelect exactly one PowerPoint shape containing text. The primary action sends it to the active Word document; if Word is unavailable, it tries OneNote. Right-click to choose Word or OneNote explicitly, refresh every connected label whose PowerPoint source is in the active presentation, or disconnect the selected source. Word labels are stored as tagged content controls. OneNote labels are stored as tagged page elements. OneNote sends stop without writing when the visible current page cannot be identified or its version changes during the update. Refresh updates open Word documents and remembered OneNote pages; it never opens or saves Office files without an explicit send action. The link stores text only and does not copy formatting, images, tables, notes, or hyperlinks.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class ConnectedTextLabelFeature
    {
        private const string SourceTag = "oPE_ConnectedTextSourceId";
        private const string WordTagPrefix = "oPE_ConnectedText:";

        public enum TargetHost
        {
            Word,
            OneNote
        }

        public static bool Execute(PowerPointManager manager)
        {
            if (!TryGetSelectedSource(manager, out Shape source, out string sourceId, out string text, out bool sourceTagCreated))
                return false;
            bool targetCreated = false;
            try
            {
                if (TrySendToWord(sourceId, text, out ConnectedTextTargetRecord record))
                {
                    targetCreated = true;
                    ConnectedOneNoteTargetStore.Save(record);
                    ShowResult("Created a connected text label in the active Word document.");
                    return true;
                }
                if (TrySendToOneNote(sourceId, text, out record, out string oneNoteError))
                {
                    targetCreated = true;
                    ConnectedOneNoteTargetStore.Save(record);
                    ShowResult("Created a connected text label on the current OneNote page.");
                    return true;
                }
                if (sourceTagCreated) ShapeMetadataService.RemoveTag(source, SourceTag);
                ShowError(oneNoteError ?? "Open a Word document or OneNote page, then run Insert Linked Text Label again.");
                return false;
            }
            catch (Exception ex)
            {
                if (!targetCreated && sourceTagCreated) ShapeMetadataService.RemoveTag(source, SourceTag);
                ExceptionLogger.Log(ex, "ConnectedTextLabelFeature.Execute");
                ShowError("The connected label could not be created.");
                return false;
            }
        }

        public static bool SendTo(PowerPointManager manager, TargetHost host)
        {
            if (!TryGetSelectedSource(manager, out Shape source, out string sourceId, out string text, out bool sourceTagCreated))
                return false;
            bool targetCreated = false;
            try
            {
                ConnectedTextTargetRecord record;
                string targetError = null;
                bool sent;
                if (host == TargetHost.Word)
                    sent = TrySendToWord(sourceId, text, out record);
                else
                    sent = TrySendToOneNote(sourceId, text, out record, out targetError);
                if (!sent)
                {
                    if (sourceTagCreated) ShapeMetadataService.RemoveTag(source, SourceTag);
                    ShowError(targetError ?? $"Open the target {host} document or page before sending the label.");
                    return false;
                }
                targetCreated = true;
                ConnectedOneNoteTargetStore.Save(record);
                ShowResult($"Created a connected text label in {host}.");
                return true;
            }
            catch (Exception ex)
            {
                if (!targetCreated && sourceTagCreated) ShapeMetadataService.RemoveTag(source, SourceTag);
                ExceptionLogger.Log(ex, "ConnectedTextLabelFeature.SendTo." + host);
                ShowError($"The connected label could not be created in {host}.");
                return false;
            }
        }

        public static bool RefreshAll(PowerPointManager manager)
        {
            try
            {
                var sources = GetPresentationSources(manager);
                if (sources.Count == 0)
                {
                    ShowResult("No connected PowerPoint text sources were found in this presentation.");
                    return false;
                }

                int wordUpdates = RefreshWordTargets(sources);
                int oneNoteUpdates = RefreshOneNoteTargets(sources);
                ShowResult($"Updated {wordUpdates + oneNoteUpdates} connected label(s) " +
                           $"({wordUpdates} Word, {oneNoteUpdates} OneNote).");
                return wordUpdates + oneNoteUpdates > 0;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ConnectedTextLabelFeature.RefreshAll");
                ShowError("Connected labels could not be refreshed completely.");
                return false;
            }
        }

        public static bool DisconnectSelected(PowerPointManager manager)
        {
            var shapes = manager?.GetSelectedShapes();
            if (shapes == null || shapes.Count != 1) return false;
            Shape source = shapes[1];
            string sourceId = ShapeMetadataService.GetTag(source, SourceTag);
            if (string.IsNullOrEmpty(sourceId))
            {
                ShowResult("The selected shape is not a connected text source.");
                return false;
            }
            ShapeMetadataService.RemoveTag(source, SourceTag);
            ConnectedOneNoteTargetStore.RemoveBySource(sourceId);
            ShowResult("Disconnected the selected PowerPoint source. Existing Word and OneNote text remains unchanged.");
            return true;
        }

        private static bool TryGetSelectedSource(
            PowerPointManager manager,
            out Shape source,
            out string sourceId,
            out string text,
            out bool sourceTagCreated)
        {
            source = null;
            sourceId = null;
            text = null;
            sourceTagCreated = false;
            try
            {
                var shapes = manager?.GetSelectedShapes();
                if (shapes == null || shapes.Count != 1)
                {
                    ShowError("Select exactly one PowerPoint shape containing text.");
                    return false;
                }
                source = shapes[1];
                if (source.HasTextFrame != Office.MsoTriState.msoTrue ||
                    source.TextFrame.HasText != Office.MsoTriState.msoTrue)
                {
                    ShowError("The selected shape does not contain text.");
                    return false;
                }
                text = source.TextFrame.TextRange.Text?.TrimEnd('\r', '\n');
                if (string.IsNullOrWhiteSpace(text))
                {
                    ShowError("The selected shape does not contain visible text.");
                    return false;
                }
                sourceId = ShapeMetadataService.GetTag(source, SourceTag);
                if (string.IsNullOrEmpty(sourceId))
                {
                    sourceId = Guid.NewGuid().ToString("N");
                    ShapeMetadataService.SetTag(source, SourceTag, sourceId);
                    sourceTagCreated = true;
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ConnectedTextLabelFeature.TryGetSelectedSource");
                return false;
            }
        }

        private static bool TrySendToWord(string sourceId, string text, out ConnectedTextTargetRecord record)
        {
            record = null;
            try
            {
                dynamic word = Marshal.GetActiveObject("Word.Application");
                dynamic document = word.ActiveDocument;
                dynamic selection = word.Selection;
                if (document == null || selection == null) return false;
                dynamic range = selection.Range;
                range.Text = text;
                dynamic control = document.ContentControls.Add(0, range); // wdContentControlRichText
                control.Title = "openEfficiency Connected Text";
                control.Tag = WordTagPrefix + sourceId;
                string path = string.Empty;
                try { path = document.FullName as string ?? string.Empty; }
                catch (COMException) { path = string.Empty; }
                record = new ConnectedTextTargetRecord
                {
                    SourceId = sourceId,
                    Host = TargetHost.Word.ToString(),
                    DocumentPath = path,
                    TargetId = control.Tag as string
                };
                return true;
            }
            catch (COMException) { return false; }
            catch (InvalidComObjectException) { return false; }
        }

        /// <summary>
        /// Sends to the user-visible active OneNote page, resolved through
        /// Windows.CurrentWindow.CurrentPageId (fail-closed; see OneNoteDesktopClient).
        /// </summary>
        private static bool TrySendToOneNote(
            string sourceId,
            string text,
            out ConnectedTextTargetRecord record,
            out string error)
        {
            record = null;
            error = null;
            if (OneNoteDesktopClient.TryAttach(out OneNoteDesktopClient client, out string attachError))
            {
                using (client)
                {
                    if (client.TryResolveActivePageId(out string directPageId, out string resolveError))
                    {
                        if (!OneNotePageUpdateService.TryAppendConnectedText(client, directPageId, sourceId, text, out error))
                            return false;
                        record = CreateOneNoteRecord(sourceId, directPageId);
                        return true;
                    }
                    error = resolveError;
                }
            }

            var bridge = new OneNoteBridgeClient();
            if (!bridge.TryGetActivePageId(out string pageId, out string bridgeError) ||
                !bridge.TryAppendConnectedText(pageId, sourceId, text, out bridgeError))
            {
                error = bridgeError ?? error ?? attachError;
                return false;
            }

            record = CreateOneNoteRecord(sourceId, pageId);
            return true;
        }

        private static ConnectedTextTargetRecord CreateOneNoteRecord(string sourceId, string pageId) =>
            new ConnectedTextTargetRecord
            {
                SourceId = sourceId,
                Host = TargetHost.OneNote.ToString(),
                DocumentPath = pageId,
                TargetId = sourceId
            };

        private static Dictionary<string, string> GetPresentationSources(PowerPointManager manager)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var presentation = manager?.GetApplication()?.ActivePresentation;
            if (presentation == null) return result;
            foreach (Slide slide in presentation.Slides)
                for (int index = 1; index <= slide.Shapes.Count; index++)
                    AddSourceRecursive(slide.Shapes[index], result);
            return result;
        }

        private static void AddSourceRecursive(Shape shape, IDictionary<string, string> result)
        {
            string sourceId = ShapeMetadataService.GetTag(shape, SourceTag);
            if (!string.IsNullOrEmpty(sourceId) &&
                shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                result[sourceId] = shape.TextFrame.TextRange.Text?.TrimEnd('\r', '\n') ?? string.Empty;
            if (shape.Type != Office.MsoShapeType.msoGroup) return;
            for (int index = 1; index <= shape.GroupItems.Count; index++)
                AddSourceRecursive(shape.GroupItems[index], result);
        }

        private static int RefreshWordTargets(IReadOnlyDictionary<string, string> sources)
        {
            int changed = 0;
            try
            {
                dynamic word = Marshal.GetActiveObject("Word.Application");
                foreach (dynamic document in word.Documents)
                {
                    foreach (dynamic control in document.ContentControls)
                    {
                        string tag = control.Tag as string;
                        if (string.IsNullOrEmpty(tag) || !tag.StartsWith(WordTagPrefix, StringComparison.OrdinalIgnoreCase)) continue;
                        string sourceId = tag.Substring(WordTagPrefix.Length);
                        if (!sources.TryGetValue(sourceId, out string text)) continue;
                        control.Range.Text = text;
                        changed++;
                    }
                }
            }
            catch (COMException) { return changed; }
            catch (InvalidComObjectException) { return changed; }
            return changed;
        }

        private static int RefreshOneNoteTargets(IReadOnlyDictionary<string, string> sources)
        {
            var records = ConnectedOneNoteTargetStore.Load()
                .Where(item => string.Equals(item.Host, TargetHost.OneNote.ToString(), StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (records.Count == 0) return 0;
            int changed = 0;
            var bridge = new OneNoteBridgeClient();
            bool bridgeAvailable = bridge.TryGetActivePageId(out _, out string bridgeAvailabilityError);
            if (bridgeAvailable)
            {
                foreach (var pageGroup in records.GroupBy(item => item.DocumentPath, StringComparer.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrEmpty(pageGroup.Key)) continue;
                    var targets = pageGroup
                        .Where(item => sources.ContainsKey(item.SourceId))
                        .Select(item => new OneNoteBridgeTarget
                        {
                            SourceId = item.SourceId,
                            TargetId = item.TargetId,
                            Text = sources[item.SourceId]
                        })
                        .ToList();
                    if (!bridge.TryRefreshConnectedTexts(pageGroup.Key, targets, out int replaced, out string error))
                    {
                        ShowError(error ?? ("A OneNote page could not be refreshed: " + pageGroup.Key));
                        continue;
                    }
                    changed += replaced;
                }
                return changed;
            }

            if (!OneNoteDesktopClient.TryAttach(out OneNoteDesktopClient client, out string attachError))
            {
                ShowError(bridgeAvailabilityError ?? attachError ?? "The running OneNote instance could not be connected.");
                return changed;
            }

            using (client)
            {
                foreach (var pageGroup in records.GroupBy(item => item.DocumentPath, StringComparer.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrEmpty(pageGroup.Key)) continue;
                    if (!OneNotePageUpdateService.TryReplacePageTexts(
                            client, pageGroup.Key, pageGroup.ToList(), sources,
                            out int replaced, out string error))
                    {
                        ShowError(error ?? ("A OneNote page could not be refreshed: " + pageGroup.Key));
                        continue;
                    }
                    changed += replaced;
                }
            }
            return changed;
        }

        private static void ShowResult(string message)
        {
            System.Windows.Forms.MessageBox.Show(message, "Insert Linked Text Label",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        private static void ShowError(string message)
        {
            System.Windows.Forms.MessageBox.Show(message, "Insert Linked Text Label",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning);
        }
    }
}
