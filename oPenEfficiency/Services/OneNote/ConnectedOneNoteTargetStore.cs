using System;
using System.Collections.Generic;
using System.Linq;
using oPenEfficiency.OneNote;

namespace oPenEfficiency.Services.OneNote
{
    /// <summary>
    /// Persistence for connected-text target records (FEAT-007 ON-S0).
    /// Owns the legacy connected_text_labels.json store (Word and OneNote records
    /// share the file). The record shape and file name are compatibility contracts:
    /// existing records remain readable and disconnect keeps destination content.
    /// </summary>
    public static class ConnectedOneNoteTargetStore
    {
        private const string StoreFileName = "connected_text_labels.json";

        public static List<ConnectedTextTargetRecord> Load()
        {
            return JsonService.Load<List<ConnectedTextTargetRecord>>(StoreFileName)
                ?? new List<ConnectedTextTargetRecord>();
        }

        /// <summary>Appends the record unless an identical target already exists; persists the file.</summary>
        public static void Save(ConnectedTextTargetRecord record)
        {
            if (record == null) return;
            var records = Load();
            bool exists = records.Any(item =>
                string.Equals(item.Host, record.Host, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(item.DocumentPath, record.DocumentPath, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(item.TargetId, record.TargetId, StringComparison.OrdinalIgnoreCase));
            if (!exists) records.Add(record);
            if (!JsonService.TrySave(records, StoreFileName))
                throw new InvalidOperationException("Connected-text target metadata could not be saved.");
        }

        /// <summary>
        /// Removes every record of a source. Destination content is deliberately left
        /// untouched — disconnect semantics keep what was sent.
        /// </summary>
        public static void RemoveBySource(string sourceId)
        {
            var records = Load();
            records.RemoveAll(item => string.Equals(item.SourceId, sourceId, StringComparison.OrdinalIgnoreCase));
            JsonService.TrySave(records, StoreFileName);
        }
    }
}
