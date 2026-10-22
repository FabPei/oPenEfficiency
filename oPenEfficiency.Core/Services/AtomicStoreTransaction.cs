using System;
using System.Collections.Generic;

namespace oPenEfficiency.Services
{
    /// <summary>Writes a replacement before deleting old values so a failed write cannot destroy known-good data.</summary>
    public static class AtomicStoreTransaction
    {
        public static void AddBeforeDelete(Action addNewValue, IEnumerable<Action> deleteOldValues)
        {
            if (addNewValue == null) throw new ArgumentNullException(nameof(addNewValue));
            if (deleteOldValues == null) throw new ArgumentNullException(nameof(deleteOldValues));

            addNewValue();
            foreach (Action deleteOldValue in deleteOldValues)
            {
                if (deleteOldValue == null) throw new ArgumentException("A delete operation is null.", nameof(deleteOldValues));
                deleteOldValue();
            }
        }
    }
}
