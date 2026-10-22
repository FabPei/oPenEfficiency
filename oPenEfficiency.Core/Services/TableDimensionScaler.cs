using System;
using System.Collections.Generic;

namespace oPenEfficiency.Services
{
    public static class TableDimensionScaler
    {
        public static float[] ScaleToTotal(IReadOnlyList<float> dimensions, float targetTotal)
        {
            if (dimensions == null) throw new ArgumentNullException(nameof(dimensions));
            if (dimensions.Count == 0) throw new ArgumentException("At least one dimension is required.", nameof(dimensions));
            if (targetTotal <= 0f) throw new ArgumentOutOfRangeException(nameof(targetTotal));

            double currentTotal = 0d;
            for (int index = 0; index < dimensions.Count; index++)
            {
                if (dimensions[index] <= 0f)
                    throw new ArgumentOutOfRangeException(nameof(dimensions), "Dimensions must be positive.");
                currentTotal += dimensions[index];
            }

            var scaled = new float[dimensions.Count];
            double assigned = 0d;
            for (int index = 0; index < dimensions.Count - 1; index++)
            {
                scaled[index] = (float)(dimensions[index] / currentTotal * targetTotal);
                assigned += scaled[index];
            }

            scaled[scaled.Length - 1] = (float)(targetTotal - assigned);
            if (scaled[scaled.Length - 1] <= 0f)
                throw new InvalidOperationException("The target total cannot produce positive dimensions.");
            return scaled;
        }
    }
}
