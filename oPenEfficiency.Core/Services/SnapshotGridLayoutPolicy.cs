using System;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free grid layout for Insert Slide Snapshots (QA-S9 external-data wave): the
    /// iterative shrink loop that fits the snapshot grid into the available slide area and
    /// the final column count. Production implementation used by SlidePasteWindow.
    /// </summary>
    public static class SnapshotGridLayoutPolicy
    {
        public const float CmToPt = 28.3465f;

        /// <summary>Safety minimum for the snapshot height (~0.5 cm).</summary>
        public const float MinimumHeightPt = 0.5f * CmToPt;

        public const int MaxLayoutIterations = 200;

        /// <summary>The number of columns that fit the given width at the item size and gap.</summary>
        public static int ColumnsThatFit(float availableWidth, float itemWidth, float gap)
        {
            return Math.Max(1, (int)Math.Floor((availableWidth + gap) / (itemWidth + gap)));
        }

        /// <summary>
        /// Fits the grid: starts from the default item size and shrinks height by 5% per
        /// iteration (keeping the width/height ratio) until the rows fit the available
        /// height, bounded by the iteration cap and the safety minimum height. The columns
        /// are recomputed from the final item size. A degenerate empty selection reports the
        /// default size and one column.
        /// </summary>
        public static void FitGrid(
            int itemCount,
            float availableWidth,
            float availableHeight,
            float gap,
            float initialWidth,
            float initialHeight,
            float widthHeightRatio,
            out float itemWidth,
            out float itemHeight,
            out int columns)
        {
            itemWidth = initialWidth;
            itemHeight = initialHeight;

            for (int attempt = 0; attempt < MaxLayoutIterations; attempt++)
            {
                int cols = ColumnsThatFit(availableWidth, itemWidth, gap);
                int rows = (int)Math.Ceiling((double)itemCount / cols);
                float totalHeight = rows * itemHeight + (rows - 1) * gap;
                if (totalHeight <= availableHeight) break;
                itemHeight *= 0.95f;
                itemWidth = itemHeight * widthHeightRatio;
                if (itemHeight < MinimumHeightPt) break;
            }

            columns = ColumnsThatFit(availableWidth, itemWidth, gap);
        }
    }
}
