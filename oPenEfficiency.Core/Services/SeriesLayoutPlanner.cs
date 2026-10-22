using System;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free layout for Generate Shape Series (QA-S9 wave 5): the required series extent per
    /// direction, the shrink-only auto-fit scale against the available bounds, and the per-item
    /// placement for horizontal, vertical and matrix runs. Production implementation used by
    /// SeriesGeneratorFeature.
    /// </summary>
    public static class SeriesLayoutPlanner
    {
        public const string DirectionHorizontal = "Horizontal";
        public const string DirectionVertical = "Vertical";
        public const string DirectionMatrix = "Matrix";

        /// <summary>
        /// Computes the extent the full series occupies: items plus the gaps between them, with
        /// matrix columns capped by the item count and rows ceiled. An unknown direction reports
        /// a zero extent, matching the historical auto-fit no-op.
        /// </summary>
        public static void RequiredSize(
            string direction,
            int itemCount,
            int matrixColumns,
            float itemWidth,
            float itemHeight,
            float spacingX,
            float spacingY,
            out float requiredWidth,
            out float requiredHeight)
        {
            requiredWidth = 0f;
            requiredHeight = 0f;
            if (itemCount <= 0) return;

            if (direction == DirectionHorizontal)
            {
                requiredWidth = (itemCount * itemWidth) + ((itemCount - 1) * spacingX);
                requiredHeight = itemHeight;
            }
            else if (direction == DirectionVertical)
            {
                requiredWidth = itemWidth;
                requiredHeight = (itemCount * itemHeight) + ((itemCount - 1) * spacingY);
            }
            else if (direction == DirectionMatrix && matrixColumns > 0)
            {
                int columns = Math.Min(itemCount, matrixColumns);
                int rows = (int)Math.Ceiling(itemCount / (double)matrixColumns);
                requiredWidth = (columns * itemWidth) + ((columns - 1) * spacingX);
                requiredHeight = (rows * itemHeight) + ((rows - 1) * spacingY);
            }
        }

        /// <summary>
        /// Scales the planned extent into the available bounds, shrinking only on the binding
        /// axis and never above 1 (auto-fit reduces, it does not enlarge).
        /// </summary>
        public static float AutoFitScale(
            float requiredWidth,
            float requiredHeight,
            float availableWidth,
            float availableHeight)
        {
            float scaleX = requiredWidth > availableWidth && requiredWidth > 0 ? availableWidth / requiredWidth : 1f;
            float scaleY = requiredHeight > availableHeight && requiredHeight > 0 ? availableHeight / requiredHeight : 1f;
            return Math.Min(1f, Math.Min(scaleX, scaleY));
        }

        /// <summary>
        /// Places the item at the given series index: a horizontal or vertical run steps by the
        /// item size plus spacing, a matrix wraps every matrixColumns-th item into the next row.
        /// An unknown direction or a matrix without columns anchors every item at the start,
        /// matching the historical behavior.
        /// </summary>
        public static void PlaceItem(
            string direction,
            int index,
            int matrixColumns,
            float startX,
            float startY,
            float itemWidth,
            float itemHeight,
            float spacingX,
            float spacingY,
            out float left,
            out float top)
        {
            left = startX;
            top = startY;

            if (direction == DirectionHorizontal)
            {
                left = startX + (index * (itemWidth + spacingX));
            }
            else if (direction == DirectionVertical)
            {
                top = startY + (index * (itemHeight + spacingY));
            }
            else if (direction == DirectionMatrix && matrixColumns > 0)
            {
                int column = index % matrixColumns;
                int row = index / matrixColumns;
                left = startX + (column * (itemWidth + spacingX));
                top = startY + (row * (itemHeight + spacingY));
            }
        }
    }
}
