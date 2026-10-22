using System;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free stacking for the four dock features (QA-S9 align-and-size wave): the
    /// zero-gap stack cursor that places every non-anchor shape against the reference
    /// shape's edge in selection order. Production implementation used by the
    /// DockToFirst*Feature family.
    /// </summary>
    public static class DockStackPolicy
    {
        public enum DockDirection
        {
            Left,
            Right,
            Above,
            Below
        }

        /// <summary>
        /// Computes the Left (horizontal directions) or Top (vertical directions) position
        /// of each non-anchor shape, in the given selection order. Left and Above start from
        /// the anchor edge and displace each shape by its own size (the cursor follows the
        /// placed edge); Right and Below place the first shape exactly on the anchor edge
        /// and advance the cursor by each size. Shapes stack with zero gap; their other
        /// coordinate, size and formatting are untouched.
        /// </summary>
        public static float[] StackSequence(
            DockDirection direction,
            float refLeft,
            float refTop,
            float refWidth,
            float refHeight,
            float[] orderedSizes)
        {
            var positions = new float[orderedSizes.Length];
            float cursor;
            int i;

            switch (direction)
            {
                case DockDirection.Left:
                    cursor = refLeft;
                    for (i = 0; i < orderedSizes.Length; i++)
                    {
                        cursor -= orderedSizes[i];
                        positions[i] = cursor;
                    }
                    break;
                case DockDirection.Above:
                    cursor = refTop;
                    for (i = 0; i < orderedSizes.Length; i++)
                    {
                        cursor -= orderedSizes[i];
                        positions[i] = cursor;
                    }
                    break;
                case DockDirection.Right:
                    cursor = refLeft + refWidth;
                    for (i = 0; i < orderedSizes.Length; i++)
                    {
                        positions[i] = cursor;
                        cursor += orderedSizes[i];
                    }
                    break;
                default: // Below
                    cursor = refTop + refHeight;
                    for (i = 0; i < orderedSizes.Length; i++)
                    {
                        positions[i] = cursor;
                        cursor += orderedSizes[i];
                    }
                    break;
            }
            return positions;
        }
    }
}
