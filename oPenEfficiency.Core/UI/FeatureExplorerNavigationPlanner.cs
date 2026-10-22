using System;

namespace oPenEfficiency.UI
{
    /// <summary>
    /// Host-neutral layout state for Feature Explorer category navigation.
    /// The host capacity deliberately excludes integrated chevron controls so
    /// their visibility cannot create an overflow hysteresis loop.
    /// </summary>
    public sealed class FeatureExplorerNavigationState
    {
        internal FeatureExplorerNavigationState(
            bool requiresNavigation,
            double viewportWidth,
            double horizontalOffset,
            double scrollableWidth)
        {
            RequiresNavigation = requiresNavigation;
            ViewportWidth = viewportWidth;
            HorizontalOffset = horizontalOffset;
            ScrollableWidth = scrollableWidth;
        }

        public bool RequiresNavigation { get; }
        public double ViewportWidth { get; }
        public double HorizontalOffset { get; }
        public double ScrollableWidth { get; }
        public bool CanScrollLeft { get { return RequiresNavigation && HorizontalOffset > FeatureExplorerNavigationPlanner.Epsilon; } }
        public bool CanScrollRight { get { return RequiresNavigation && HorizontalOffset < ScrollableWidth - FeatureExplorerNavigationPlanner.Epsilon; } }
    }

    public static class FeatureExplorerNavigationPlanner
    {
        public const double Epsilon = 0.5;

        /// <summary>
        /// Creates navigation state from the total category content width, the complete category-row width
        /// before chevrons consume space, and the viewport currently rendered by the host.
        /// </summary>
        public static FeatureExplorerNavigationState Create(
            double contentWidth,
            double hostCapacity,
            double renderedViewportWidth,
            double horizontalOffset)
        {
            double normalizedContentWidth = NonNegative(contentWidth);
            double normalizedHostCapacity = NonNegative(hostCapacity);
            bool requiresNavigation = normalizedHostCapacity > 0 &&
                normalizedContentWidth > normalizedHostCapacity + Epsilon;

            if (!requiresNavigation)
            {
                return new FeatureExplorerNavigationState(false, normalizedHostCapacity, 0, 0);
            }

            double viewportWidth = renderedViewportWidth > Epsilon
                ? Math.Min(normalizedHostCapacity, renderedViewportWidth)
                : normalizedHostCapacity;
            double scrollableWidth = Math.Max(0, normalizedContentWidth - viewportWidth);
            double normalizedOffset = Clamp(horizontalOffset, 0, scrollableWidth);
            return new FeatureExplorerNavigationState(true, viewportWidth, normalizedOffset, scrollableWidth);
        }

        /// <summary>
        /// Returns the smallest valid scroll offset that exposes the complete category item when possible.
        /// </summary>
        public static double CalculateRevealOffset(
            double contentWidth,
            double viewportWidth,
            double currentOffset,
            double itemStart,
            double itemWidth)
        {
            double normalizedContentWidth = NonNegative(contentWidth);
            double normalizedViewportWidth = Math.Min(normalizedContentWidth, NonNegative(viewportWidth));
            double scrollableWidth = Math.Max(0, normalizedContentWidth - normalizedViewportWidth);
            if (scrollableWidth <= Epsilon || normalizedViewportWidth <= Epsilon)
            {
                return 0;
            }

            double normalizedOffset = Clamp(currentOffset, 0, scrollableWidth);
            double normalizedStart = Clamp(itemStart, 0, normalizedContentWidth);
            double normalizedEnd = Clamp(normalizedStart + NonNegative(itemWidth), normalizedStart, normalizedContentWidth);
            if (normalizedStart < normalizedOffset - Epsilon)
            {
                return Clamp(normalizedStart, 0, scrollableWidth);
            }

            if (normalizedEnd > normalizedOffset + normalizedViewportWidth + Epsilon)
            {
                return Clamp(normalizedEnd - normalizedViewportWidth, 0, scrollableWidth);
            }

            return normalizedOffset;
        }

        private static double NonNegative(double value)
        {
            return double.IsNaN(value) || double.IsInfinity(value) || value < 0 ? 0 : value;
        }

        private static double Clamp(double value, double minimum, double maximum)
        {
            if (double.IsNaN(value) || double.IsInfinity(value)) return minimum;
            return Math.Max(minimum, Math.Min(maximum, value));
        }
    }
}
