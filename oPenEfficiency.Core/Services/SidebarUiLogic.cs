using System;
using System.Collections.Generic;
using System.Linq;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Services
{
    public sealed class SidebarSelectionSnapshot
    {
        public OfficeSelectionType Type { get; set; }
        public int ShapeCount { get; set; }
        public int TextShapeCount { get; set; }
        public int TextOrTableShapeCount { get; set; }
        public bool HasTextRange { get; set; }
        public bool HasTable { get; set; }
        public bool HasHiddenShapes { get; set; }
    }

    public sealed class SidebarFeatureStateRule
    {
        public int MinSelection { get; set; }
        public int MaxSelection { get; set; }
        public OfficeSelectionType RequiredType { get; set; }
        public bool RequiresTable { get; set; }
        public bool RequiresHiddenShapes { get; set; }
    }

    public static class SidebarStateEvaluator
    {
        public static bool IsEnabled(
            SidebarFeatureStateRule rule,
            SidebarSelectionSnapshot selection)
        {
            if (rule == null) return false;
            if (rule.RequiresHiddenShapes && selection?.HasHiddenShapes != true) return false;

            if (selection == null)
                return rule.MinSelection == 0 && !rule.RequiresTable && !rule.RequiresHiddenShapes;

            if (rule.RequiredType == OfficeSelectionType.Shapes &&
                selection.Type != OfficeSelectionType.Shapes &&
                selection.Type != OfficeSelectionType.Text)
                return false;

            if (rule.RequiresTable && !selection.HasTable) return false;

            if (selection.Type == OfficeSelectionType.Shapes ||
                selection.Type == OfficeSelectionType.Text)
            {
                int count = selection.ShapeCount;
                if (selection.Type == OfficeSelectionType.Text && count == 0) count = 1;
                if (rule.MinSelection > 0 && count < rule.MinSelection) return false;
                if (rule.MaxSelection > 0 && count > rule.MaxSelection) return false;
            }

            return true;
        }

        public static bool HasUsableTextTarget(
            string featureId,
            SidebarSelectionSnapshot selection)
        {
            if (selection == null) return false;
            if (selection.Type == OfficeSelectionType.Text)
                return featureId != "BtnSplitByParagraphs" || selection.HasTextRange;
            if (selection.Type != OfficeSelectionType.Shapes) return false;

            int targetCount = featureId == "BtnApplyTextTool"
                ? selection.TextOrTableShapeCount
                : selection.TextShapeCount;
            if (featureId == "BtnSpecialChars")
                return selection.ShapeCount == 1 && targetCount == 1;
            if (featureId == "BtnSplitByParagraphs")
                return selection.ShapeCount == 1 ? targetCount == 1 : targetCount >= 2;
            return targetCount > 0;
        }

        public static bool? ResolveAggregateToggleState(
            bool isEnabled,
            int itemCount,
            int activeItemCount)
        {
            if (!isEnabled || itemCount <= 0 || activeItemCount <= 0) return false;
            if (activeItemCount >= itemCount) return true;
            return null;
        }
    }

    public static class SidebarSearchService
    {
        public static IReadOnlyList<FeatureCatalogRecord> Search(
            IEnumerable<FeatureCatalogRecord> catalog,
            string searchTerm)
        {
            string query = searchTerm?.Trim();
            if (string.IsNullOrEmpty(query)) return Array.Empty<FeatureCatalogRecord>();

            return (catalog ?? Enumerable.Empty<FeatureCatalogRecord>())
                .Where(feature => feature != null &&
                    feature.UserVisible &&
                    !feature.IsObsolete &&
                    feature.ExecutionKind == FeatureExecutionKind.Command)
                .Select(feature => new
                {
                    Feature = feature,
                    Score = FeatureCatalogSearch.GetMatchScore(feature, query)
                })
                .Where(result => result.Score < int.MaxValue)
                .OrderBy(result => result.Score)
                .ThenBy(result => result.Feature.Name, StringComparer.OrdinalIgnoreCase)
                .Select(result => result.Feature)
                .ToList();
        }
    }
}
