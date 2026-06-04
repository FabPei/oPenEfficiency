using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnExploreFeatures",
        Name = "Explore Features",
        Tooltip = "Explore Features",
        IconData = "M12,2A10,10 0 1,0 22,12A10,10 0 0,0 12,2M12,17A1,1 0 1,1 11,16A1,1 0 0,1 12,17M13,14H11V9H13V14Z",
        Color = "#6366F1",
        Description = "The Feature Explorer provides a searchable index of all features available in oPenEfficiency. It shows detailed usage instructions, animated examples, and current shortcuts.",
        DetailedHelpText = "The Feature Explorer provides a searchable index of all features available in oPenEfficiency. It shows detailed usage instructions, animated examples, and current shortcuts.",
        Keywords = "search, index, help, discover, guide, manual",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class ExploreFeaturesFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            Globals.ThisAddIn.OpenFeatureExplorer();
            return true;
        }
    }
}
