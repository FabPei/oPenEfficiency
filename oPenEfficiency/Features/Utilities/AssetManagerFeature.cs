using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnAssetManager",
        Name = "Asset Library",
        Tooltip = "Asset Library (Slides, Images, Shapes)",
        IconData = "M21.54,15L17,15L17,19.54 M7,3.34L7,5C7,6.1 7.9,7 9,7C9,5.9 9.9,5 11,5L14.17,5 M11,21.95L11,18L11,17L2.05,17 M12,2C17.522847498307936,2.0 22.0,6.477152501692064 22.0,12.0C22.0,17.522847498307936 17.522847498307936,22.0 12.0,22.0C6.477152501692064,22.0 2.0,17.522847498307936 2.0,12.0C2.0,6.477152501692064 6.477152501692064,2.0 12.0,2.0",
        Color = "#8B5CF6",
        Description = "Central repository for your reusable slide templates, custom shapes, and high-quality image assets. Searchable and categorized.",
        DetailedHelpText = "Central repository for your reusable slide templates, custom shapes, and high-quality image assets. Searchable and categorized.",
        Keywords = "templates, reusable shapes, image library, slide library, assets",
        MinSelection = 0,
        IsToggle = true,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class AssetManagerFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            Globals.ThisAddIn.ToggleAssetManager();
            return true;
        }
    }
}
