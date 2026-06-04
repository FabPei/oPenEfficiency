using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnCloneTop",
        Name = "Clone Top",
        Tooltip = "Clone Top",
        IconData = "M19,19H5V5H19M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M12,16H10V12H6V10H10V6H12V10H16V12H12",
        Color = "#FBBF24",
        Description = "Duplicates the selection and places the copy adjacent to the top edge.",
        DetailedHelpText = "Duplicates the selection and places the copy adjacent to the top edge.",
        Keywords = "clone, duplicate, copy, above, top",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class CloneTopFeature
    {
        public static bool Execute(PowerPointManager manager)
            => CloneSelectionFeature.Execute(manager, CloneSelectionFeature.Direction.Top);
    }
}
