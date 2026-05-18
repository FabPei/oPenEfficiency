using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnPickTextColor",
        Name = "Pick Text Color",
        Tooltip = "Eyedropper: Pick and apply Text Color",
        Color = "#F59E0B",
        Description = "Sample a color from any pixel on the slide/screen and apply it as the Text Color to selected shapes.",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class PickTextColorFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return PickColorFeature.Execute(manager, PickColorFeature.ApplyTarget.Text);
        }
    }
}
