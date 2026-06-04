using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnPickLineColor",
        Name = "Pick Line Color",
        Tooltip = "Eyedropper: Pick and apply Line Color",
        Color = "#3B82F6",
        Description = "Sample a color from any pixel on the slide/screen and apply it as the Line Color to selected shapes.",
        Keywords = "color, picker, line, border, stroke, eyedropper",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class PickLineColorFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return PickColorFeature.Execute(manager, PickColorFeature.ApplyTarget.Line);
        }
    }
}
