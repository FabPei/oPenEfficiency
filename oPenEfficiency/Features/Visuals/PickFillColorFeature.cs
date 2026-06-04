using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnPickFillColor",
        Name = "Pick Fill Color",
        Tooltip = "Eyedropper: Pick and apply Fill Color",
        Color = "#06B6D4",
        Description = "Sample a color from any pixel on the slide/screen and apply it as the Fill Color to selected shapes.",
        Keywords = "color, picker, fill, eyedropper, sample",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class PickFillColorFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return PickColorFeature.Execute(manager, PickColorFeature.ApplyTarget.Fill);
        }
    }
}
