using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnMagicResizer",
        Name = "Magic Resizer",
        Tooltip = "Magic Resizer",
        IconData = "M136,112L48,112L48,200L136,200L136,120ZM128,200L56,200L56,128L128,128ZM216,184L216,200L176,200L200,200L200,184ZM216,112L216,144L216,112ZM216,56L216,72L216,56L184,56L200,56ZM152,48L112,48L144,48ZM40,80L40,56L72,56L56,56L56,80Z",
        Color = "#6366F1",
        Description = "Propagates a resize transformation across multiple objects relative to their individual centers, preserving layout spacing.",
        DetailedHelpText = "Propagates a resize transformation across multiple objects relative to their individual centers, preserving layout spacing.",
        Keywords = "scale, transform, relative, individual, size, maintain",
        MinSelection = 1,
        IsToggle = true,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class MagicResizerFeature
    {
        // Window lifecycle managed by sidebar via ToggleFloatingWindow.
        public static bool Execute(PowerPointManager manager) => false;
    }
}
