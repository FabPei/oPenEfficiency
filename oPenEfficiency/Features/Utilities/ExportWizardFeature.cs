using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnExportWizard",
        Name = "Export Wizard",
        Tooltip = "Export Wizard",
        Color = "#10B981",
        Description = "Automates the process of saving individual slides or selections as PDF, PNG, or separate PPTX files.",
        DetailedHelpText = "Automates the process of saving individual slides or selections as PDF, PNG, or separate PPTX files.",
        Keywords = "save, pdf, png, extract, slides, convert",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class ExportWizardFeature
    {
        // Window lifecycle managed by sidebar via ToggleFloatingWindow.
        public static bool Execute(PowerPointManager manager) => false;
    }
}
