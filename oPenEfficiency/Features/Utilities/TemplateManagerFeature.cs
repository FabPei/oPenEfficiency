using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnTemplateManager",
        Name = "Template Manager",
        Tooltip = "PowerPoint Template Manager",
        Color = "#3B82F6",
        Description = "Manage and quickly open PowerPoint templates from default locations and custom directories.",
        DetailedHelpText = "Manage and quickly open PowerPoint templates from default locations and custom directories.\n\nFeatures:\n• Discover default Windows templates\n• Add watched folders\n• Add individual template files\n• Open as new presentation",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class TemplateManagerFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var win = new oPenEfficiency.UI.Dialogs.TemplateManagerWindow(manager);
            win.ShowDialog();
            return true;
        }
    }
}
