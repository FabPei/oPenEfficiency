using System;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnExcelLinkManager",
        Name = "Excel Link Manager",
        Tooltip = "Review and manage Excel links in the active presentation",
        IconData = "M14,2H6A2,2 0 0,0 4,4V20A2,2 0 0,0 6,22H18A2,2 0 0,0 20,20V8L14,2M15.8,20H14L12,16.6L10,20H8.2L11.1,15.5L8.2,11H10L12,14.4L14,11H15.8L12.9,15.5L15.8,20M13,9V3.5L18.5,9H13Z",
        Color = "#10B981",
        Description = "Lists linked Excel objects and native linked tables, with refresh, source navigation, relinking, and unlinking actions.",
        DetailedHelpText = "### Excel Link Manager\nOpen the manager to inspect Excel links across the active presentation. Select one or more entries to refresh them or jump to their slide. Native linked tables also support Open Source and Relink Source. Unlink is destructive and keeps the current PowerPoint content while removing its source connection. Embedded charts are intentionally excluded because they have no external Excel link.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class ExcelLinkManagerFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, oPenEfficiency.UI.ExcelLinkManagerStartAction.ShowAll);
        }

        public static bool Execute(PowerPointManager manager, oPenEfficiency.UI.ExcelLinkManagerStartAction startAction)
        {
            if (manager == null) return false;

            try
            {
                var app = manager.GetApplication();
                if (app?.ActivePresentation == null) return false;

                var window = new oPenEfficiency.UI.ExcelLinkManagerWindow(manager, startAction);
                new System.Windows.Interop.WindowInteropHelper(window).Owner = (IntPtr)app.HWND;
                window.ShowDialog();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelLinkManagerFeature.Execute");
                throw;
            }
        }
    }
}
