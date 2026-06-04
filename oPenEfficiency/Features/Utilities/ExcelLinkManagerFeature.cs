using System;
using oPenEfficiency.Services.Attributes;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnExcelLinkManager",
        Name = "Excel Link Manager",
        Tooltip = "Advanced Excel Link Manager",
        IconData = "M14,2H6A2,2 0 0,0 4,4V20A2,2 0 0,0 6,22H18A2,2 0 0,0 20,20V8L14,2M15.8,20H14L12,16.6L10,20H8.2L11.1,15.5L8.2,11H10L12,14.4L14,11H15.8L12.9,15.5L15.8,20M13,9V3.5L18.5,9H13Z",
        Color = "#10B981",
        Description = "Manage all Excel links with Split Screen and Relative Paths.",
        DetailedHelpText = "Manage your Excel links.",
        Keywords = "spreadsheet, sync, update, paths, data, connection",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class ExcelLinkManagerFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            var window = new oPenEfficiency.UI.ExcelLinkManagerWindow(manager);
            window.ShowDialog();
            return true;
        }
    }
}
