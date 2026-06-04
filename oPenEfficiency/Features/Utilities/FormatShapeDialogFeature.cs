using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnFormatShapeDialog",
        Name = "Format Shape Dialog",
        Tooltip = "Open Format Shape",
        IconData = "M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V5h14v14z M7 7h10v2H7V7zm0 4h10v2H7v-2zm0 4h7v2H7v-2z",
        Color = "#8B5CF6",
        Description = "Opens the native PowerPoint 'Format Shape' pane to adjust fill, line, and effects.",
        DetailedHelpText = "### Format Shape Dialog\nOpens a consolidated formatting dialog combining fill, line, shadow, size, and position options in a single panel.",
        Keywords = "pane, fill, line, shadow, properties, native",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class FormatShapeDialogFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app != null && app.CommandBars != null)
                {
                    app.CommandBars.ExecuteMso("ObjectFormatDialog");
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FormatShapeDialogFeature.Execute");
                return false;
            }
        }
    }
}
