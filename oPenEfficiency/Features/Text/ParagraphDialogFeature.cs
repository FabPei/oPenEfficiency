using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnParagraphDialog",
        Name = "Paragraph Dialog",
        Tooltip = "Paragraph Dialog",
        Color = "#FDE047",
        Description = "Opens the native PowerPoint paragraph formatting dialog to adjust line spacing, indentation, and alignment.",
        DetailedHelpText = "Opens the native PowerPoint paragraph formatting dialog.",
        Keywords = "line spacing, indentation, alignment, text format, native dialog",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class ParagraphDialogFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var app = manager?.GetApplication();
                if (app?.CommandBars == null) return false;
                app.CommandBars.ExecuteMso("ParagraphDialog");
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ParagraphDialogFeature.Execute");
                return false;
            }
        }
    }
}
