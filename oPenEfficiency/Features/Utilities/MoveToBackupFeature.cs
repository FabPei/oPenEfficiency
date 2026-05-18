using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnMoveToBackup",
        Name = "Move to Backup",
        Tooltip = "Move selected slides to the end behind the backup slide",
        IconData = "M14,3V5H17.59L7.76,14.83L9.17,16.24L19,6.41V10H21V3M19,19H5V5H12V3H5C3.89,3 3,3.9 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V12H19V19Z",
        Color = "#F59E0B",
        Description = "Moves the selected slides to the end of the presentation.",
        DetailedHelpText = "Moves selected slides behind the Backup slide.",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionSlides)]
    public static class MoveToBackupFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;
                var selection = app.ActiveWindow.Selection;
                if (selection == null || selection.Type != PpSelectionType.ppSelectionSlides) return false;
                
                var pres = app.ActivePresentation;
                int targetIndex = pres.Slides.Count;
                
                // Find backup slide
                for (int i = pres.Slides.Count; i >= 1; i--)
                {
                    try {
                        if (pres.Slides[i].Tags["EE4P_BACKUP_SLIDE"] == "True")
                        {
                            targetIndex = pres.Slides.Count;
                            break;
                        }
                    } catch { }
                }
                
                // Move slides to end
                foreach (Slide slide in selection.SlideRange)
                {
                    slide.MoveTo(targetIndex);
                }
                
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MoveToBackupFeature");
                return false;
            }
        }
    }
}