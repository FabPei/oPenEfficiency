using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnHideMasterObjects",
        Name = "Hide / Show Master Objects",
        Tooltip = "Toggle Master Objects",
        IconData = "M12,4.5C7,4.5 2.73,7.61 1,12C2.73,16.39 7,19.5 12,19.5C17,19.5 21.27,16.39 23,12C21.27,7.61 17,4.5 12,4.5M12,17C9.24,17 7,14.76 7,12C7,9.24 9.24,7 12,7C14.76,7 17,9.24 17,12C17,14.76 14.76,17 12,17M12,9A3,3 0 0,0 9,12A3,3 0 0,0 12,15A3,3 0 0,0 15,12A3,3 0 0,0 12,9Z", // Eye icon
        Color = "#8B5CF6", // Utilities color
        MinSelection = 0,
        RequiredType = PowerPoint.PpSelectionType.ppSelectionNone,
        Description = "Toggles the visibility of Slide Master shapes and backgrounds on the currently selected slides.",
        DetailedHelpText = "### Hide / Show Master Objects\n\nToggles the visibility of Slide Master background elements on the current slide, allowing full canvas access for precise positioning."
    )]
    public static class HideMasterObjectsFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var selection = manager.GetApplication().ActiveWindow.Selection;
                PowerPoint.SlideRange slideRange = null;

                try
                {
                    slideRange = selection.SlideRange;
                }
                catch
                {
                    // If SlideRange is not available (e.g. some views), try to get the active slide
                    var slide = manager.GetApplication().ActiveWindow.View.Slide as PowerPoint.Slide;
                    if (slide != null)
                    {
                        ToggleSlideMaster(slide);
                        return true;
                    }
                    return false;
                }

                if (slideRange != null && slideRange.Count > 0)
                {
                    foreach (PowerPoint.Slide slide in slideRange)
                    {
                        ToggleSlideMaster(slide);
                    }
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex.Message, "HideMasterObjectsFeature");
                return false;
            }
        }

        private static void ToggleSlideMaster(PowerPoint.Slide slide)
        {
            try
            {
                if (slide.DisplayMasterShapes == MsoTriState.msoTrue)
                {
                    // Currently visible -> Hide them
                    slide.DisplayMasterShapes = MsoTriState.msoFalse;
                    slide.FollowMasterBackground = MsoTriState.msoFalse;
                }
                else
                {
                    // Currently hidden -> Show them
                    slide.DisplayMasterShapes = MsoTriState.msoTrue;
                    slide.FollowMasterBackground = MsoTriState.msoTrue;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log("Failed to toggle master objects: " + ex.Message, "HideMasterObjectsFeature");
            }
        }
    }
}
