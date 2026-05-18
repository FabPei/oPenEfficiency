using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Alignment
{
    [FeatureMetadata(Id = "BtnDistributeHorizontal", Name = "Distribute Horizontally", Tooltip = "Distribute Horizontally (Right-click for options)", IconData = "M4 22H2V2h2v20zM22 2v20h-2V2h2zm-4 4H6v12h12V6z", Color = "#10B981", Description = "Distributes selected shapes evenly horizontally. Right-click to toggle between 'Align to Slide' or 'Align Selected Objects'.", MinSelection = 2, RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class DistributeHorizontalFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, false, false);
        }

        public static bool Execute(PowerPointManager manager, bool relativeToSlide, bool alignMiddles = false)
        {
            var app = manager?.GetApplication();
            if (app?.ActiveWindow?.Selection?.Type != PpSelectionType.ppSelectionShapes) return false;
            try
            {
                var sr = app.ActiveWindow.Selection.ShapeRange;
                if (sr == null || (sr.Count < 3 && !relativeToSlide)) return false;

                var rel = relativeToSlide ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                sr.Distribute(Office.MsoDistributeCmd.msoDistributeHorizontally, rel);

                if (alignMiddles)
                {
                    sr.Align(Office.MsoAlignCmd.msoAlignMiddles, Office.MsoTriState.msoFalse);
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DistributeHorizontalFeature.Execute");
                return false;
            }
        }
    }

    [FeatureMetadata(Id = "BtnDistributeVertical", Name = "Distribute Vertically", Tooltip = "Distribute Vertically (Right-click for options)", IconData = "M22 2v2H2V2h20zm0 18v2H2v-2h20zm-4-4H6V8h12v8z", Color = "#10B981", Description = "Distributes selected shapes evenly vertically. Right-click to toggle between 'Align to Slide' or 'Align Selected Objects'.", MinSelection = 2, RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class DistributeVerticalFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, false, false);
        }

        public static bool Execute(PowerPointManager manager, bool relativeToSlide, bool alignCenters = false)
        {
            var app = manager?.GetApplication();
            if (app?.ActiveWindow?.Selection?.Type != PpSelectionType.ppSelectionShapes) return false;
            try
            {
                var sr = app.ActiveWindow.Selection.ShapeRange;
                if (sr == null || (sr.Count < 3 && !relativeToSlide)) return false;

                var rel = relativeToSlide ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                sr.Distribute(Office.MsoDistributeCmd.msoDistributeVertically, rel);

                if (alignCenters)
                {
                    sr.Align(Office.MsoAlignCmd.msoAlignCenters, Office.MsoTriState.msoFalse);
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DistributeVerticalFeature.Execute");
                return false;
            }
        }
    }
}
