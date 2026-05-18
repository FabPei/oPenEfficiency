using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Shapes
{
    [FeatureMetadata(
        Id = "BtnCreateMotionPath",
        Name = "Create Motion Path",
        Tooltip = "Create Motion Path (From Shape A to Shape B)",
        IconData = "M14,2L20,8L14,14V11H10V7H14V2M11,19A6,6 0 0,0 5,13A6,6 0 0,0 11,7V9A4,4 0 0,1 7,13A4,4 0 0,1 11,17H19V19H11Z",
        Color = "#FBBF24",
        MinSelection = 2,
        RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes,
        Description = "Generates an automatic animation. The first selected shape will move precisely to the X/Y coordinates of the second selected shape.",
        DetailedHelpText = "### Create Motion Path\n\nSelect exactly two shapes. This feature injects a PowerPoint animation path so the first shape automatically travels to the exact coordinates of the second shape when the slide is presented."
    )]
    public static class CreateMotionPathFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var selection = manager.GetApplication().ActiveWindow.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count < 2) return false;

                // According to EE Spec, it's typically "First selected moves to Last Selected coordinates"
                // But ShapeRange doesn't cleanly preserve selection order.
                // We'll use ShapeRange[1] as From and ShapeRange[2] as To.
                PowerPoint.Shape shapeFrom = selection.ShapeRange[1];
                PowerPoint.Shape shapeTo = selection.ShapeRange[2];
                
                PowerPoint.Slide slide = shapeFrom.Parent as PowerPoint.Slide;
                if (slide == null) return false;
                
                var pres = manager.GetApplication().ActivePresentation;
                float slideWidth = pres.PageSetup.SlideWidth;
                float slideHeight = pres.PageSetup.SlideHeight;

                // Calculate center-to-center delta
                float fromCenterX = shapeFrom.Left + (shapeFrom.Width / 2);
                float fromCenterY = shapeFrom.Top + (shapeFrom.Height / 2);
                float toCenterX = shapeTo.Left + (shapeTo.Width / 2);
                float toCenterY = shapeTo.Top + (shapeTo.Height / 2);

                float dx = toCenterX - fromCenterX;
                float dy = toCenterY - fromCenterY;

                // Convert to percentage of slide dimensions relative to starting position
                float pctX = dx / slideWidth;
                float pctY = dy / slideHeight;

                string path = $"M 0 0 L {pctX.ToString(System.Globalization.CultureInfo.InvariantCulture)} {pctY.ToString(System.Globalization.CultureInfo.InvariantCulture)}";

                PowerPoint.Effect effect;
                try
                {
                    effect = slide.TimeLine.MainSequence.AddEffect(shapeFrom, PowerPoint.MsoAnimEffect.msoAnimEffectCustom, 
                                                                   PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                }
                catch
                {
                    // Fallback to path down if custom fails
                    effect = slide.TimeLine.MainSequence.AddEffect(shapeFrom, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, 
                                                                   PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                }

                PowerPoint.AnimationBehavior behavior = effect.Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeMotion);
                behavior.MotionEffect.Path = path;

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex.Message, "CreateMotionPathFeature");
                return false;
            }
        }
    }
}
