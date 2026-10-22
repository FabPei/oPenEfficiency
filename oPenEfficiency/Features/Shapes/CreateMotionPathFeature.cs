using System;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using MotionPathMode = oPenEfficiency.Core.Services.MotionPathPolicy.MotionPathMode;
using MotionPathPolicy = oPenEfficiency.Core.Services.MotionPathPolicy;
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
        MaxSelection = 2,
        RequiredType = OfficeSelectionType.Shapes,
        Description = "Adds an on-click motion path from the first selected shape's center to the second selected shape's center.",
        DetailedHelpText = "### Create Motion Path\n\nSelect exactly two shapes. The primary action creates an on-click motion path from the first selected shape's center to the second. Right-click can reverse the direction or create both directions in one Undo step. Existing animations remain unchanged, the shapes do not move while editing, and incomplete effects are removed if creation fails."
    )]
    public static class CreateMotionPathFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, MotionPathMode.FirstToSecond);
        }

        public static bool Execute(PowerPointManager manager, MotionPathMode mode)
        {
            PowerPoint.Sequence sequence = null;
            int initialEffectCount = 0;
            try
            {
                var app = manager?.GetApplication();
                var selection = app?.ActiveWindow?.Selection;
                if (selection?.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count != 2) return false;

                PowerPoint.Shape first = selection.ShapeRange[1];
                PowerPoint.Shape second = selection.ShapeRange[2];
                
                PowerPoint.Slide slide = first.Parent as PowerPoint.Slide;
                PowerPoint.Slide targetSlide = second.Parent as PowerPoint.Slide;
                if (slide == null || targetSlide == null || targetSlide.SlideID != slide.SlideID) return false;
                
                var pres = app.ActivePresentation;
                float slideWidth = pres.PageSetup.SlideWidth;
                float slideHeight = pres.PageSetup.SlideHeight;
                if (slideWidth <= 0 || slideHeight <= 0) return false;

                sequence = slide.TimeLine.MainSequence;
                initialEffectCount = sequence.Count;
                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "CreateMotionPathFeature.StartNewUndoEntry"); }

                if (MotionPathPolicy.MovesFirstToSecond(mode))
                    AddMotionPath(sequence, first, second, slideWidth, slideHeight);
                if (MotionPathPolicy.MovesSecondToFirst(mode))
                    AddMotionPath(sequence, second, first, slideWidth, slideHeight);

                return true;
            }
            catch (Exception ex)
            {
                if (sequence != null)
                {
                    try
                    {
                        while (sequence.Count > initialEffectCount)
                            sequence[sequence.Count].Delete();
                    }
                    catch (Exception cleanupEx) { ExceptionLogger.Log(cleanupEx, "CreateMotionPathFeature.Cleanup"); }
                }
                ExceptionLogger.Log(ex, "CreateMotionPathFeature.Execute");
                return false;
            }
        }

        private static void AddMotionPath(
            PowerPoint.Sequence sequence,
            PowerPoint.Shape shapeFrom,
            PowerPoint.Shape shapeTo,
            float slideWidth,
            float slideHeight)
        {
            // Read both shapes' bounds up front, then delegate the slide-relative delta and
            // the invariant-culture path string to the host-free policy.
            MotionPathPolicy.ComputeNormalizedDelta(
                shapeFrom.Left, shapeFrom.Top, shapeFrom.Width, shapeFrom.Height,
                shapeTo.Left, shapeTo.Top, shapeTo.Width, shapeTo.Height,
                slideWidth, slideHeight,
                out float pctX, out float pctY);
            string path = MotionPathPolicy.BuildPathString(pctX, pctY);

            var effect = sequence.AddEffect(
                    shapeFrom,
                    PowerPoint.MsoAnimEffect.msoAnimEffectPathDown,
                    PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);

                PowerPoint.AnimationBehavior behavior = null;
                for (int i = 1; i <= effect.Behaviors.Count; i++)
                {
                    if (effect.Behaviors[i].Type == PowerPoint.MsoAnimType.msoAnimTypeMotion)
                    {
                        behavior = effect.Behaviors[i];
                        break;
                    }
                }
                if (behavior == null) behavior = effect.Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeMotion);
                behavior.MotionEffect.Path = path;
        }
    }
}
