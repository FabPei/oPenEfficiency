using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnHideMasterObjects",
        Name = "Hide/Show Master Objects",
        SearchAliases = new[] { "Hide / Show Master Objects" },
        Tooltip = "Hide or restore master objects",
        IconData = "M12,4.5C7,4.5 2.73,7.61 1,12C2.73,16.39 7,19.5 12,19.5C17,19.5 21.27,16.39 23,12C21.27,7.61 17,4.5 12,4.5M12,17C9.24,17 7,14.76 7,12C7,9.24 9.24,7 12,7C14.76,7 17,9.24 17,12C17,14.76 14.76,17 12,17M12,9A3,3 0 0,0 9,12A3,3 0 0,0 12,15A3,3 0 0,0 15,12A3,3 0 0,0 12,9Z",
        Color = "#8B5CF6",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None,
        Description = "Hides or restores master shapes and the master background on the selected slides or current slide.",
        DetailedHelpText = "### Hide/Show Master Objects\nHides master shapes and the master background on the selected slides, or on the current slide when no slide range is selected. A mixed selection is normalized by hiding all selected slides. The next click restores each slide's original master-shape and master-background settings. Right-click to hide, show, or toggle explicitly. Slide-level content is not changed.")]
    public static class HideMasterObjectsFeature
    {
        public enum MasterVisibilityMode { Toggle, Hide, Show }

        private const string HiddenTag = "oPE_MasterObjectsHidden";
        private const string OriginalShapesTag = "oPE_MasterShapesOriginal";
        private const string OriginalBackgroundTag = "oPE_MasterBackgroundOriginal";

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, MasterVisibilityMode.Toggle);
        }

        public static bool Execute(PowerPointManager manager, MasterVisibilityMode mode)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app?.ActiveWindow == null || app.ActivePresentation == null) return false;
                var slides = GetTargetSlides(app);
                if (slides.Count == 0) return false;
                bool hide = mode == MasterVisibilityMode.Hide ||
                    (mode == MasterVisibilityMode.Toggle && slides.Any(slide => !HasTag(slide, HiddenTag)));
                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "HideMasterObjectsFeature.StartUndo"); }
                foreach (var slide in slides)
                {
                    if (mode == MasterVisibilityMode.Hide || (mode == MasterVisibilityMode.Toggle && hide))
                        Hide(slide);
                    else if (mode == MasterVisibilityMode.Show)
                        Show(slide);
                    else
                        Restore(slide);
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "HideMasterObjectsFeature.Execute");
                return false;
            }
        }

        private static List<PowerPoint.Slide> GetTargetSlides(PowerPoint.Application app)
        {
            try
            {
                var range = app.ActiveWindow.Selection?.SlideRange;
                if (range != null && range.Count > 0) return range.Cast<PowerPoint.Slide>().ToList();
            }
            catch { /* Unsupported undo grouping does not block toggling master objects. */ }
            var active = app.ActiveWindow.View.Slide as PowerPoint.Slide;
            return active == null ? new List<PowerPoint.Slide>() : new List<PowerPoint.Slide> { active };
        }

        private static void Hide(PowerPoint.Slide slide)
        {
            if (!HasTag(slide, HiddenTag))
            {
                slide.Tags.Add(OriginalShapesTag, ((int)slide.DisplayMasterShapes).ToString());
                slide.Tags.Add(OriginalBackgroundTag, ((int)slide.FollowMasterBackground).ToString());
                slide.Tags.Add(HiddenTag, "true");
            }
            slide.DisplayMasterShapes = MsoTriState.msoFalse;
            slide.FollowMasterBackground = MsoTriState.msoFalse;
        }

        private static void Restore(PowerPoint.Slide slide)
        {
            slide.DisplayMasterShapes = ParseTriState(slide.Tags[OriginalShapesTag], MsoTriState.msoTrue);
            slide.FollowMasterBackground = ParseTriState(slide.Tags[OriginalBackgroundTag], MsoTriState.msoTrue);
            TryDeleteTag(slide, HiddenTag);
            TryDeleteTag(slide, OriginalShapesTag);
            TryDeleteTag(slide, OriginalBackgroundTag);
        }

        private static void Show(PowerPoint.Slide slide)
        {
            slide.DisplayMasterShapes = MsoTriState.msoTrue;
            slide.FollowMasterBackground = MsoTriState.msoTrue;
            TryDeleteTag(slide, HiddenTag);
            TryDeleteTag(slide, OriginalShapesTag);
            TryDeleteTag(slide, OriginalBackgroundTag);
        }

        private static MsoTriState ParseTriState(string value, MsoTriState fallback)
        {
            return int.TryParse(value, out int parsed) ? (MsoTriState)parsed : fallback;
        }

        private static bool HasTag(PowerPoint.Slide slide, string name)
        {
            try { return string.Equals(slide.Tags[name], "true", StringComparison.OrdinalIgnoreCase); }
            catch { return false; }
        }

        private static void TryDeleteTag(PowerPoint.Slide slide, string name)
        {
            try { slide.Tags.Delete(name); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "HideMasterObjectsFeature.TryDeleteTag"); }
        }
    }
}
