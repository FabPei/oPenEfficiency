using System;
using System.Windows.Interop;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnSlidePaste",
        Name = "Insert Slide Snapshots",
        SearchAliases = new[] { "Slide Paste" },
        Tooltip = "Paste Slides as Images",
        IconData = "M3,3H9V9H3V3M11,3H17V9H11V3M3,11H9V17H3V11M11,11H17V17H11V11Z",
        Color = "#8B5CF6",
        Description = "Selects slides from the current deck and inserts stable image references on the active slide in an automatically sized grid.",
        DetailedHelpText = "### Insert Slide Snapshots\nOpens thumbnails for the presentation that was active when the window loaded. Right-click presets the hyperlink and shadow combination. Select source slides, then confirm Insert Selected to place PNG snapshots on the active slide. The command aborts if you switch presentations, resolves reordered source slides by Slide ID, groups the insertion into one PowerPoint Undo entry, tags each image with its source Slide ID, and reports skipped exports. Temporary thumbnail and export files are removed when the window closes.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class SlidePasteFeature
    {
        private static SlidePasteWindow _window;

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, false, false);
        }

        public static bool Execute(PowerPointManager manager, bool addLinks, bool addShadows)
        {
            if (manager == null) return false;
            try
            {
                if (_window != null && _window.IsLoaded)
                {
                    _window.SetOptions(addLinks, addShadows);
                    _window.Activate();
                    return true;
                }

                _window = new SlidePasteWindow(manager, addLinks, addShadows);
                var app = manager.GetApplication();
                if (app != null) new WindowInteropHelper(_window).Owner = (IntPtr)app.HWND;
                _window.Closed += (s, e) => _window = null;
                _window.Show();
                return true;
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "SlidePasteFeature.Execute");
                var failedWindow = _window;
                _window = null;
                try { failedWindow?.Close(); }
                catch (Exception closeEx)
                {
                    oPenEfficiency.Utils.ExceptionLogger.Log(
                        closeEx,
                        "SlidePasteFeature.Execute.FailureCleanup");
                }
                throw;
            }
        }
    }
}
