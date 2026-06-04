using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnSlidePaste",
        Name = "Slide Paste",
        Tooltip = "Paste Slides as Images",
        IconData = "M3,3H9V9H3V3M11,3H17V9H11V3M3,11H9V17H3V11M11,11H17V17H11V11Z",
        Color = "#8B5CF6",
        Description = "Opens a panel to select slides from the current deck and paste them as images on the active slide in an auto-sized grid layout. Supports optional hyperlinks and shadow backdrop.",
        DetailedHelpText = "### Slide Paste\nPastes the current clipboard content simultaneously onto all slides currently selected in the Slide Sorter view.",
        Keywords = "slide paste, image grid, link, thumbnails, overview, reference",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class SlidePasteFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            var dialog = new SlidePasteWindow(manager);
            dialog.Show();
            return true;
        }
    }
}
