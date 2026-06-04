using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnHideSelected",
        Name = "Hide Selected",
        Tooltip = "Hide selected",
        IconData = "M12,4.5C7,4.5 2.73,7.61 1,12C2.73,16.39 7,19.5 12,19.5C13.2,19.5 14.3,19.2 15.3,18.7L4.3,7.7C3.3,8.8 2.5,10.3 2,12C2.73,16.39 7,19.5 12,19.5C17,19.5 21.27,16.39 23,12C21.27,7.61 17,4.5 12,4.5M12,17.5C14.33,17.5 16.46,16.41 17.84,14.65L12,8.81L9,11.81V13H10.5V14.09L13.09,16.68C12.7,17.14 12.2,17.5 12,17.5M17.84,14.65L19.49,4.72L18.04,6.17C16.4,5.1 14.28,4.5 12,4.5C13.23,8 14.33,8.39 15.22,9.05L13.72,10.55C13.2,10.2 12.63,10 12,10C10.9,10 10,10.9 10,12C10,12.63 10.2,13.2 10.55,13.72L8.5,15.77C8.18,14.64 8,13.34 8,12A4,4 0 0,1 12,8Z",
        Color = "#D946EF",
        Description = "Hides all currently selected shapes on the slide. Use 'Show Hidden' to make them visible again.",
        DetailedHelpText = "### Hide Selected\nTemporarily hides selected shapes without deleting them. The shapes remain recoverable via the Show Hidden feature.",
        Keywords = "invisible, conceal, disappear, stealth, temporarily, remove",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class HideSelectedFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return VisibilityFeature.HideSelectedShapes(manager);
        }
    }
}
