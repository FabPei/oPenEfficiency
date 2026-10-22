using System;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Visuals
{
    [FeatureMetadata(
        Id = "BtnIconBrowser",
        Name = "Icon Library",
        SearchAliases = new[] { "Icon Browser" },
        Tooltip = "Browse and insert SVG icons",
        IconData = "M4,5A3,3 0 0,1 7,2H17A3,3 0 0,1 20,5V15A3,3 0 0,1 17,18H14L9,22V18H7A3,3 0 0,1 4,15V5M8,6V14H16V6H8M10,8H14V12H10V8Z",
        Color = "#0EA5E9",
        Description = "Browse licensed icon packs, filter by style and theme, recolor icons, and insert them as SVG.",
        DetailedHelpText = "Browse licensed icon packs, filter by style and theme, recolor icons, and insert them as SVG images or editable PowerPoint shapes into the active slide. Right-click opens focused search, Favorites, or a predefined insert mode.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class IconBrowserFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                Globals.ThisAddIn.OpenIconBrowser();
                return true;
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IconBrowserFeature.Execute");
                throw;
            }
        }
    }
}
