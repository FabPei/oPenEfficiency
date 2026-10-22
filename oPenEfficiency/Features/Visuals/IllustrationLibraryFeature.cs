using System;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Visuals
{
    [FeatureMetadata(
        Id = "BtnIllustrationLibrary",
        Name = "Illustration Library",
        SearchAliases = new[] { "Illustrations", "Illustration Sources", "SVG Import", "Import Illustrations" },
        Tooltip = "Browse and insert bundled illustrations, or import your own SVG artwork",
        IconData = "M5,4H19A2,2 0 0,1 21,6V18A2,2 0 0,1 19,20H5A2,2 0 0,1 3,18V6A2,2 0 0,1 5,4M5,6V16L9,12L12,15L16,11L19,14V6H5M8,8A1.5,1.5 0 1,1 8,11A1.5,1.5 0 0,1 8,8Z",
        Color = "#8B5CF6",
        Description = "Pick a bundled or imported illustration, recolour it, switch its parts off or on, and insert it into the slide - or open the evaluated external sources and import your own SVG artwork.",
        DetailedHelpText = "The Illustration Library has two views. Figures: every bundled and imported illustration with its palette (each colour can be replaced) and its named parts (each part can be left out); insert it as a scalable SVG picture or as converted editable shapes, and the inserted shape carries its source in the shape tags. Sources & import: the bundled CC0 collections plus every evaluated external illustration source with its license, attribution duty and technical notes. Open a source in your browser, download the artwork you are licensed to use, and import the SVG files: the import screens each file for scripts, embedded images and external references, refuses files it cannot recolour, and records the source and license in the user library below %LOCALAPPDATA%\\oPenEfficiency\\Illustrations. Only the bundled CC0 collections may be redistributed with the add-in; licensing of imported artwork stays your own responsibility.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class IllustrationLibraryFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                Globals.ThisAddIn.OpenIllustrationLibrary(manager);
                return true;
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationLibraryFeature.Execute");
                throw;
            }
        }
    }
}
