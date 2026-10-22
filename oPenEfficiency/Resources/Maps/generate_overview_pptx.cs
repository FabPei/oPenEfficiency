using System;
using System.IO;
using Microsoft.Office.Core;
using oPenEfficiency;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

internal static class GenerateOverviewPptx
{
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length != 3)
        {
            Console.Error.WriteLine("Usage: GenerateOverviewPptx <blank-input.pptx> <output.pptx> <Mercator|Natural>");
            return 2;
        }

        string inputPath = Path.GetFullPath(args[0]);
        string outputPath = Path.GetFullPath(args[1]);
        MapProjection projection;
        if (!Enum.TryParse(args[2], true, out projection))
        {
            Console.Error.WriteLine("Projection must be Mercator or Natural.");
            return 2;
        }
        PowerPoint.Application app = null;
        PowerPoint.Presentation presentation = null;
        try
        {
            app = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            presentation = app.Presentations.Open(inputPath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
            if (presentation.Slides.Count != 8)
            {
                throw new InvalidOperationException("The blank overview asset must contain exactly eight slides.");
            }

            var manager = new PowerPointManager(app);
            presentation.Slides[1].Select();
            PowerPoint.Shape worldMap = MapAssetService.InsertWorldCountriesFromGeometry(manager, projection);
            RequireMap(worldMap, "WORLD-ADM0");
            Console.WriteLine("Generated WORLD-ADM0");

            string[] continents = { "Africa", "Asia", "Europe", "North America", "South America", "Oceania", "Antarctica" };
            for (int index = 0; index < continents.Length; index++)
            {
                presentation.Slides[index + 2].Select();
                MapProjection continentProjection = string.Equals(continents[index], "Antarctica", StringComparison.OrdinalIgnoreCase)
                    ? MapProjection.Natural
                    : projection;
                PowerPoint.Shape map = MapAssetService.InsertContinentCountriesFromGeometry(manager, continents[index], continentProjection);
                RequireMap(map, "CONTINENT-" + continents[index].ToUpperInvariant().Replace(' ', '-'));
                Console.WriteLine("Generated " + map.Tags[MapAssetService.MapIdTag]);
            }

            presentation.SaveAs(outputPath);
            Console.WriteLine(outputPath);
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
        finally
        {
            if (presentation != null) try { presentation.Close(); } catch { /* Generator cleanup continues if PowerPoint already closed the presentation. */ }
            if (app != null) try { app.Quit(); } catch { /* Generator cleanup continues if the temporary PowerPoint process already exited. */ }
        }
    }

    private static void RequireMap(PowerPoint.Shape map, string expectedMapId)
    {
        if (map == null || !string.Equals(map.Tags[MapAssetService.MapIdTag], expectedMapId, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Failed to generate " + expectedMapId + ".");
        }
    }
}
