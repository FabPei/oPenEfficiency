using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Swap vertical positions of two shapes.
    /// </summary>
    public static class SwapPositionsVerticalFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count != 2) return false;
            try
            {
                SwapPositionsFeature.SwapAxis(shapes[1], shapes[2], horizontal: false);
                return true;
            }
            catch (System.Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "SwapPositionsVerticalFeature.Execute");
                return false;
            }
        }
    }
}
