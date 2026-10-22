using System;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using DiagonalPlacementPolicy = oPenEfficiency.Core.Services.DiagonalPlacementPolicy;

namespace oPenEfficiency.Features.Alignment
{
    [FeatureMetadata(Id = "BtnDiagonalAlign", Name = "Align Diagonally", SearchAliases = new[] { "Diagonal Align" }, Tooltip = "Align selection diagonally at a specific degree", IconData = "M21 3L3 21M21 3l-6 0M21 3l0 6", Color = "#8B5CF6", Description = "Aligns selected shapes in a diagonal line (e.g., 45 degrees). Use the right-click menu to choose specific angles.", DetailedHelpText = "### Align Diagonally\n\nPlaces the selected shapes in an evenly spaced line at the chosen angle. The toolbar lets you use common angles and choose whether the first shape, last shape, or the selection center remains anchored.", MinSelection = 2, RequiredType = OfficeSelectionType.Shapes)]
    public static class DiagonalAlignFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var window = new UI.Toolbars.DiagonalAlignToolbar(manager);
                window.Show();
                return true;
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "DiagonalAlignFeature.Execute");
                throw;
            }
        }

        public static bool ExecuteSpecific(PowerPointManager manager, float angleDeg, string anchorType = "First")
        {
            try
            {
                if (manager == null || float.IsNaN(angleDeg) || float.IsInfinity(angleDeg)) return false;

                var sr = manager.GetSelectedShapes();
                if (sr == null || sr.Count < 2) return false;

                // Read every shape's original bounds up front, then delegate the placement
                // (projection, ordering, spacing and anchor arithmetic) to the host-free policy.
                var centerX = new float[sr.Count];
                var centerY = new float[sr.Count];
                var width = new float[sr.Count];
                var height = new float[sr.Count];
                for (int i = 1; i <= sr.Count; i++)
                {
                    var s = sr[i];
                    centerX[i - 1] = s.Left + (s.Width / 2f);
                    centerY[i - 1] = s.Top + (s.Height / 2f);
                    width[i - 1] = s.Width;
                    height[i - 1] = s.Height;
                }

                DiagonalPlacementPolicy.ComputePlacements(
                    centerX, centerY, width, height,
                    angleDeg, DiagonalPlacementPolicy.ResolveAnchor(anchorType),
                    out float[] newLeft, out float[] newTop);

                bool allShapesMoved = true;
                for (int i = 0; i < sr.Count; i++)
                {
                    var shp = sr[i + 1];
                    try
                    {
                        shp.Left = newLeft[i];
                        shp.Top = newTop[i];
                    }
                    catch (Exception ex)
                    {
                        Utils.ExceptionLogger.Log(ex, "DiagonalAlignFeature.ExecuteSpecific.MoveShape");
                        allShapesMoved = false;
                    }
                }

                return allShapesMoved;
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "DiagonalAlignFeature.ExecuteSpecific");
                return false;
            }
        }
    }
}
