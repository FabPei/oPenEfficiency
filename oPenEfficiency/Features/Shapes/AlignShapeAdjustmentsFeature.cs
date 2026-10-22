using System;
using System.Collections.Generic;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using ShapeAdjustmentTransferPolicy = oPenEfficiency.Core.Services.ShapeAdjustmentTransferPolicy;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Shapes
{
    internal static class ShapeAdjustmentTransfer
    {
        public static bool Execute(PowerPointManager manager, string context)
        {
            return Execute(manager, context, null);
        }

        public static bool Execute(PowerPointManager manager, string context, bool? useLastSelected)
        {
            if (manager == null) return false;
            var snapshots = new List<Tuple<PowerPoint.Shape, float[]>>();
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2)
                {
                    ShowInfo("Select at least two shapes.");
                    return false;
                }

                int? explicitReferenceIndex = ShapeAdjustmentTransferPolicy.ResolveExplicitReferenceIndex(
                    shapeRange.Count, useLastSelected);
                var reference = explicitReferenceIndex.HasValue
                    ? shapeRange[explicitReferenceIndex.Value]
                    : manager.GetReferenceShape();
                if (reference == null || reference.Adjustments == null || reference.Adjustments.Count == 0)
                {
                    ShowInfo("The configured anchor shape has no adjustable geometry.");
                    return false;
                }

                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var target = shapeRange[i];
                    if (target.Id == reference.Id) continue;
                    try
                    {
                        bool sameShapeType = target.Type == reference.Type;
                        bool sameAutoShapeType = target.AutoShapeType == reference.AutoShapeType;
                        if (target.Adjustments == null ||
                            !ShapeAdjustmentTransferPolicy.IsTransferCompatible(
                                sameShapeType, sameAutoShapeType,
                                target.Adjustments.Count, reference.Adjustments.Count))
                            continue;

                        var values = new float[target.Adjustments.Count];
                        for (int adjustment = 1; adjustment <= target.Adjustments.Count; adjustment++)
                            values[adjustment - 1] = target.Adjustments[adjustment];
                        snapshots.Add(Tuple.Create(target, values));
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, context + ".InspectTarget");
                    }
                }

                if (snapshots.Count == 0)
                {
                    ShowInfo("No other selected shape has the same adjustable geometry as the configured anchor.");
                    return false;
                }

                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, context + ".StartNewUndoEntry"); }

                foreach (var snapshot in snapshots)
                {
                    for (int adjustment = 1; adjustment <= reference.Adjustments.Count; adjustment++)
                        snapshot.Item1.Adjustments[adjustment] = reference.Adjustments[adjustment];
                }
                return true;
            }
            catch (Exception ex)
            {
                foreach (var snapshot in snapshots)
                {
                    try
                    {
                        for (int adjustment = 1; adjustment <= snapshot.Item2.Length; adjustment++)
                            snapshot.Item1.Adjustments[adjustment] = snapshot.Item2[adjustment - 1];
                    }
                    catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, context + ".Rollback"); }
                }
                ExceptionLogger.Log(ex, context);
                System.Windows.Forms.MessageBox.Show(
                    "PowerPoint could not transfer the shape geometry. No partial geometry should remain.",
                    "Apply Anchor Shape Adjustments",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        private static void ShowInfo(string message)
        {
            System.Windows.Forms.MessageBox.Show(
                message,
                "Apply Anchor Shape Adjustments",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }
    }

    [FeatureMetadata(
        Id = "BtnAlignShapeAdjustments",
        Name = "Copy Shape Adjustments",
        SearchAliases = new[]
        {
            "Apply Anchor Shape Adjustments",
            "Match Shape Adjustments",
            "Match Shape Geometry",
            "Shape Transformation / Vertex Align",
            "Align Shape Proportions"
        },
        Tooltip = "Copy yellow-diamond adjustments from the anchor",
        IconData = "M12,2L4.5,20.29L5.21,21L12,18L18.79,21L19.5,20.29L12,2M12,5.83L17.5,19L12,16.62L6.5,19L12,5.83Z", // Compass/Transform icon
        Color = "#FBBF24",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes,
        Description = "Copies yellow-diamond adjustment values from the configured anchor to compatible selected shapes.",
        DetailedHelpText = "### Copy Shape Adjustments\n\nUses the first or last selected anchor configured in Settings. Right-click to use the first or last selected shape explicitly without changing Settings. Every other selected shape with the same PowerPoint shape type and the same number of yellow-diamond adjustments receives the anchor's exact adjustment values. Size, position, rotation, text, and formatting remain unchanged; incompatible shapes are skipped and write failures roll back all copied values."
    )]
    public static class AlignShapeAdjustmentsFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return ShapeAdjustmentTransfer.Execute(manager, "AlignShapeAdjustmentsFeature.Execute");
        }

        public static bool ExecuteUsingFirstSelected(PowerPointManager manager)
        {
            return ShapeAdjustmentTransfer.Execute(manager, "AlignShapeAdjustmentsFeature.ExecuteUsingFirstSelected", false);
        }

        public static bool ExecuteUsingLastSelected(PowerPointManager manager)
        {
            return ShapeAdjustmentTransfer.Execute(manager, "AlignShapeAdjustmentsFeature.ExecuteUsingLastSelected", true);
        }
    }
}
