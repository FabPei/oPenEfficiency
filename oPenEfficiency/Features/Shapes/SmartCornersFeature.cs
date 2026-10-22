using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnSmartCorners",
        Name = "Toggle Rounded Corners",
        SearchAliases = new[] { "Smart Corners" },
        Tooltip = "Switch selected rectangles between rounded and square corners",
        IconData = "M3,3H21V21H3V3M7,7V17H17V7H7Z",
        Color = "#F43F5E",
        Description = "Synchronizes selected native rectangles between square and uniformly rounded corners.",
        DetailedHelpText = "### Toggle Rounded Corners\n\nSelect one or more native Rectangle or Rounded Rectangle shapes. The primary action toggles all compatible shapes between square and uniformly rounded. Right-click to force Rounded, force Square, or Toggle. Incompatible shapes are skipped. Shape objects, text, formatting, tags, links, animations, and Z-order are preserved because the native AutoShape type changes in place.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class SmartCornersFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            var shapes = GetCompatibleShapes(manager);
            if (shapes.Count == 0)
            {
                ShowInfo("Select at least one native Rectangle or Rounded Rectangle.");
                return false;
            }

            bool makeRounded = false;
            foreach (var shape in shapes)
            {
                if (shape.AutoShapeType != MsoAutoShapeType.msoShapeRoundedRectangle)
                {
                    makeRounded = true;
                    break;
                }
            }
            return SetRounded(manager, shapes, makeRounded, "SmartCornersFeature.Execute");
        }

        public static void ToggleCorner(PowerPointManager manager, int cornerIndex)
        {
            if (cornerIndex == -1)
            {
                Execute(manager);
                return;
            }

            ShowInfo("Individual corner toggles would require replacing the PowerPoint shape and are intentionally not supported. Use Toggle Rounded Corners to change all four native corners safely.");
        }

        public static void ResetToSquare(PowerPointManager manager)
        {
            var shapes = GetCompatibleShapes(manager);
            if (shapes.Count == 0)
            {
                ShowInfo("Select at least one native Rectangle or Rounded Rectangle.");
                return;
            }
            SetRounded(manager, shapes, false, "SmartCornersFeature.ResetToSquare");
        }

        public static bool MakeRounded(PowerPointManager manager)
        {
            var shapes = GetCompatibleShapes(manager);
            if (shapes.Count == 0) return false;
            return SetRounded(manager, shapes, true, "SmartCornersFeature.MakeRounded");
        }

        public static bool MakeSquare(PowerPointManager manager)
        {
            var shapes = GetCompatibleShapes(manager);
            if (shapes.Count == 0) return false;
            return SetRounded(manager, shapes, false, "SmartCornersFeature.MakeSquare");
        }

        public static bool[] GetCommonRoundedState(PowerPointManager manager)
        {
            var shapes = GetCompatibleShapes(manager);
            if (shapes.Count == 0) return new bool[4];
            foreach (var shape in shapes)
            {
                if (shape.AutoShapeType != MsoAutoShapeType.msoShapeRoundedRectangle)
                    return new bool[4];
            }
            return new[] { true, true, true, true };
        }

        private static List<PowerPoint.Shape> GetCompatibleShapes(PowerPointManager manager)
        {
            var result = new List<PowerPoint.Shape>();
            var selection = manager?.GetSelectedShapes();
            if (selection == null) return result;

            for (int i = 1; i <= selection.Count; i++)
            {
                var shape = selection[i];
                if (shape.Type == MsoShapeType.msoAutoShape &&
                    (shape.AutoShapeType == MsoAutoShapeType.msoShapeRectangle ||
                     shape.AutoShapeType == MsoAutoShapeType.msoShapeRoundedRectangle))
                    result.Add(shape);
            }
            return result;
        }

        private static bool SetRounded(PowerPointManager manager, IList<PowerPoint.Shape> shapes, bool rounded, string context)
        {
            var originalTypes = new List<Tuple<PowerPoint.Shape, MsoAutoShapeType>>();
            try
            {
                foreach (var shape in shapes)
                    originalTypes.Add(Tuple.Create(shape, shape.AutoShapeType));

                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, context + ".StartNewUndoEntry"); }

                var targetType = rounded
                    ? MsoAutoShapeType.msoShapeRoundedRectangle
                    : MsoAutoShapeType.msoShapeRectangle;
                foreach (var shape in shapes) shape.AutoShapeType = targetType;
                return true;
            }
            catch (Exception ex)
            {
                foreach (var original in originalTypes)
                {
                    try { original.Item1.AutoShapeType = original.Item2; }
                    catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, context + ".Rollback"); }
                }
                ExceptionLogger.Log(ex, context);
                System.Windows.Forms.MessageBox.Show(
                    "PowerPoint could not update the selected rectangle corners.",
                    "Toggle Rounded Corners",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        private static void ShowInfo(string message)
        {
            System.Windows.Forms.MessageBox.Show(
                message,
                "Toggle Rounded Corners",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }
    }
}
