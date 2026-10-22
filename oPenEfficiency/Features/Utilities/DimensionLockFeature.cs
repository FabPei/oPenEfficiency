using System;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Lock or unlock the dimensions (width and height) of selected shapes.
    /// This allows shapes to be moved but prevents resizing, preserving their aspect ratio and size.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnLockDimensions",
        Name = "Lock/Unlock Aspect Ratio",
        SearchAliases = new[] { "Aspect Ratio Lock" },
        Tooltip = "Toggle proportional resizing for selected shapes",
        IconData = "M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M19,19H5V5H19V19M12,7H14V12H12V7M12,14H14V17H12V14Z",
        Color = "#6B7280",
        Description = "Locks or unlocks the aspect ratio of selected shapes while still allowing proportional resizing.",
        DetailedHelpText = "### Lock/Unlock Aspect Ratio\nLocks the width-to-height ratio of all selected shapes. Shapes can still be moved and resized proportionally. If the selection has mixed states, clicking locks all selected shapes; clicking again unlocks all. Right-click to lock or unlock every slide shape in the presentation.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes,
        IsToggle = true)]
    public static class DimensionLockFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - toggles dimension lock state.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return SetSelectionDimensionLock(manager, forceState: null);
        }

        public static bool SetSelectionDimensionLock(PowerPointManager manager, bool? forceState = null)
        {
            if (manager == null) return false;
            try
            {
                var selection = manager.GetSelectedShapes();
                if (selection == null || selection.Count == 0) return false;

                bool targetState = forceState ?? selection.Cast<Shape>().Any(shape => !IsDimensionLocked(shape));

                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "DimensionLockFeature.SetSelectionDimensionLock.StartUndo"); }

                for (int i = 1; i <= selection.Count; i++)
                {
                    SetDimensionLock(selection[i], targetState);
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DimensionLockFeature.SetSelectionDimensionLock");
                return false;
            }
        }

        public static bool SetGlobalDimensionLock(PowerPointManager manager, bool? forceState = null)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActivePresentation == null) return false;

                var shapes = app.ActivePresentation.Slides.Cast<Slide>()
                    .SelectMany(slide => slide.Shapes.Cast<Shape>()).ToList();
                if (shapes.Count == 0) return false;
                bool targetState = forceState ?? shapes.Any(shape => !IsDimensionLocked(shape));

                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "DimensionLockFeature.SetGlobalDimensionLock.StartUndo"); }

                foreach (var shape in shapes) SetDimensionLock(shape, targetState);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DimensionLockFeature.SetGlobalDimensionLock");
                return false;
            }
        }

        public static bool IsDimensionLocked(Shape shape)
        {
            try
            {
                return shape.Tags["oPE_DimensionLocked"] == "true" ||
                       shape.LockAspectRatio == Office.MsoTriState.msoTrue;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DimensionLockFeature.IsDimensionLocked");
                return false;
            }
        }

        private static void SetDimensionLock(Shape shape, bool lockState)
        {
            try
            {
                // Lock aspect ratio to prevent dimension changes
                shape.LockAspectRatio = lockState ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DimensionLockFeature.SetDimensionLock.LockAspectRatio");
            }

            try
            {
                if (lockState) shape.Tags.Add("oPE_DimensionLocked", "true");
                else shape.Tags.Delete("oPE_DimensionLocked");
            }
            catch (Exception)
            {
                try
                {
                    // If tag already exists, delete it and add again
                    shape.Tags.Delete("oPE_DimensionLocked");
                    if (lockState) shape.Tags.Add("oPE_DimensionLocked", "true");
                }
                catch (Exception ex2)
                {
                    ExceptionLogger.Log(ex2, "DimensionLockFeature.SetDimensionLock.TagUpdate");
                }
            }
        }
    }
}
