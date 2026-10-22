using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Set or toggle the locking state of shapes.
    /// Uses native PowerPoint shape locking when it is available.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnShapeLock",
        Name = "Lock/Unlock Shapes",
        SearchAliases = new[] { "Shape Locking" },
        Tooltip = "Lock or unlock the selected shapes",
        IconData = "M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M19,19H5V5H19V19M12,7H14V12H12V7M12,14H14V17H12V14Z",
        Color = "#6B7280",
        Description = "Locks or unlocks the selected shapes to prevent accidental edits.",
        DetailedHelpText = "### Lock/Unlock Shapes\nLocks or unlocks the selected shapes using PowerPoint's native shape lock. If the selection contains both states, left-click locks every selected shape; left-click again unlocks them.\n\n**Right-click options:**\n* **Toggle Lock All Shapes**: Lock all shapes when any shape is unlocked; unlock all when every shape is locked.\n* **Lock All Shapes**: Lock all shapes on every slide in the presentation.\n* **Unlock All Shapes**: Unlock all shapes on every slide in the presentation.\n\nNative shape locking requires a PowerPoint version that supports it.",
        MinSelection = 0,
        IsToggle = true)]
    public static class ShapeLockingFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - toggles shape lock state for the selection.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return SetSelectionShapeLock(manager, forceState: null);
        }

        public static bool SetGlobalShapeLock(PowerPointManager manager, bool? forceState = null)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActivePresentation == null) return false;

                bool found = false;
                bool allLocked = true;
                foreach (Slide slide in app.ActivePresentation.Slides)
                {
                    foreach (Shape shape in slide.Shapes)
                    {
                        found = true;
                        if (!IsShapeLocked(shape)) allLocked = false;
                    }
                }
                if (!found) return false;

                bool targetState = forceState ?? !allLocked;
                bool allSucceeded = true;

                foreach (Slide slide in app.ActivePresentation.Slides)
                {
                    foreach (Shape shape in slide.Shapes)
                    {
                        if (!SetShapeLock(shape, targetState)) allSucceeded = false;
                    }
                }
                return allSucceeded;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ShapeLockingFeature.SetGlobalShapeLock");
                return false;
            }
        }

        public static bool SetSelectionShapeLock(PowerPointManager manager, bool? forceState = null)
        {
            if (manager == null) return false;
            try
            {
                var selection = manager.GetSelectedShapes();
                if (selection == null || selection.Count == 0) return false;

                bool allLocked = true;
                for (int i = 1; i <= selection.Count; i++)
                {
                    if (!IsShapeLocked(selection[i])) allLocked = false;
                }
                bool targetState = forceState ?? !allLocked;
                bool allSucceeded = true;

                for (int i = 1; i <= selection.Count; i++)
                {
                    if (!SetShapeLock(selection[i], targetState)) allSucceeded = false;
                }
                return allSucceeded;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ShapeLockingFeature.SetSelectionShapeLock");
                return false;
            }
        }

        public static bool IsShapeLocked(Shape shape)
        {
            return TryGetNativeLockState(shape, out bool locked) && locked;
        }

        private static bool SetShapeLock(Shape shape, bool lockState)
        {
            if (shape == null) return false;
            try
            {
                ((dynamic)shape).Locked = lockState
                    ? Office.MsoTriState.msoTrue
                    : Office.MsoTriState.msoFalse;
                bool succeeded = TryGetNativeLockState(shape, out bool currentState) && currentState == lockState;
                if (succeeded)
                {
                    try { shape.Tags.Add("oPE_Locked", lockState ? "true" : "false"); }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "ShapeLockingFeature.SetShapeLock.Tag"); }
                }
                return succeeded;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ShapeLockingFeature.SetShapeLock");
                return false;
            }
        }

        private static bool TryGetNativeLockState(Shape shape, out bool locked)
        {
            locked = false;
            if (shape == null) return false;
            try
            {
                locked = (int)((dynamic)shape).Locked != 0;
                return true;
            }
            catch
            {
                // Unsupported PowerPoint versions are an expected compatibility case. The
                // mutating path logs failures, while this frequently-polled query stays quiet.
                return false;
            }
        }
    }
}
