using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Set or toggle the locking state of shapes.
    /// Handles native PowerPoint locking (Office 365/2021+) and fallback metadata-based locking.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnShapeLock",
        Name = "Shape Locking",
        Tooltip = "Shape locking",
        IconData = "M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M19,19H5V5H19V19M12,7H14V12H12V7M12,14H14V17H12V14Z",
        Color = "#6B7280",
        Description = "Locks or unlocks all shapes in the presentation to prevent accidental edits.",
        DetailedHelpText = "### Shape Locking\nLocks the position and size of the selected shapes to prevent accidental edits.\n**Right-Click Options:**\n* **Global Lock**: Lock all shapes on every slide in the presentation.\n* **Global Unlock**: Unlock everything in the presentation instantly for mass editing.",
        MinSelection = 0,
        IsToggle = true)]
    public static class ShapeLockingFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - toggles shape lock state for all shapes.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return SetGlobalShapeLock(manager, forceState: null);
        }

        public static bool SetGlobalShapeLock(PowerPointManager manager, bool? forceState = null)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActivePresentation == null) return false;

                bool targetState = true;
                if (forceState == null)
                {
                    bool found = false;
                    foreach (Slide slide in app.ActivePresentation.Slides)
                    {
                        if (slide.Shapes.Count > 0)
                        {
                            targetState = !IsShapeLocked(slide.Shapes[1]);
                            found = true;
                            break;
                        }
                    }
                    if (!found) return false;
                }
                else targetState = forceState.Value;

                foreach (Slide slide in app.ActivePresentation.Slides)
                {
                    foreach (Shape shape in slide.Shapes) SetShapeLock(shape, targetState);
                }
                return true;
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

                bool targetState = true;
                if (forceState == null)
                {
                    targetState = !IsShapeLocked(selection[1]);
                }
                else targetState = forceState.Value;

                for (int i = 1; i <= selection.Count; i++)
                {
                    SetShapeLock(selection[i], targetState);
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ShapeLockingFeature.SetSelectionShapeLock");
                return false;
            }
        }

        public static bool IsShapeLocked(Shape shape)
        {
            try { return ((dynamic)shape).Locked; }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ShapeLockingFeature.IsShapeLocked.Dynamic");
                try { return shape.Tags["oPE_Locked"] == "true"; }
                catch (Exception ex2)
                {
                    ExceptionLogger.Log(ex2, "ShapeLockingFeature.IsShapeLocked.Fallback");
                    return false;
                }
            }
        }

        private static void SetShapeLock(Shape shape, bool lockState)
        {
            try { ((dynamic)shape).Locked = lockState; }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ShapeLockingFeature.SetShapeLock.Locked"); }

            try { shape.LockAspectRatio = lockState ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse; }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ShapeLockingFeature.SetShapeLock.LockAspectRatio"); }

            try { shape.Tags.Add("oPE_Locked", lockState.ToString().ToLower()); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ShapeLockingFeature.SetShapeLock.Tag"); }
        }
    }
}
