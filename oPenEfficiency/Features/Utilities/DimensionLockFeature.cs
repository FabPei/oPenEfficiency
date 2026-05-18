using System;
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
        Name = "Dimension Lock",
        Tooltip = "Dimension lock",
        IconData = "M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M19,19H5V5H19V19M12,7H14V12H12V7M12,14H14V17H12V14Z",
        Color = "#6B7280",
        Description = "Locks or unlocks the dimensions (width and height) of selected shapes.",
        DetailedHelpText = "### Dimension Lock\nPrevents the width and height of selected shapes from being changed while allowing them to be moved.\n**Right-Click Options:**\n* **Global Toggle**: Apply or remove dimension locks for all shapes across the entire presentation.",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes,
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

                bool targetState = true;
                if (forceState == null)
                {
                    targetState = !IsDimensionLocked(selection[1]);
                }
                else targetState = forceState.Value;

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

                bool targetState = true;
                if (forceState == null)
                {
                    bool found = false;
                    foreach (Slide slide in app.ActivePresentation.Slides)
                    {
                        if (slide.Shapes.Count > 0)
                        {
                            targetState = !IsDimensionLocked(slide.Shapes[1]);
                            found = true;
                            break;
                        }
                    }
                    if (!found) return false;
                }
                else targetState = forceState.Value;

                foreach (Slide slide in app.ActivePresentation.Slides)
                {
                    foreach (Shape shape in slide.Shapes) SetDimensionLock(shape, targetState);
                }
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
                // Store custom tag for dimension lock state
                shape.Tags.Add("oPE_DimensionLocked", lockState.ToString().ToLower());
            }
            catch (Exception)
            {
                try
                {
                    // If tag already exists, delete it and add again
                    shape.Tags.Delete("oPE_DimensionLocked");
                    shape.Tags.Add("oPE_DimensionLocked", lockState.ToString().ToLower());
                }
                catch (Exception ex2)
                {
                    ExceptionLogger.Log(ex2, "DimensionLockFeature.SetDimensionLock.TagUpdate");
                }
            }
        }
    }
}
