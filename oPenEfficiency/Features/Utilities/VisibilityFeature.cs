using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Linq;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Hide selected shapes or reveal hidden shapes on the current slide.
    /// </summary>
    public static class VisibilityFeature
    {
        private const string HiddenByFeatureTag = "oPE_HiddenByFeature";

        public static bool HideSelectedShapes(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var shapes = manager.GetSelectedShapes();
                if (shapes == null || shapes.Count == 0) return false;

                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "VisibilityFeature.HideSelectedShapes.StartUndo"); }

                bool changed = false;
                foreach (Shape s in shapes)
                {
                    if (s.Visible == Office.MsoTriState.msoFalse) continue;
                    // Do not hide a shape unless it can be identified for safe restoration.
                    s.Tags.Add(HiddenByFeatureTag, "true");
                    s.Visible = Office.MsoTriState.msoFalse;
                    changed = true;
                }
                return changed;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "VisibilityFeature.HideSelectedShapes");
                return false;
            }
        }

        public static bool ShowAllHiddenShapes(PowerPointManager manager, bool includeShapesHiddenOutsideOpenEfficiency = false)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null || app.ActiveWindow.View.Slide == null) return false;

                var slide = app.ActiveWindow.View.Slide as Slide;
                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "VisibilityFeature.ShowAllHiddenShapes.StartUndo"); }
                bool found = false;
                foreach (Shape s in slide.Shapes)
                {
                    bool hiddenByFeature = HasTag(s, HiddenByFeatureTag);
                    if (s.Visible == Office.MsoTriState.msoFalse && (includeShapesHiddenOutsideOpenEfficiency || hiddenByFeature))
                    {
                        s.Visible = Office.MsoTriState.msoTrue;
                        if (hiddenByFeature) TryDeleteTag(s, HiddenByFeatureTag);
                        found = true;
                    }
                }
                return found;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "VisibilityFeature.ShowAllHiddenShapes");
                return false;
            }
        }

        public static bool IsShapeHidden(Shape shape)
        {
            return shape != null && shape.Visible == Office.MsoTriState.msoFalse;
        }

        private static bool HasTag(Shape shape, string name)
        {
            try { return string.Equals(shape.Tags[name], "true", StringComparison.OrdinalIgnoreCase); }
            catch { return false; }
        }

        private static void TryDeleteTag(Shape shape, string name)
        {
            try { shape.Tags.Delete(name); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "VisibilityFeature.TryDeleteTag"); }
        }
    }
}
