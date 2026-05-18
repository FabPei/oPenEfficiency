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
        public static bool HideSelectedShapes(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var shapes = manager.GetSelectedShapes();
                if (shapes == null || shapes.Count == 0) return false;

                foreach (Shape s in shapes)
                {
                    s.Visible = Office.MsoTriState.msoFalse;
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "VisibilityFeature.HideSelectedShapes");
                return false;
            }
        }

        public static bool ShowAllHiddenShapes(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null || app.ActiveWindow.View.Slide == null) return false;

                var slide = app.ActiveWindow.View.Slide as Slide;
                bool found = false;
                foreach (Shape s in slide.Shapes)
                {
                    if (s.Visible == Office.MsoTriState.msoFalse)
                    {
                        s.Visible = Office.MsoTriState.msoTrue;
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
    }
}
