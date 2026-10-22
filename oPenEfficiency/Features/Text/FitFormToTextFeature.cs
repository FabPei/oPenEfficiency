using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Text
{
    [FeatureMetadata(
        Id = "BtnFitFormToText",
        Name = "Fit Shape to Text",
        Tooltip = "Fit shape tightly to its text",
        IconData = "M3,17H21V19H3V17 M3,5H21V7H3V5 M13,9V15H11V9H13 M15,9V11H17V9H15 M15,13V15H17V13H15 M7,9V11H9V9H7 M7,13V15H9V13H7", // A custom representation
        Color = "#FDE047",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes,
        Description = "Fits selected text-bearing shapes tightly around their text, including supported shapes inside groups.",
        DetailedHelpText = "### Fit Shape to Text\n\nFor each selected text-bearing shape, including supported children of selected groups, disables word wrapping and lets PowerPoint resize the shape to fit its text. The primary action also removes all four text margins. Right-click to choose fitting with zero margins or while preserving the existing margins. Shapes without text are skipped and the complete action creates one PowerPoint Undo entry."
    )]
    public static class FitFormToTextFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, clearMargins: true);
        }

        public static bool Execute(PowerPointManager manager, bool clearMargins)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app == null || app.ActiveWindow == null) return false;
                var selection = app.ActiveWindow.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return false;

                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "FitFormToTextFeature.StartNewUndoEntry"); }

                bool modified = false;
                bool failed = false;

                for (int i = 1; i <= selection.ShapeRange.Count; i++)
                {
                    modified |= ProcessShape(selection.ShapeRange[i], clearMargins, ref failed);
                }

                return modified && !failed;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FitFormToTextFeature.Execute");
                return false;
            }
        }

        private static bool ProcessShape(PowerPoint.Shape shape, bool clearMargins, ref bool failed)
        {
            bool modified = false;

            if (shape.Type == MsoShapeType.msoGroup)
            {
                for (int i = 1; i <= shape.GroupItems.Count; i++)
                {
                    modified |= ProcessShape(shape.GroupItems[i], clearMargins, ref failed);
                }
                return modified;
            }

            if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
            {
                try
                {
                    if (clearMargins)
                    {
                        shape.TextFrame.MarginLeft = 0;
                        shape.TextFrame.MarginRight = 0;
                        shape.TextFrame.MarginTop = 0;
                        shape.TextFrame.MarginBottom = 0;
                    }
                    shape.TextFrame.WordWrap = MsoTriState.msoFalse;
                    shape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    modified = true;
                }
                catch (Exception ex)
                {
                    failed = true;
                    ExceptionLogger.Log(ex, "FitFormToTextFeature.ProcessShape");
                }
            }

            return modified;
        }
    }
}
