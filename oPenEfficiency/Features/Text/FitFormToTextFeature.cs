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
        Name = "Fit Form to Text",
        Tooltip = "Fit Form to Text",
        IconData = "M3,17H21V19H3V17 M3,5H21V7H3V5 M13,9V15H11V9H13 M15,9V11H17V9H15 M15,13V15H17V13H15 M7,9V11H9V9H7 M7,13V15H9V13H7", // A custom representation
        Color = "#FDE047",
        MinSelection = 1,
        RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes,
        Description = "Adjusts the bounding box of the selected shapes to exactly fit the text string within, removing asymmetrical margins and empty whitespace.",
        DetailedHelpText = "### Fit Form to Text\n\nSnaps the bounding box of the selected shape exactly to the dimensions of the contained text. Eliminates unwanted whitespace padding and ensures the shape tightly wraps its content.",
        Keywords = "auto size, shrink to fit, adjust bounding box, remove padding, zero margins, tight fit"
    )]
    public static class FitFormToTextFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var selection = manager.GetApplication().ActiveWindow.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return false;

                bool modified = false;

                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    modified |= ProcessShape(shape);
                }

                return modified;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex.Message, "FitFormToTextFeature");
                return false;
            }
        }

        private static bool ProcessShape(PowerPoint.Shape shape)
        {
            bool modified = false;

            if (shape.Type == MsoShapeType.msoGroup)
            {
                foreach (PowerPoint.Shape child in shape.GroupItems)
                {
                    modified |= ProcessShape(child);
                }
                return modified;
            }

            if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
            {
                shape.TextFrame.MarginLeft = 0;
                shape.TextFrame.MarginRight = 0;
                shape.TextFrame.MarginTop = 0;
                shape.TextFrame.MarginBottom = 0;
                
                // Disable WordWrap temporarily so the shape shrinks to exactly the line width
                shape.TextFrame.WordWrap = MsoTriState.msoFalse;
                shape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                
                modified = true;
            }

            return modified;
        }
    }
}
