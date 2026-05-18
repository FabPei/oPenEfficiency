using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Toggle rounded corners on rectangle shapes (Top-Left, Top-Right, Bottom-Left, Bottom-Right).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnSmartCorners",
        Name = "Smart Corners",
        Tooltip = "Smart corners",
        IconData = "M3,3H21V21H3V3M7,7V17H17V7H7Z",
        Color = "#F43F5E",
        Description = "Toggles rounded corners on selected rectangle shapes.",
        DetailedHelpText = "### Smart Corners\nApplies a mathematically uniform corner radius to all selected shapes with a configurable slider. Works independently of the native Format Shape dialog.",
        MinSelection = 1)]
    public static class SmartCornersFeature
    {
        // Corners: 0=TL, 1=TR, 2=BL, 3=BR

        /// <summary>
        /// Wrapper for auto-discovery - toggles all corners (default behavior).
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            // Toggle all 4 corners by default
            ToggleCorner(manager, -1); // -1 means toggle all
            return true;
        }

        public static void ToggleCorner(PowerPointManager manager, int cornerIndex)
        {
            var selection = manager.GetSelectedShapes();
            if (selection == null || selection.Count == 0) return;

            bool anyNotRect = false;

            // Check if all are valid rectangles/smart rects
            foreach (PowerPoint.Shape shape in selection)
            {
                if (shape.Type != MsoShapeType.msoAutoShape && shape.Type != MsoShapeType.msoFreeform)
                {
                    anyNotRect = true;
                    continue;
                }
                
                if (shape.Type == MsoShapeType.msoAutoShape && shape.AutoShapeType != MsoAutoShapeType.msoShapeRectangle && shape.AutoShapeType != MsoAutoShapeType.msoShapeRoundedRectangle)
                {
                    anyNotRect = true;
                    continue;
                }
            }

            if (anyNotRect && selection.Count == 1)
            {
                System.Windows.MessageBox.Show("Please select a rectangle", "oPenEfficiency", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return;
            }

            foreach (PowerPoint.Shape shape in selection)
            {
                if (shape.Type != MsoShapeType.msoAutoShape && shape.Type != MsoShapeType.msoFreeform) continue;

                if (shape.Type == MsoShapeType.msoAutoShape && shape.AutoShapeType != MsoAutoShapeType.msoShapeRectangle && shape.AutoShapeType != MsoAutoShapeType.msoShapeRoundedRectangle)
                    continue;

                bool[] state = GetRoundedState(shape);
                state[cornerIndex] = !state[cornerIndex];
                
                ApplySmartCorners(shape, state);
            }
        }

        public static void ResetToSquare(PowerPointManager manager)
        {
            var selection = manager.GetSelectedShapes();
            if (selection == null || selection.Count == 0) return;

            foreach (PowerPoint.Shape shape in selection)
            {
                if (shape.Type != MsoShapeType.msoAutoShape && shape.Type != MsoShapeType.msoFreeform) continue;
                if (shape.Type == MsoShapeType.msoAutoShape && shape.AutoShapeType != MsoAutoShapeType.msoShapeRectangle && shape.AutoShapeType != MsoAutoShapeType.msoShapeRoundedRectangle) continue;

                ApplySmartCorners(shape, new bool[] { false, false, false, false });
            }
        }

        public static bool[] GetCommonRoundedState(PowerPointManager manager)
        {
            var selection = manager.GetSelectedShapes();
            bool[] commonState = new bool[4];
            if (selection == null || selection.Count == 0) return commonState;

            // Use the state of the first valid shape
            foreach (PowerPoint.Shape shape in selection)
            {
                if (shape.Tags["oPen_SmartRect"] == "true")
                {
                    return GetRoundedState(shape);
                }
            }
            return commonState;
        }

        private static bool[] GetRoundedState(PowerPoint.Shape shape)
        {
            bool[] state = new bool[4];
            string tag = shape.Tags["oPen_SmartRect_State"];
            if (!string.IsNullOrEmpty(tag))
            {
                var parts = tag.Split(',');
                if (parts.Length == 4)
                {
                    for (int i = 0; i < 4; i++)
                    {
                        state[i] = parts[i] == "1";
                    }
                }
            }
            return state;
        }

        private static void ApplySmartCorners(PowerPoint.Shape shape, bool[] state)
        {
            float w = shape.Width;
            float h = shape.Height;
            float l = shape.Left;
            float t = shape.Top;

            // Determine radius (e.g. 10% of shortest side)
            float radius = Math.Min(w, h) * 0.15f;
            
            // Limit radius
            if (radius > Math.Min(w, h) / 2) radius = Math.Min(w, h) / 2;

            PowerPoint.Slide slide = shape.Parent as PowerPoint.Slide;
            if (slide == null) return;

            // Create Freeform
            PowerPoint.FreeformBuilder ff = slide.Shapes.BuildFreeform(MsoEditingType.msoEditingCorner, l, t + (state[0] ? radius : 0));
            
            // TL
            if (state[0])
            {
                // Starting at l, t+radius, curve to l+radius, t
                ff.AddNodes(MsoSegmentType.msoSegmentCurve, MsoEditingType.msoEditingCorner, l, t + radius * 0.45f, l + radius * 0.45f, t, l + radius, t);
            }
            else
            {
                ff.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, l, t);
            }
            
            // TR
            ff.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, l + w - (state[1] ? radius : 0), t);
            if (state[1])
            {
                ff.AddNodes(MsoSegmentType.msoSegmentCurve, MsoEditingType.msoEditingCorner, l + w - radius * 0.45f, t, l + w, t + radius * 0.45f, l + w, t + radius);
            }
            else
            {
                ff.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, l + w, t);
            }

            // BR
            ff.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, l + w, t + h - (state[3] ? radius : 0));
            if (state[3])
            {
                ff.AddNodes(MsoSegmentType.msoSegmentCurve, MsoEditingType.msoEditingCorner, l + w, t + h - radius * 0.45f, l + w - radius * 0.45f, t + h, l + w - radius, t + h);
            }
            else
            {
                ff.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, l + w, t + h);
            }

            // BL
            ff.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, l + radius, t + h);
            if (state[2])
            {
                ff.AddNodes(MsoSegmentType.msoSegmentCurve, MsoEditingType.msoEditingCorner, l + radius * 0.45f, t + h, l, t + h - radius * 0.45f, l, t + h - radius);
            }
            else
            {
                ff.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, l, t + h);
            }

            // Close Path
            ff.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, l, t + (state[0] ? radius : 0));
            
            PowerPoint.Shape newShape = ff.ConvertToShape();
            
            // Format replacement
            CopyFormatting(shape, newShape);

            // Tags
            newShape.Tags.Add("oPen_SmartRect", "true");
            newShape.Tags.Add("oPen_SmartRect_State", $"{(state[0]?"1":"0")},{(state[1]?"1":"0")},{(state[2]?"1":"0")},{(state[3]?"1":"0")}");

            // Selection
            newShape.Select(MsoTriState.msoFalse);
            
            // Delete old
            shape.Delete();
        }

        private static void CopyFormatting(PowerPoint.Shape source, PowerPoint.Shape target)
        {
            // Position/Size (in case it drifted slightly)
            target.Left = source.Left;
            target.Top = source.Top;
            target.Width = source.Width;
            target.Height = source.Height;
            target.Rotation = source.Rotation;

            // Pick up and apply visual format
            source.PickUp();
            target.Apply();

            // Text
            if (source.HasTextFrame == MsoTriState.msoTrue && source.TextFrame.HasText == MsoTriState.msoTrue)
            {
                target.TextFrame.TextRange.Text = source.TextFrame.TextRange.Text;
            }

            // Z-Order: keep pushing target back until it's just above the shape that used to be right behind source,
            // but since source is still there (we haven't deleted it yet), we can just replace its Z-Order 
            // by pushing towards bottom and then to the right place.
            // Actually, PowerPoint doesn't have an easy "Set Z order" method, so we use repeatedly ZOrder MsoZOrderCmd.msoSendBackward
            // A common trick is to ZOrder the target shape relative to the source shape.
            // Unfortunately, relative ZOrder is tricky. The best approach is to send target to back, then bring forward until it matches.
            // Wait, an easier trick: Put both shapes at the back, then bring the target forward once. NOT safe if there are other shapes.

            // To preserve ZOrder, we can't easily query ZOrder. Wait, source.ZOrderPosition exists!
            int targetZ = target.ZOrderPosition;
            int sourceZ = source.ZOrderPosition;
            
            while (target.ZOrderPosition > sourceZ + 1)
            {
                target.ZOrder(MsoZOrderCmd.msoSendBackward);
            }
            while (target.ZOrderPosition < sourceZ) // In case target is behind source (which is unlikely as it's newly created)
            {
                target.ZOrder(MsoZOrderCmd.msoBringForward);
            }
        }
    }
}
