using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Services;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert or update a Traffic Light infographic.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTrafficLight",
        Name = "Traffic Light",
        Tooltip = "Traffic light",
        IconData = "M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M19,19H5V5H19V19M7,7H9V17H7V7M11,7H13V17H11V7M15,7H17V17H15V7Z",
        Color = "#F43F5E",
        Description = "Inserts a traffic light infographic (red/yellow/green states). Default is Green state.",
        DetailedHelpText = "### Traffic Light\nInserts a 3-state traffic light indicator (Red / Yellow / Green) as a native shape group. Used for RAG status reporting.",
        Keywords = "traffic lights, rag status, red yellow green, status indicator",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class TrafficLightFeature
    {
        public enum TrafficLightState { Red, Yellow, Green, Off }
        public const string TrafficLightTag = "oPE_TrafficLight";

        // Colors (BGR format for PowerPoint COM)
        private static readonly int TL_Red    = 0x3232E8;   // #E83232 -> BGR 0x3232E8
        private static readonly int TL_Yellow = 0x17C8F5;   // #F5C817 -> BGR 0x17C8F5
        private static readonly int TL_Green  = 0x3EAF4C;   // #4CAF3E -> BGR 0x3EAF4C
        private static readonly int TL_Dim    = 0xCCCCCC;   // Grey dimmed
        private static readonly int TL_Frame  = 0x404040;   // Dark grey frame

        /// <summary>
        /// Wrapper for auto-discovery - inserts traffic light with default Green state.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, TrafficLightState.Green, horizontal: true, hasBorder: false, borderWeight: 1f);
        }

        public static bool Execute(PowerPointManager manager, TrafficLightState state = TrafficLightState.Green, bool horizontal = true, bool hasBorder = false, float borderWeight = 1f)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;

                // Update existing
                var selection = manager.GetSelectedShapes();
                if (selection != null && selection.Count == 1)
                {
                    var selected = selection[1];
                    string opeState = ShapeMetadataService.GetTag(selected, "oPE_State");
                    bool isOpe = (opeState != null && opeState.StartsWith(TrafficLightTag)) || 
                                 (selected.AlternativeText != null && selected.AlternativeText.StartsWith(TrafficLightTag));
                    bool isEE4P = ShapeMetadataService.GetTag(selected, "EE4P_SMART_ELEMENT") == "TrafficLight";

                    if (isOpe || isEE4P)
                    {
                        UpdateTrafficLight(selected, state, horizontal, hasBorder, borderWeight);
                        return true;
                    }
                }

                // Insert new one
                var slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                float circleSize = 18f;
                float padding = 4f;
                float gap = 3f;

                float frameWidth, frameHeight;
                if (horizontal)
                {
                    frameWidth = padding * 2 + circleSize * 3 + gap * 2;
                    frameHeight = padding * 2 + circleSize;
                }
                else
                {
                    frameWidth = padding * 2 + circleSize;
                    frameHeight = padding * 2 + circleSize * 3 + gap * 2;
                }

                float left = (app.ActivePresentation.PageSetup.SlideWidth / 2) - (frameWidth / 2);
                float top = (app.ActivePresentation.PageSetup.SlideHeight / 2) - (frameHeight / 2);

                // Frame
                var frame = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                    left, top, frameWidth, frameHeight);
                frame.Fill.ForeColor.RGB = TL_Frame;
                frame.Line.Visible = hasBorder ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                if (hasBorder)
                {
                    frame.Line.ForeColor.RGB = 0x000000;
                    frame.Line.Weight = borderWeight;
                }
                try { frame.Adjustments[1] = 0.2f; } catch (Exception ex) { ExceptionLogger.Log(ex, "TrafficLightFeature.Execute.FrameAdjustments"); }

                // Circles
                float c1Left, c1Top, c2Left, c2Top, c3Left, c3Top;
                if (horizontal)
                {
                    float cy = top + padding;
                    c1Left = left + padding;                       c1Top = cy;
                    c2Left = left + padding + circleSize + gap;    c2Top = cy;
                    c3Left = left + padding + (circleSize + gap) * 2; c3Top = cy;
                }
                else
                {
                    float cx = left + padding;
                    c1Left = cx; c1Top = top + padding;
                    c2Left = cx; c2Top = top + padding + circleSize + gap;
                    c3Left = cx; c3Top = top + padding + (circleSize + gap) * 2;
                }

                var redCircle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, c1Left, c1Top, circleSize, circleSize);
                var yelCircle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, c2Left, c2Top, circleSize, circleSize);
                var grnCircle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, c3Left, c3Top, circleSize, circleSize);

                foreach (Shape c in new[] { redCircle, yelCircle, grnCircle })
                {
                    c.Line.Visible = Office.MsoTriState.msoFalse;
                    c.Fill.ForeColor.RGB = TL_Dim;
                }

                switch (state)
                {
                    case TrafficLightState.Red:    redCircle.Fill.ForeColor.RGB = TL_Red; break;
                    case TrafficLightState.Yellow: yelCircle.Fill.ForeColor.RGB = TL_Yellow; break;
                    case TrafficLightState.Green:  grnCircle.Fill.ForeColor.RGB = TL_Green; break;
                    case TrafficLightState.Off:    break;
                }

                // Group
                var names = new[] { frame.Name, redCircle.Name, yelCircle.Name, grnCircle.Name };
                var group = slide.Shapes.Range(names).Group();
                string orient = horizontal ? "H" : "V";
                group.AlternativeText = $"{TrafficLightTag}|{(int)state}|{orient}|{(hasBorder ? 1 : 0)}|{borderWeight}";
                group.Name = "TrafficLight_" + Guid.NewGuid().ToString().Substring(0, 8);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TrafficLightFeature.Execute");
                return false;
            }
        }

        private static void UpdateTrafficLight(Shape group, TrafficLightState state, bool horizontal, bool hasBorder, float borderWeight)
        {
            try
            {
                if (group.Type != Office.MsoShapeType.msoGroup) return;
                var items = group.GroupItems;
                if (items.Count < 4) return;

                var frame = items[1];
                var red   = items[2];
                var yel   = items[3];
                var grn   = items[4];

                red.Fill.ForeColor.RGB = (state == TrafficLightState.Red) ? TL_Red : TL_Dim;
                yel.Fill.ForeColor.RGB = (state == TrafficLightState.Yellow) ? TL_Yellow : TL_Dim;
                grn.Fill.ForeColor.RGB = (state == TrafficLightState.Green) ? TL_Green : TL_Dim;

                frame.Line.Visible = hasBorder ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                if (hasBorder)
                {
                    frame.Line.ForeColor.RGB = 0x000000;
                    frame.Line.Weight = borderWeight;
                }

                float circleSize = red.Width;
                float padding = 4f;
                float gap = 3f;

                float frameWidth, frameHeight;
                if (horizontal)
                {
                    frameWidth = padding * 2 + circleSize * 3 + gap * 2;
                    frameHeight = padding * 2 + circleSize;
                }
                else
                {
                    frameWidth = padding * 2 + circleSize;
                    frameHeight = padding * 2 + circleSize * 3 + gap * 2;
                }

                frame.Width = frameWidth;
                frame.Height = frameHeight;

                if (horizontal)
                {
                    float cy = frame.Top + padding;
                    red.Left = frame.Left + padding;                          red.Top = cy;
                    yel.Left = frame.Left + padding + circleSize + gap;       yel.Top = cy;
                    grn.Left = frame.Left + padding + (circleSize + gap) * 2; grn.Top = cy;
                }
                else
                {
                    float cx = frame.Left + padding;
                    red.Left = cx; red.Top = frame.Top + padding;
                    yel.Left = cx; yel.Top = frame.Top + padding + circleSize + gap;
                    grn.Left = cx; grn.Top = frame.Top + padding + (circleSize + gap) * 2;
                }

                string orient = horizontal ? "H" : "V";
                ShapeMetadataService.SetTag(group, "oPE_State", $"{TrafficLightTag}|{(int)state}|{orient}|{(hasBorder ? 1 : 0)}|{borderWeight}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"TrafficLightFeature.SetState error: {ex.Message}");
            }
        }
    }
}
