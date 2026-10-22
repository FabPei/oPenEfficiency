using System;
using System.Drawing;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Services;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnProgressSeries",
        Name = "Progress Series",
        Tooltip = "Progress series",
        IconData = "M4,2H20V20H4V2M6,4V18H18V4H6M8,6H16V16H8V6Z",
        Color = "#F43F5E",
        Description = "Inserts a progress series visualization with configurable shape type and layout.",
        DetailedHelpText = "### Progress Series\nInserts an editable numbered or bullet progress series. Right-click for common three-step, five-step, horizontal, vertical, and bullet presets; selecting an existing series updates it in place. If you manually recolor the active segment in PowerPoint, opening the contextual toolbar reads that visible fill without ungrouping the series and retains it until you choose another color.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class ProgressSeriesFeature
    {
        public const string SeriesTag = oPenEfficiency.Core.Services.ProgressSeriesStatePolicy.SeriesTag;

        public struct SeriesSettings
        {
            public int Count;
            public int ActiveIndex;
            public string Mode; // "Numbers", "Letters", "SmallLetters", "Roman", "SmallRoman", "Bullets"
            public string Layout; // "H" or "V"
            public string HighlightColor; // Hex string e.g. "#0078D4"
            public string ShapeType; // "Square", "Circle", "Triangle", "Star", "Check", "Cross", "Dot"

            public static SeriesSettings Default() => new SeriesSettings
            {
                Count = oPenEfficiency.Core.Services.ProgressSeriesStatePolicy.DefaultCount,
                ActiveIndex = oPenEfficiency.Core.Services.ProgressSeriesStatePolicy.DefaultActiveIndex,
                Mode = oPenEfficiency.Core.Services.ProgressSeriesStatePolicy.DefaultMode,
                Layout = oPenEfficiency.Core.Services.ProgressSeriesStatePolicy.DefaultLayout,
                HighlightColor = oPenEfficiency.Core.Services.ProgressSeriesStatePolicy.DefaultHighlightColor,
                ShapeType = oPenEfficiency.Core.Services.ProgressSeriesStatePolicy.DefaultShapeType
            };

            public string ToTag() => oPenEfficiency.Core.Services.ProgressSeriesStatePolicy.FormatTag(Count, ActiveIndex, Mode, Layout, HighlightColor, ShapeType);

            public static SeriesSettings FromTag(string tag)
            {
                var settings = Default();
                if (oPenEfficiency.Core.Services.ProgressSeriesStatePolicy.TryParseTag(tag, out int count, out int activeIndex, out string mode, out string layout, out string highlightColor, out string shapeType))
                {
                    settings.Count = count;
                    settings.ActiveIndex = activeIndex;
                    settings.Mode = mode;
                    settings.Layout = layout;
                    settings.HighlightColor = highlightColor;
                    settings.ShapeType = shapeType;
                }

                return settings;
            }
        }

        /// <summary>
        /// Wrapper for auto-discovery - creates progress series with default settings (5 items, 3 active, numbers, horizontal, squares).
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            Execute(manager, SeriesSettings.Default());
            return true;
        }

        public static void Create(PowerPointManager manager)
        {
            Execute(manager);
        }

        public static bool IsProgressSeries(PowerPoint.Shape shape)
        {
            if (shape == null) return false;
            if (shape.AlternativeText != null && shape.AlternativeText.StartsWith(SeriesTag)) return true;
            
            string ee4pTag = ShapeMetadataService.GetTag(shape, "EE4P_SMART_ELEMENT");
            if (ee4pTag == "PercentBar" || ee4pTag == "Indicator") return true;

            return false;
        }

        public static void Execute(PowerPointManager manager, SeriesSettings settings, PowerPoint.Shape targetShape = null)
        {
            if (manager == null) return;
            
            PowerPoint.Slide slide = null;
            PowerPoint.ShapeRange selection = null;
            PowerPoint.Shape existingGroup = targetShape;
            
            try
            {
                slide = manager.GetCurrentSlide();
                if (slide == null) return;

                // If we have a single selected series, we update it by deleting and recreating or modifying
                // Recreating is safer for layout changes
                float left = 100, top = 100;

                if (existingGroup == null)
                {
                    selection = manager.GetSelectedShapes();
                    if (selection != null && selection.Count == 1)
                    {
                        var s = selection[1];
                        if (IsProgressSeries(s))
                        {
                            existingGroup = s;
                        }
                        else if (s.Child == Office.MsoTriState.msoTrue)
                        {
                            try
                            {
                                var parent = s.ParentGroup;
                                if (IsProgressSeries(parent))
                                {
                                    existingGroup = parent;
                                }
                            }
                            catch { /* Optional parent-group detection falls back to the selected segment. */ }
                        }
                    }
                }

                if (existingGroup != null)
                {
                    left = existingGroup.Left;
                    top = existingGroup.Top;
                }

            float size = 30;
            float gap = 10;
            var shapeNames = new System.Collections.Generic.List<string>();

            for (int i = 1; i <= settings.Count; i++)
            {
                float x = left;
                float y = top;

                if (settings.Layout == "H")
                    x += (i - 1) * (size + gap);
                else
                    y += (i - 1) * (size + gap);

                var pptShapeType = Office.MsoAutoShapeType.msoShapeRectangle;
                switch (settings.ShapeType)
                {
                    case "Circle": pptShapeType = Office.MsoAutoShapeType.msoShapeOval; break;
                    case "Diamond": pptShapeType = Office.MsoAutoShapeType.msoShapeDiamond; break;
                    case "Triangle": pptShapeType = Office.MsoAutoShapeType.msoShapeIsoscelesTriangle; break;
                    case "Star": pptShapeType = Office.MsoAutoShapeType.msoShape5pointStar; break;
                    case "Check": pptShapeType = Office.MsoAutoShapeType.msoShapeMathEqual; break;
                    case "Cross": pptShapeType = Office.MsoAutoShapeType.msoShapeRectangle; break;
                    case "Dot": pptShapeType = Office.MsoAutoShapeType.msoShapeOval; break;
                    default: pptShapeType = Office.MsoAutoShapeType.msoShapeRectangle; break;
                }

                var shape = slide.Shapes.AddShape(pptShapeType, x, y, size, size);
                
                // Content
                string text = GetTextForIndex(i, settings.Mode);
                
                // Override text for Cross shape to show ✕ symbol
                if (settings.ShapeType == "Cross")
                {
                    text = "✕";
                }
                
                shape.TextFrame.TextRange.Text = text;
                shape.TextFrame.MarginLeft = 0;
                shape.TextFrame.MarginRight = 0;
                shape.TextFrame.MarginTop = 0;
                shape.TextFrame.MarginBottom = 0;
                shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                shape.TextFrame.TextRange.Font.Size = 14;
                shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF;

                // Color
                if (i == settings.ActiveIndex)
                {
                    shape.Fill.ForeColor.RGB = ColorTranslator.ToWin32(ColorTranslator.FromHtml(settings.HighlightColor));
                }
                else
                {
                    // Light Gray for non-active
                    shape.Fill.ForeColor.RGB = ColorTranslator.ToWin32(Color.FromArgb(210, 210, 210));
                    shape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToWin32(Color.FromArgb(100, 100, 100));
                }

                shape.Line.Visible = Office.MsoTriState.msoFalse;
                shapeNames.Add(shape.Name);
            }

            var range = slide.Shapes.Range(shapeNames.ToArray());
            var group = range.Group();
            string tag = settings.ToTag();
            ShapeMetadataService.SetTag(group, "oPE_State", tag);
            group.AlternativeText = tag;
            group.Name = "oPE_Series_" + Guid.NewGuid().ToString("N").Substring(0, 8);

            if (existingGroup != null)
            {
                existingGroup.Delete();
            }

            group.Select();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ProgressSeriesFeature.Execute error: {ex.Message}");
            }
            finally
            {
                // Release COM objects to prevent ObjectDisposedException
                try { if (existingGroup != null) Marshal.ReleaseComObject(existingGroup); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"ProgressSeriesFeature.Release existingGroup error: {ex.Message}"); }
                try { if (selection != null) Marshal.ReleaseComObject(selection); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"ProgressSeriesFeature.Release selection error: {ex.Message}"); }
                try { if (slide != null) Marshal.ReleaseComObject(slide); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"ProgressSeriesFeature.Release slide error: {ex.Message}"); }
            }
        }

        private static string GetTextForIndex(int index, string mode)
        {
            switch (mode)
            {
                case "Letters":
                    return GetLetter(index);
                case "SmallLetters":
                    return GetLetter(index).ToLower();
                case "Roman":
                    return GetRomanNumeral(index);
                case "SmallRoman":
                    return GetRomanNumeral(index).ToLower();
                case "Bullets":
                    return "•";
                default:
                    return index.ToString();
            }
        }

        private static string GetLetter(int index)
        {
            if (index <= 0) return "A";
            int value = index - 1;
            string result = "";
            do
            {
                result = (char)('A' + (value % 26)) + result;
                value = (value / 26) - 1;
            } while (value >= 0);
            return result;
        }

        private static string GetRomanNumeral(int index)
        {
            if (index <= 0) return "I";
            if (index > 10) return index.ToString(); // Fallback for large numbers
            
            string[] romanNumerals = { "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X" };
            return index < romanNumerals.Length ? romanNumerals[index] : index.ToString();
        }
    }
}
