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
        DetailedHelpText = "### Progress Series\nInserts a horizontal progress bar or segmented series bar with adjustable fill percentage. Ideal for KPI tracking slides and dashboards.",
        Keywords = "progress bar, segmented bar, kpi tracker, linear progress",
        MinSelection = 0,
        RequiredType = PowerPoint.PpSelectionType.ppSelectionNone)]
    public static class ProgressSeriesFeature
    {
        public const string SeriesTag = "oPE_Series";

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
                Count = 5,
                ActiveIndex = 3,
                Mode = "Numbers",
                Layout = "H",
                HighlightColor = "#E53935",
                ShapeType = "Square"
            };

            public string ToTag() => $"{SeriesTag}|{Count}|{ActiveIndex}|{Mode}|{Layout}|{HighlightColor}|{ShapeType}";

            public static SeriesSettings FromTag(string tag)
            {
                var settings = Default();
                if (string.IsNullOrEmpty(tag) || !tag.StartsWith(SeriesTag)) return settings;

                var parts = tag.Split('|');
                if (parts.Length >= 2) int.TryParse(parts[1], out settings.Count);
                if (parts.Length >= 3) int.TryParse(parts[2], out settings.ActiveIndex);
                if (parts.Length >= 4) settings.Mode = parts[3];
                if (parts.Length >= 5) settings.Layout = parts[4];
                if (parts.Length >= 6) settings.HighlightColor = parts[5];
                if (parts.Length >= 7) settings.ShapeType = parts[6];

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
                            catch { }
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
