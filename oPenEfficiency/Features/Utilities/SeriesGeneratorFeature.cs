using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Controls.Primitives;
using System.Windows.Interop;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnSeriesGenerator",
        Name = "Generate Shape Series",
        SearchAliases = new[] { "Series Generator" },
        Tooltip = "Generate a formatted shape series",
        IconData = "M4 6H2v14c0 1.1.9 2 2 2h14v-2H4V6zm16-4H8c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm0 14H8V4h12v12zm-7-2h2v-2h2v-2h-2V8h-2v2h-2v2h2v2z",
        Color = "#3B82F6",
        Description = "Creates or updates a horizontal, vertical, or matrix series of numbered, lettered, dated, or quarter-labeled shapes.",
        DetailedHelpText = "### Generate Shape Series\n\nOptionally select a template shape first, then create up to 1,000 number, letter, day, month, year, or quarter labels. Right-click opens common 1–5 number, A–E letter, vertical, or 3×3 matrix presets while keeping the preview editable. Choose a horizontal, vertical, or matrix layout and optionally fit it to the slide or a second selected boundary shape. Generated shapes retain series metadata, and a complete generation or update creates one PowerPoint Undo entry.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None,
        IsToggle = true)]
    public static class SeriesGeneratorFeature
    {
        private const int MaximumSeriesItems = 1000;
        public const string TagSeriesId = "oPE_SeriesId";
        public const string TagSeriesConfig = "oPE_SeriesConfig";
        private const string LegacyTagSeriesId = "OPE_SERIES_ID";
        private const string LegacyTagSeriesConfig = "OPE_SERIES_CONFIG";
        private static UI.Dialogs.SeriesGeneratorWindow _window;
        private static ToggleButton _toggleButton;

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, null);
        }

        public static bool Execute(PowerPointManager manager, ToggleButton toggleButton)
        {
            return ExecuteInternal(manager, toggleButton, null);
        }

        public static bool ExecutePreset(
            PowerPointManager manager,
            string seriesType,
            string direction,
            int count,
            int matrixColumns = 3)
        {
            string startValue = string.Equals(seriesType, "Letters", StringComparison.Ordinal)
                ? "A"
                : "1";
            var config = new SeriesConfig
            {
                SeriesType = seriesType,
                Language = "en-US",
                StartValue = startValue,
                EndValue = string.Empty,
                Count = count,
                Step = 1,
                FormatString = string.Equals(seriesType, "Letters", StringComparison.Ordinal) ? string.Empty : "0",
                Direction = direction,
                MatrixColumns = matrixColumns,
                SpacingX = 15,
                SpacingY = 15,
                AutoFitToBounds = true,
                BoundaryType = "Slide"
            };
            return ExecuteInternal(manager, null, config);
        }

        private static bool ExecuteInternal(
            PowerPointManager manager,
            ToggleButton toggleButton,
            SeriesConfig initialConfig)
        {
            if (manager == null) return false;
            try
            {
                if (toggleButton != null && toggleButton.IsChecked == false)
                {
                    _window?.Close();
                    return true;
                }

                if (_window != null && _window.IsLoaded)
                {
                    if (toggleButton != null)
                    {
                        _toggleButton = toggleButton;
                        toggleButton.IsChecked = true;
                    }
                    if (initialConfig != null) _window.SetPreset(initialConfig);
                    _window.Activate();
                    return true;
                }

                _toggleButton = toggleButton;
                _window = new UI.Dialogs.SeriesGeneratorWindow(manager, initialConfig);
                var app = manager.GetApplication();
                if (app != null) new WindowInteropHelper(_window).Owner = (IntPtr)app.HWND;
                _window.Closed += (sender, args) =>
                {
                    if (_toggleButton != null) _toggleButton.IsChecked = false;
                    _toggleButton = null;
                    _window = null;
                };
                if (toggleButton != null) toggleButton.IsChecked = true;
                _window.Show();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SeriesGeneratorFeature.Execute");
                if (toggleButton != null) toggleButton.IsChecked = false;
                return false;
            }
        }

        public static List<string> GenerateSeriesData(SeriesConfig config)
        {
            var result = new List<string>();
            if (config == null || config.Step == 0 ||
                (config.Count <= 0 && string.IsNullOrWhiteSpace(config.EndValue)))
                return result;

            try
            {
                switch (config.SeriesType)
                {
                    case "Numbers":
                        if (int.TryParse(config.StartValue, out int startNum))
                        {
                            int limit = config.Count;
                            bool useEnd = int.TryParse(config.EndValue, out int endNum);
                            int i = 0;
                            while (i < limit || useEnd)
                            {
                                int val = startNum + (i * config.Step);
                                if (useEnd)
                                {
                                    if (config.Step > 0 && val > endNum) break;
                                    if (config.Step < 0 && val < endNum) break;
                                }
                                result.Add(FormatNumber(val, config.FormatString));
                                i++;
                                if (i >= MaximumSeriesItems) break;
                            }
                        }
                        break;

                    case "Letters":
                        char startChar = !string.IsNullOrEmpty(config.StartValue) ? config.StartValue[0] : 'A';
                        if (!IsAsciiLetter(startChar)) startChar = 'A';
                        bool isLower = char.IsLower(startChar);
                        int startCode = char.ToUpper(startChar) - 64; // A=1
                        
                        int endCode = 9999;
                        bool useEndStr = false;
                        if (!string.IsNullOrEmpty(config.EndValue))
                        {
                            string ev = config.EndValue.Trim();
                            if (ev.Length > 0 && IsAsciiLetter(ev[0]))
                            {
                                endCode = char.ToUpper(ev[0]) - 64;
                                useEndStr = true;
                            }
                        }

                        int iL = 0;
                        int limitStr = config.Count;
                        while (iL < limitStr || useEndStr)
                        {
                            int val = startCode + (iL * config.Step);
                            if (useEndStr)
                            {
                                if (config.Step > 0 && val > endCode) break;
                                if (config.Step < 0 && val < endCode) break;
                            }
                            string letterStr = ConvertToLetter(val);
                            if (isLower) letterStr = letterStr.ToLowerInvariant();
                            result.Add(FormatString(letterStr, config.FormatString));
                            iL++;
                            if (iL >= MaximumSeriesItems) break;
                        }
                        break;

                    case "Days":
                    case "Months":
                    case "Years":
                        if (DateTime.TryParse(config.StartValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate))
                        {
                            int iD = 0;
                            bool useEnd = DateTime.TryParse(config.EndValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate);
                            int limit = config.Count;
                            while (iD < limit || useEnd)
                            {
                                DateTime dt = startDate;
                                if (config.SeriesType == "Days") dt = startDate.AddDays(iD * config.Step);
                                else if (config.SeriesType == "Months") dt = startDate.AddMonths(iD * config.Step);
                                else if (config.SeriesType == "Years") dt = startDate.AddYears(iD * config.Step);

                                if (useEnd)
                                {
                                    if (config.Step > 0 && dt > endDate) break;
                                    if (config.Step < 0 && dt < endDate) break;
                                }

                                result.Add(FormatDate(dt, config.FormatString, config.Language));
                                iD++;
                                if (iD >= MaximumSeriesItems) break;
                            }
                        }
                        break;

                    case "Quarters":
                        int startQ = 1;
                        int startY = DateTime.Now.Year;
                        var parts = (config.StartValue ?? string.Empty).Split('-');
                        int parsedValue;
                        if (parts.Length >= 1 && int.TryParse(parts[0], out parsedValue) && parsedValue >= 1)
                            startY = parsedValue;
                        if (parts.Length >= 2 &&
                            int.TryParse(parts[1].Replace("Q", "").Replace("q", ""), out parsedValue) &&
                            parsedValue >= 1 && parsedValue <= 4)
                            startQ = parsedValue;

                        bool useEndQ = false;
                        int endY = 9999, endQ = 4;
                        if (!string.IsNullOrEmpty(config.EndValue))
                        {
                            var eParts = config.EndValue.Split('-');
                            int parsedEndYear;
                            int parsedEndQuarter;
                            if (eParts.Length >= 2 &&
                                int.TryParse(eParts[0], out parsedEndYear) && parsedEndYear >= 1 &&
                                int.TryParse(eParts[1].Replace("Q", "").Replace("q", ""), out parsedEndQuarter) &&
                                parsedEndQuarter >= 1 && parsedEndQuarter <= 4)
                            {
                                endY = parsedEndYear;
                                endQ = parsedEndQuarter;
                                useEndQ = true;
                            }
                        }
                        
                        int iQ = 0;
                        int limitQ = config.Count;
                        int totalStartQ = (startY * 4) + (startQ - 1);
                        int totalEndQ = (endY * 4) + (endQ - 1);

                        while (iQ < limitQ || useEndQ)
                        {
                            int totalQs = totalStartQ + (iQ * config.Step);
                            if (useEndQ)
                            {
                                if (config.Step > 0 && totalQs > totalEndQ) break;
                                if (config.Step < 0 && totalQs < totalEndQ) break;
                            }
                            int y = totalQs / 4;
                            int q = (totalQs % 4) + 1;
                            result.Add(FormatQuarter(y, q, config.FormatString));
                            iQ++;
                            if (iQ >= MaximumSeriesItems) break;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SeriesGeneratorFeature.GenerateSeriesData");
            }
            return result;
        }

        private static string FormatNumber(int value, string format)
        {
            if (string.IsNullOrWhiteSpace(format)) return value.ToString(CultureInfo.InvariantCulture);
            try { return string.Format(CultureInfo.InvariantCulture, format.Contains("{0") ? format : "{0:" + format + "}", value); }
            catch { return value.ToString(); } // Fallback
        }

        private static string FormatString(string value, string format)
        {
            if (string.IsNullOrWhiteSpace(format)) return value;
            try { return string.Format(CultureInfo.InvariantCulture, format.Contains("{0") ? format : "{0}", value); }
            catch { return value; }
        }

        private static string FormatDate(DateTime dt, string format, string language)
        {
            var culture = CultureInfo.CurrentCulture;
            if (!string.IsNullOrWhiteSpace(language))
            {
                try { culture = CultureInfo.GetCultureInfo(language); } catch { /* Unknown language identifiers use invariant series formatting. */ }
            }
            if (string.IsNullOrWhiteSpace(format)) return dt.ToString("d", culture);
            try { return dt.ToString(format, culture); }
            catch { return dt.ToString("d", culture); }
        }

        private static string FormatQuarter(int year, int quarter, string format)
        {
            if (string.IsNullOrWhiteSpace(format)) return $"Q{quarter} {year}";
            
            string result = format;
            // simple custom replacements for quarter formatting
            result = result.Replace("{Q}", $"Q{quarter}");
            result = result.Replace("{q}", $"{quarter}");
            result = result.Replace("{YYYY}", year.ToString());
            result = result.Replace("{yyyy}", year.ToString());
            result = result.Replace("{YY}", (year % 100).ToString("00"));
            result = result.Replace("{yy}", (year % 100).ToString("00"));

            return result;
        }

        private static string ConvertToLetter(int value)
        {
            string result = string.Empty;
            while (--value >= 0)
            {
                result = (char)('A' + value % 26) + result;
                value /= 26;
            }
            return result;
        }

        public static string GetSeriesConfigTag(Shape shape)
        {
            return GetTagValue(shape, TagSeriesConfig, LegacyTagSeriesConfig);
        }

        private static string GetTagValue(Shape shape, string tagName, string legacyTagName)
        {
            if (shape == null) return string.Empty;
            try
            {
                string value = shape.Tags[tagName];
                if (!string.IsNullOrEmpty(value)) return value;
            }
            catch { /* Missing current tag falls back to the supported legacy tag. */ }
            try { return shape.Tags[legacyTagName] ?? string.Empty; }
            catch { return string.Empty; }
        }

        private static bool IsAsciiLetter(char value)
        {
            char upper = char.ToUpperInvariant(value);
            return upper >= 'A' && upper <= 'Z';
        }

        public static bool GenerateShapes(PowerPointManager manager, SeriesConfig config, Shape templateShape = null, Shape boundaryShape = null)
        {
            var createdShapes = new List<Shape>();
            bool oldSeriesDeletionStarted = false;
            try
            {
                if (manager == null || config == null) return false;
                var app = manager.GetApplication();
                if (app == null || app.ActiveWindow == null || app.ActiveWindow.View.Slide == null) return false;
                Slide slide = (Slide)app.ActiveWindow.View.Slide;

                var data = GenerateSeriesData(config);
                if (data.Count == 0) return false;
                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "SeriesGeneratorFeature.StartNewUndoEntry"); }

            string configJson = JsonService.Serialize(config);

            // Find existing series to delete
            var existingShapes = new List<Shape>();
            if (!string.IsNullOrEmpty(config.SeriesId))
            {
                for (int i = slide.Shapes.Count; i >= 1; i--)
                {
                    Shape s = slide.Shapes[i];
                    if (s.Tags != null)
                    {
                        string id = GetTagValue(s, TagSeriesId, LegacyTagSeriesId);
                        if (id == config.SeriesId)
                        {
                            existingShapes.Add(s);
                        }
                    }
                }
            }

            if (string.IsNullOrEmpty(config.SeriesId))
            {
                config.SeriesId = Guid.NewGuid().ToString();
                configJson = JsonService.Serialize(config); // re-serialize with ID
            }

            // Extract template properties if template provided, or from first existing shape
            Shape activeTemplate = templateShape;
            if (activeTemplate == null && existingShapes.Count > 0)
            {
                activeTemplate = existingShapes.Last(); // the one created first (visually)
            }

            float startX = 100f;
            float startY = 100f;
            float width = 100f;
            float height = 50f;
            float fontSize = 14f;

            if (activeTemplate != null)
            {
                startX = activeTemplate.Left;
                startY = activeTemplate.Top;
                width = activeTemplate.Width;
                height = activeTemplate.Height;
                if (activeTemplate.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    float templateFontSize = activeTemplate.TextFrame.TextRange.Font.Size;
                    if (templateFontSize > 0) fontSize = templateFontSize;
                }
            }

            // --- Auto-Fit to Bounds Math ---
            float spacingX = (float)config.SpacingX;
            float spacingY = (float)config.SpacingY;
            float scaleFactor = 1f;

            if (config.AutoFitToBounds)
            {
                oPenEfficiency.Core.Services.SeriesLayoutPlanner.RequiredSize(
                    config.Direction, data.Count, config.MatrixColumns, width, height, spacingX, spacingY,
                    out float reqWidth, out float reqHeight);

                float availWidth = app.ActivePresentation.PageSetup.SlideWidth - startX;
                float availHeight = app.ActivePresentation.PageSetup.SlideHeight - startY;

                if (config.BoundaryType == "SelectedShape" && boundaryShape != null)
                {
                    startX = boundaryShape.Left;
                    startY = boundaryShape.Top;
                    availWidth = boundaryShape.Width;
                    availHeight = boundaryShape.Height;
                }

                // Shrinks only on the binding axis; never enlarges.
                scaleFactor = oPenEfficiency.Core.Services.SeriesLayoutPlanner.AutoFitScale(reqWidth, reqHeight, availWidth, availHeight);

                if (scaleFactor < 1f)
                {
                    width *= scaleFactor;
                    height *= scaleFactor;
                    spacingX *= scaleFactor;
                    spacingY *= scaleFactor;
                    fontSize *= scaleFactor;
                }
            }
            // --- End Auto-Fit Math ---
            bool resizeTemplateFont = activeTemplate != null && scaleFactor < 1f && fontSize > 0f;

            // Generate new ones
            for (int i = 0; i < data.Count; i++)
            {
                Shape newShape;
                if (activeTemplate != null)
                {
                    // Duplicate template
                    try
                    {
                        var sr = activeTemplate.Duplicate();
                        newShape = sr[1];
                        newShape.Width = width;
                        newShape.Height = height;
                        if (resizeTemplateFont && newShape.HasTextFrame == Office.MsoTriState.msoTrue)
                            newShape.TextFrame.TextRange.Font.Size = fontSize;
                    }
                    catch
                    {
                        newShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, width, height);
                    }
                }
                else
                {
                    newShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, width, height);
                    newShape.Fill.ForeColor.RGB = 0xD9D9D9; // Light grey
                    newShape.Line.Visible = Office.MsoTriState.msoFalse;
                    newShape.TextFrame.TextRange.Font.Color.RGB = 0x000000;
                    newShape.TextFrame.TextRange.Font.Size = fontSize;
                    newShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                    newShape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                }
                createdShapes.Add(newShape);

                // Position logic
                oPenEfficiency.Core.Services.SeriesLayoutPlanner.PlaceItem(
                    config.Direction, i, config.MatrixColumns, startX, startY, width, height, spacingX, spacingY,
                    out float x, out float y);

                newShape.Left = x;
                newShape.Top = y;

                // Set text
                if (newShape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    newShape.TextFrame.TextRange.Text = data[i];
                }

                // Add tags
                newShape.Tags.Add(TagSeriesId, config.SeriesId);
                newShape.Tags.Add(TagSeriesConfig, configJson);
            }

            // Keep the old series intact until all replacement shapes exist.
            oldSeriesDeletionStarted = existingShapes.Count > 0;
            foreach (var existingShape in existingShapes)
            {
                if (existingShape == boundaryShape) continue;
                existingShape.Delete();
            }

            // Select new shapes
            if (createdShapes.Count > 0)
            {
                try
                {
                    var arr = createdShapes.Select(s => s.Name).ToArray();
                    slide.Shapes.Range(arr).Select();
                }
                catch { /* A stale generated shape is skipped while restoring series metadata. */ }
            }
                return true;
            }
            catch (Exception ex)
            {
                if (!oldSeriesDeletionStarted)
                {
                    foreach (var createdShape in createdShapes)
                    {
                        try { createdShape.Delete(); }
                        catch (Exception cleanupEx)
                        {
                            ExceptionLogger.Log(cleanupEx, "SeriesGeneratorFeature.GenerateShapes.Cleanup");
                        }
                    }
                }
                ExceptionLogger.Log(ex, "SeriesGeneratorFeature.GenerateShapes");
                return false;
            }
        }
    }
}
