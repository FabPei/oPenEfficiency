using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(Id = "BtnSeriesGenerator", Name = "Series Generator", Tooltip = "Generate Series (Numbers, Dates, Letters)", IconData = "M4 6H2v14c0 1.1.9 2 2 2h14v-2H4V6zm16-4H8c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm0 14H8V4h12v12zm-7-2h2v-2h2v-2h-2V8h-2v2h-2v2h2v2z", Color = "#3B82F6", Description = "Generates a sequence of numbers, letters, dates, or quarters based on a template shape.", Keywords = "series, sequence, timeline, dates, quarters, generate", MinSelection = 0, RequiredType = PpSelectionType.ppSelectionNone)]
    public static class SeriesGeneratorFeature
    {
        public const string TagSeriesId = "OPE_SERIES_ID";
        public const string TagSeriesConfig = "OPE_SERIES_CONFIG";

        public static bool Execute(PowerPointManager manager)
        {
            var window = new UI.Dialogs.SeriesGeneratorWindow(manager);
            window.Show();
            return true;
        }

        public static List<string> GenerateSeriesData(SeriesConfig config)
        {
            var result = new List<string>();
            if (config.Count <= 0 && string.IsNullOrWhiteSpace(config.EndValue)) return result;

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
                                if (i > 1000) break; // sanity cap
                            }
                        }
                        break;

                    case "Letters":
                        char startChar = config.StartValue.Length > 0 ? config.StartValue[0] : 'A';
                        if (!char.IsLetter(startChar)) startChar = 'A';
                        bool isLower = char.IsLower(startChar);
                        int startCode = char.ToUpper(startChar) - 64; // A=1
                        
                        int endCode = 9999;
                        bool useEndStr = false;
                        if (!string.IsNullOrEmpty(config.EndValue))
                        {
                            string ev = config.EndValue.Trim();
                            if (ev.Length > 0 && char.IsLetter(ev[0]))
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
                            if (iL > 1000) break;
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
                                if (iD > 1000) break;
                            }
                        }
                        break;

                    case "Quarters":
                        int startQ = 1;
                        int startY = DateTime.Now.Year;
                        var parts = config.StartValue.Split('-');
                        if (parts.Length >= 1) int.TryParse(parts[0], out startY);
                        if (parts.Length >= 2) int.TryParse(parts[1].Replace("Q", "").Replace("q", ""), out startQ);

                        bool useEndQ = false;
                        int endY = 9999, endQ = 4;
                        if (!string.IsNullOrEmpty(config.EndValue))
                        {
                            var eParts = config.EndValue.Split('-');
                            if (eParts.Length >= 1) int.TryParse(eParts[0], out endY);
                            if (eParts.Length >= 2) int.TryParse(eParts[1].Replace("Q", "").Replace("q", ""), out endQ);
                            useEndQ = true;
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
                            if (iQ > 1000) break;
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
            if (string.IsNullOrWhiteSpace(format)) return value.ToString();
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
                try { culture = CultureInfo.GetCultureInfo(language); } catch { }
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

        public static void GenerateShapes(PowerPointManager manager, SeriesConfig config, Shape templateShape = null, Shape boundaryShape = null)
        {
            var app = manager.GetApplication();
            if (app == null || app.ActiveWindow == null || app.ActiveWindow.View.Slide == null) return;
            Slide slide = (Slide)app.ActiveWindow.View.Slide;

            var data = GenerateSeriesData(config);
            if (data.Count == 0) return;

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
                        string id = "";
                        try { id = s.Tags[TagSeriesId]; } catch { }
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
                    fontSize = activeTemplate.TextFrame.TextRange.Font.Size;
                }
            }

            // --- Auto-Fit to Bounds Math ---
            float spacingX = (float)config.SpacingX;
            float spacingY = (float)config.SpacingY;
            float scaleFactor = 1f;

            if (config.AutoFitToBounds)
            {
                float reqWidth = 0f;
                float reqHeight = 0f;
                int c = data.Count;

                if (config.Direction == "Horizontal")
                {
                    reqWidth = (c * width) + ((c - 1) * spacingX);
                    reqHeight = height;
                }
                else if (config.Direction == "Vertical")
                {
                    reqWidth = width;
                    reqHeight = (c * height) + ((c - 1) * spacingY);
                }
                else if (config.Direction == "Matrix" && config.MatrixColumns > 0)
                {
                    int cols = Math.Min(c, config.MatrixColumns);
                    int rows = (int)Math.Ceiling(c / (double)config.MatrixColumns);
                    reqWidth = (cols * width) + ((cols - 1) * spacingX);
                    reqHeight = (rows * height) + ((rows - 1) * spacingY);
                }

                float availWidth = app.ActivePresentation.PageSetup.SlideWidth - startX;
                float availHeight = app.ActivePresentation.PageSetup.SlideHeight - startY;

                if (config.BoundaryType == "SelectedShape" && boundaryShape != null)
                {
                    startX = boundaryShape.Left;
                    startY = boundaryShape.Top;
                    availWidth = boundaryShape.Width;
                    availHeight = boundaryShape.Height;
                }

                float scaleX = reqWidth > availWidth && reqWidth > 0 ? availWidth / reqWidth : 1f;
                float scaleY = reqHeight > availHeight && reqHeight > 0 ? availHeight / reqHeight : 1f;
                scaleFactor = Math.Min(scaleX, scaleY);
                
                // Cap to max 1f (we don't want to enlarge, only shrink)
                scaleFactor = Math.Min(1f, scaleFactor);

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

            // Delete old shapes if updating
            foreach (var s in existingShapes)
            {
                if (s != activeTemplate && s != boundaryShape) // Don't delete the template or boundary
                {
                    try { s.Delete(); } catch { }
                }
            }

            var createdShapes = new List<Shape>();

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
                        if (newShape.HasTextFrame == Office.MsoTriState.msoTrue) newShape.TextFrame.TextRange.Font.Size = fontSize;
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

                // Position logic
                float x = startX;
                float y = startY;

                if (config.Direction == "Horizontal")
                {
                    x = startX + (i * (width + spacingX));
                }
                else if (config.Direction == "Vertical")
                {
                    y = startY + (i * (height + spacingY));
                }
                else if (config.Direction == "Matrix" && config.MatrixColumns > 0)
                {
                    int col = i % config.MatrixColumns;
                    int row = i / config.MatrixColumns;
                    x = startX + (col * (width + spacingX));
                    y = startY + (row * (height + spacingY));
                }

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
                
                createdShapes.Add(newShape);
            }

            // Cleanup template if we duplicated it during an update and it wasn't part of the newly generated list
            if (activeTemplate != null && existingShapes.Contains(activeTemplate))
            {
                try { activeTemplate.Delete(); } catch { }
            }

            // Select new shapes
            if (createdShapes.Count > 0)
            {
                try
                {
                    var arr = createdShapes.Select(s => s.Name).ToArray();
                    slide.Shapes.Range(arr).Select();
                }
                catch { }
            }
        }
    }
}
