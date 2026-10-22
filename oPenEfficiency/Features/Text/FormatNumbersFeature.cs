using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Linq;
using oPenEfficiency.Utils;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Reformat numbers within selected text, shapes, or slides.
    /// Supports thousand separators, decimal separators, and currency placement.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnFormatNumbers",
        Name = "Format Numbers",
        SearchAliases = new[] { "Number Formatter" },
        Tooltip = "Format numbers",
        IconData = "M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M19,19H5V5H19V19M7.5,13V15H9.5V13H7.5M12.5,11V15H14.5V11H12.5M17.5,9V15H19.5V9H17.5M7.5,7V11H9.5V7H7.5M12.5,5V9H14.5V5H12.5Z",
        Color = "#FBBF24",
        Description = "Reformats detected numbers with configurable grouping, decimal separators, and currency placement.",
        DetailedHelpText = "### Format Numbers\n\nChoose standard comma, period, space, apostrophe, or Indian grouping and optionally reposition, replace, or remove a detected currency symbol. Right-click for direct US, German, French-space, or Swiss-apostrophe separator presets. The command affects the active text selection, selected shapes, selected slides, or the current slide. Other text and formatting remain in place, and successful changes create one Undo entry.",
        MinSelection = 0)]
    public static class FormatNumbersFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - opens dialog for interactive formatting.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var window = new oPenEfficiency.UI.FormatNumbersWindow(manager);
                var app = manager.GetApplication();
                if (app?.ActiveWindow != null)
                    new System.Windows.Interop.WindowInteropHelper(window).Owner = new IntPtr(app.HWND);
                window.Show();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FormatNumbersFeature.ExecuteDialog");
                throw;
            }
        }

        public static bool Execute(PowerPointManager manager, string targetThousandSep, string targetDecimalSep, string groupingType = "Standard", string currencyPlacement = "InPlace", string targetCurrency = "")
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;

                var selection = app.ActiveWindow.Selection;
                List<Shape> targetShapes = new List<Shape>();
                int changed = 0;

                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "FormatNumbersFeature.StartNewUndoEntry"); }

                if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    return ReformatTextNumbers(selection.TextRange, targetThousandSep, targetDecimalSep, groupingType, currencyPlacement, targetCurrency) > 0;
                }
                else if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    foreach (Shape s in selection.ShapeRange) GetAllShapesRecursive(s, targetShapes);
                }
                else if (selection.Type == PpSelectionType.ppSelectionSlides)
                {
                    for (int i = 1; i <= selection.SlideRange.Count; i++)
                    {
                        var slide = selection.SlideRange[i];
                        for (int j = 1; j <= slide.Shapes.Count; j++)
                            GetAllShapesRecursive(slide.Shapes[j], targetShapes);
                    }
                }
                else
                {
                    var slide = app.ActiveWindow.View.Slide as Slide;
                    if (slide != null)
                    {
                        for (int i = 1; i <= slide.Shapes.Count; i++)
                            GetAllShapesRecursive(slide.Shapes[i], targetShapes);
                    }
                }

                foreach (var shape in targetShapes)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        changed += ReformatTextNumbers(shape.TextFrame.TextRange, targetThousandSep, targetDecimalSep, groupingType, currencyPlacement, targetCurrency);
                    }
                    
                    if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        var table = shape.Table;
                        for (int r = 1; r <= table.Rows.Count; r++)
                        {
                            for (int c = 1; c <= table.Columns.Count; c++)
                            {
                                var cell = table.Cell(r, c);
                                if (cell.Shape.HasTextFrame == Office.MsoTriState.msoTrue && cell.Shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                {
                                    changed += ReformatTextNumbers(cell.Shape.TextFrame.TextRange, targetThousandSep, targetDecimalSep, groupingType, currencyPlacement, targetCurrency);
                                }
                            }
                        }
                    }
                }
                return changed > 0;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FormatNumbersFeature.Execute");
                return false;
            }
        }

        private static int ReformatTextNumbers(TextRange range, string tSep, string dSep, string groupingType, string currencyPlacement, string targetCurrency)
        {
            string text = range.Text;
            if (string.IsNullOrEmpty(text)) return 0;

            var regex = new System.Text.RegularExpressions.Regex(@"(?<![0-9\.,'])([$€£₹¥]?\s?-?\d+(?:[.,' ]\d{2,3})*(?:[.,]\d+)?\s?[$€£₹¥]?)(?![0-9\.,'])");
            var matches = regex.Matches(text);
            if (matches.Count == 0) return 0;

            int changed = 0;
            for (int m = matches.Count - 1; m >= 0; m--)
            {
                var match = matches[m];
                if (NumberTextFormatter.TryFormatToken(
                    match.Value,
                    tSep,
                    dSep,
                    groupingType,
                    currencyPlacement,
                    targetCurrency,
                    out string result))
                {
                    if (!string.Equals(result, match.Value, StringComparison.Ordinal))
                    {
                        range.Characters(match.Index + 1, match.Length).Text = result;
                        changed++;
                    }
                }
            }
            return changed;
        }

        private static bool TryNormalizeNumber(string input, out decimal result, out int decimalPlaces)
        {
            result = 0;
            decimalPlaces = 0;
            if (string.IsNullOrWhiteSpace(input)) return false;

            string normalized = input.Replace(" ", "").Replace("'", "");
            int lastComma = normalized.LastIndexOf(',');
            int lastDot = normalized.LastIndexOf('.');

            if (lastComma >= 0 && lastDot >= 0)
            {
                char decimalSeparator = lastComma > lastDot ? ',' : '.';
                char groupingSeparator = decimalSeparator == ',' ? '.' : ',';
                int decimalIndex = normalized.LastIndexOf(decimalSeparator);
                decimalPlaces = normalized.Length - decimalIndex - 1;
                normalized = normalized.Replace(groupingSeparator.ToString(), "");
                if (decimalSeparator == ',') normalized = normalized.Replace(',', '.');
            }
            else if (lastComma >= 0 || lastDot >= 0)
            {
                char separator = lastComma >= 0 ? ',' : '.';
                bool negative = normalized.StartsWith("-", StringComparison.Ordinal);
                string unsigned = negative ? normalized.Substring(1) : normalized;
                string[] groups = unsigned.Split(separator);
                bool grouping = IsGroupingPattern(groups);
                if (grouping)
                {
                    normalized = (negative ? "-" : "") + string.Concat(groups);
                }
                else
                {
                    decimalPlaces = groups[groups.Length - 1].Length;
                    normalized = (negative ? "-" : "") +
                        string.Concat(groups.Take(groups.Length - 1)) +
                        "." +
                        groups[groups.Length - 1];
                }
            }

            return decimal.TryParse(
                normalized,
                System.Globalization.NumberStyles.AllowLeadingSign |
                System.Globalization.NumberStyles.AllowDecimalPoint,
                System.Globalization.CultureInfo.InvariantCulture,
                out result);
        }

        private static bool IsGroupingPattern(string[] groups)
        {
            if (groups == null || groups.Length < 2 || groups.Any(group => group.Length == 0))
                return false;

            if (groups.Length == 2)
                return groups[0].Length >= 1 && groups[0].Length <= 3 && groups[1].Length == 3;

            bool standard = groups[0].Length >= 1 && groups[0].Length <= 3 &&
                groups.Skip(1).All(group => group.Length == 3);
            bool indian = groups[0].Length >= 1 && groups[0].Length <= 2 &&
                groups.Skip(1).Take(groups.Length - 2).All(group => group.Length == 2) &&
                groups[groups.Length - 1].Length == 3;
            return standard || indian;
        }

        private static string ApplyIndianGrouping(decimal num, int decimals, string tSep, string dSep)
        {
            string sign = num < 0 ? "-" : "";
            decimal absNum = Math.Abs(num);
            string[] parts = absNum.ToString("F" + decimals, System.Globalization.CultureInfo.InvariantCulture).Split('.');
            string integerPart = parts[0];
            string decimalPart = parts.Length > 1 ? parts[1] : "";

            string result = "";
            if (integerPart.Length <= 3) result = integerPart;
            else
            {
                string lastThree = integerPart.Substring(integerPart.Length - 3);
                string remaining = integerPart.Substring(0, integerPart.Length - 3);
                result = lastThree;
                while (remaining.Length > 0)
                {
                    int take = Math.Min(2, remaining.Length);
                    string chunk = remaining.Substring(remaining.Length - take);
                    result = chunk + tSep + result;
                    remaining = remaining.Substring(0, remaining.Length - take);
                }
            }
            return sign + result + (decimals > 0 ? dSep + decimalPart : "");
        }

        private static void GetAllShapesRecursive(Shape parent, List<Shape> list)
        {
            if (parent.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (Shape child in parent.GroupItems) GetAllShapesRecursive(child, list);
            }
            else list.Add(parent);
        }

    }
}
