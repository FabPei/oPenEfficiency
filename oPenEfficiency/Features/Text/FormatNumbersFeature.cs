using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Linq;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Reformat numbers within selected text, shapes, or slides.
    /// Supports thousand separators, decimal separators, and currency placement.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnFormatNumbers",
        Name = "Number Formatter",
        Tooltip = "Format numbers",
        IconData = "M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M19,19H5V5H19V19M7.5,13V15H9.5V13H7.5M12.5,11V15H14.5V11H12.5M17.5,9V15H19.5V9H17.5M7.5,7V11H9.5V7H7.5M12.5,5V9H14.5V5H12.5Z",
        Color = "#FBBF24",
        Description = "Reformats numbers with configurable thousand/decimal separators and currency symbols.",
        DetailedHelpText = "### Number Formatter\nFormats all numeric text in the selected shapes according to a configurable pattern, including decimal places, thousands separators, currency symbols, and percentage formatting.",
        MinSelection = 0)]
    public static class FormatNumbersFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - opens dialog for interactive formatting.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            new oPenEfficiency.UI.FormatNumbersWindow(manager).Show();
            return true;
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

                if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    ReformatTextNumbers(selection.TextRange, targetThousandSep, targetDecimalSep, groupingType, currencyPlacement, targetCurrency);
                    return true;
                }
                else if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    foreach (Shape s in selection.ShapeRange) GetAllShapesRecursive(s, targetShapes);
                }
                else
                {
                    var slides = GetTargetSlides(app, true).ToList();
                    foreach (Slide s in slides)
                        foreach (Shape sh in s.Shapes) GetAllShapesRecursive(sh, targetShapes);
                }

                foreach (var shape in targetShapes)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        ReformatTextNumbers(shape.TextFrame.TextRange, targetThousandSep, targetDecimalSep, groupingType, currencyPlacement, targetCurrency);
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
                                    ReformatTextNumbers(cell.Shape.TextFrame.TextRange, targetThousandSep, targetDecimalSep, groupingType, currencyPlacement, targetCurrency);
                                }
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FormatNumbersFeature.Execute");
                return false;
            }
        }

        private static void ReformatTextNumbers(TextRange range, string tSep, string dSep, string groupingType, string currencyPlacement, string targetCurrency)
        {
            string text = range.Text;
            if (string.IsNullOrEmpty(text)) return;

            var regex = new System.Text.RegularExpressions.Regex(@"(?<![0-9\.,'])([$€£₹¥]?\s?-?\d+(?:[.,' ]\d{2,3})*(?:[.,]\d+)?\s?[$€£₹¥]?)(?![0-9\.,'])");
            var matches = regex.Matches(text);
            if (matches.Count == 0) return;

            for (int m = matches.Count - 1; m >= 0; m--)
            {
                var match = matches[m];
                string val = match.Value;
                
                string originalCurrency = "";
                var curRegex = new System.Text.RegularExpressions.Regex(@"[$€£₹¥]");
                var curMatch = curRegex.Match(val);
                if (curMatch.Success) originalCurrency = curMatch.Value;
                
                string finalCurrency = !string.IsNullOrEmpty(targetCurrency) ? targetCurrency : originalCurrency;
                string numberPart = curRegex.Replace(val, "").Trim();

                if (TryNormalizeNumber(numberPart, out double num, out int decimals))
                {
                    string formattedNumber = "";
                    if (groupingType == "Indian") formattedNumber = ApplyIndianGrouping(num, decimals, tSep, dSep);
                    else
                    {
                        var nfi = new System.Globalization.NumberFormatInfo
                        {
                            NumberGroupSeparator = tSep,
                            NumberDecimalSeparator = dSep,
                            NumberGroupSizes = new[] { 3 }
                        };
                        formattedNumber = num.ToString("N" + decimals, nfi);
                    }

                    if (tSep == " ") formattedNumber = formattedNumber.Replace("\u00A0", " ");
                    if (decimals == 0 && formattedNumber.Contains(dSep))
                        formattedNumber = formattedNumber.Substring(0, formattedNumber.IndexOf(dSep));

                    string result = formattedNumber;
                    if (!string.IsNullOrEmpty(finalCurrency) && currencyPlacement != "Remove")
                    {
                        switch (currencyPlacement)
                        {
                            case "Prefix": result = finalCurrency + (finalCurrency.Length > 1 ? " " : "") + formattedNumber; break;
                            case "Suffix": result = formattedNumber + " " + finalCurrency; break;
                            case "InPlace":
                                if (!string.IsNullOrEmpty(dSep) && formattedNumber.Contains(dSep))
                                    result = formattedNumber.Replace(dSep, finalCurrency);
                                else result = formattedNumber + " " + finalCurrency;
                                break;
                        }
                    }
                    range.Characters(match.Index + 1, match.Length).Text = result;
                }
            }
        }

        private static bool TryNormalizeNumber(string input, out double result, out int decimalPlaces)
        {
            result = 0;
            decimalPlaces = 0;
            if (string.IsNullOrWhiteSpace(input)) return false;

            string normalized = input.Replace(" ", "").Replace("'", "");
            if (normalized.Contains(",") && normalized.Contains("."))
            {
                if (normalized.LastIndexOf(",") > normalized.LastIndexOf(".")) normalized = normalized.Replace(".", "");
                else normalized = normalized.Replace(",", "");
            }

            int lastComma = normalized.LastIndexOf(',');
            int lastDot = normalized.LastIndexOf('.');
            
            if (lastComma > lastDot) normalized = normalized.Replace(".", "").Replace(',', '.');
            else if (lastDot > lastComma) normalized = normalized.Replace(",", "");

            if (double.TryParse(normalized, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out result))
            {
                int dotIdx = normalized.IndexOf('.');
                if (dotIdx >= 0) decimalPlaces = normalized.Length - dotIdx - 1;
                return true;
            }
            return false;
        }

        private static string ApplyIndianGrouping(double num, int decimals, string tSep, string dSep)
        {
            string sign = num < 0 ? "-" : "";
            double absNum = Math.Abs(num);
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

        private static IEnumerable<Slide> GetTargetSlides(Application app, bool currentOnly)
        {
            if (currentOnly && app.ActiveWindow != null && app.ActiveWindow.View.Slide != null)
            {
                yield return (Slide)app.ActiveWindow.View.Slide;
            }
            else
            {
                foreach (Slide s in app.ActivePresentation.Slides) yield return s;
            }
        }
    }
}
