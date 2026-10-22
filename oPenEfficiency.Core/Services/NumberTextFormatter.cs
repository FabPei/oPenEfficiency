using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace oPenEfficiency.Services
{
    public static class NumberTextFormatter
    {
        private static readonly Regex CurrencyPattern = new Regex(@"[$€£₹¥]", RegexOptions.Compiled);

        public static bool TryFormatToken(
            string token,
            string thousandSeparator,
            string decimalSeparator,
            string groupingType,
            string currencyPlacement,
            string targetCurrency,
            out string formatted)
        {
            formatted = token;
            if (string.IsNullOrWhiteSpace(token)) return false;

            Match currencyMatch = CurrencyPattern.Match(token);
            string originalCurrency = currencyMatch.Success ? currencyMatch.Value : string.Empty;
            string finalCurrency = string.IsNullOrEmpty(targetCurrency) ? originalCurrency : targetCurrency;
            string numberPart = CurrencyPattern.Replace(token, string.Empty).Trim();
            if (!TryNormalizeNumber(numberPart, out decimal number, out int decimalPlaces)) return false;

            string formattedNumber;
            if (string.Equals(groupingType, "Indian", StringComparison.Ordinal))
            {
                formattedNumber = ApplyIndianGrouping(
                    number,
                    decimalPlaces,
                    thousandSeparator,
                    decimalSeparator);
            }
            else
            {
                var numberFormat = new NumberFormatInfo
                {
                    NumberGroupSeparator = thousandSeparator ?? string.Empty,
                    NumberDecimalSeparator = decimalSeparator ?? string.Empty,
                    NumberGroupSizes = new[] { 3 }
                };
                formattedNumber = number.ToString("N" + decimalPlaces, numberFormat);
            }

            if (thousandSeparator == " ")
                formattedNumber = formattedNumber.Replace("\u00A0", " ");
            if (decimalPlaces == 0 && !string.IsNullOrEmpty(decimalSeparator))
            {
                int decimalIndex = formattedNumber.IndexOf(decimalSeparator, StringComparison.Ordinal);
                if (decimalIndex >= 0) formattedNumber = formattedNumber.Substring(0, decimalIndex);
            }

            formatted = ApplyCurrency(
                formattedNumber,
                decimalSeparator,
                finalCurrency,
                currencyPlacement);
            return true;
        }

        private static string ApplyCurrency(
            string formattedNumber,
            string decimalSeparator,
            string currency,
            string placement)
        {
            if (string.IsNullOrEmpty(currency) || string.Equals(placement, "Remove", StringComparison.Ordinal))
                return formattedNumber;

            switch (placement)
            {
                case "Prefix":
                    return currency + (currency.Length > 1 ? " " : string.Empty) + formattedNumber;
                case "Suffix":
                    return formattedNumber + " " + currency;
                case "InPlace":
                    return !string.IsNullOrEmpty(decimalSeparator) && formattedNumber.Contains(decimalSeparator)
                        ? formattedNumber.Replace(decimalSeparator, currency)
                        : formattedNumber + " " + currency;
                default:
                    return formattedNumber;
            }
        }

        private static bool TryNormalizeNumber(string input, out decimal result, out int decimalPlaces)
        {
            result = 0;
            decimalPlaces = 0;
            if (string.IsNullOrWhiteSpace(input)) return false;

            string normalized = input.Replace(" ", string.Empty).Replace("'", string.Empty);
            int lastComma = normalized.LastIndexOf(',');
            int lastDot = normalized.LastIndexOf('.');

            if (lastComma >= 0 && lastDot >= 0)
            {
                char decimalMarker = lastComma > lastDot ? ',' : '.';
                char groupingMarker = decimalMarker == ',' ? '.' : ',';
                int decimalIndex = normalized.LastIndexOf(decimalMarker);
                decimalPlaces = normalized.Length - decimalIndex - 1;
                normalized = normalized.Replace(groupingMarker.ToString(), string.Empty);
                if (decimalMarker == ',') normalized = normalized.Replace(',', '.');
            }
            else if (lastComma >= 0 || lastDot >= 0)
            {
                char separator = lastComma >= 0 ? ',' : '.';
                bool negative = normalized.StartsWith("-", StringComparison.Ordinal);
                string unsigned = negative ? normalized.Substring(1) : normalized;
                string[] groups = unsigned.Split(separator);
                if (IsGroupingPattern(groups))
                {
                    normalized = (negative ? "-" : string.Empty) + string.Concat(groups);
                }
                else
                {
                    decimalPlaces = groups[groups.Length - 1].Length;
                    normalized = (negative ? "-" : string.Empty) +
                        string.Concat(groups.Take(groups.Length - 1)) + "." + groups[groups.Length - 1];
                }
            }

            return decimal.TryParse(
                normalized,
                NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture,
                out result);
        }

        private static bool IsGroupingPattern(string[] groups)
        {
            if (groups == null || groups.Length < 2 || groups.Any(group => group.Length == 0)) return false;
            if (groups.Length == 2)
                return groups[0].Length >= 1 && groups[0].Length <= 3 && groups[1].Length == 3;

            bool standard = groups[0].Length >= 1 && groups[0].Length <= 3 &&
                groups.Skip(1).All(group => group.Length == 3);
            bool indian = groups[0].Length >= 1 && groups[0].Length <= 2 &&
                groups.Skip(1).Take(groups.Length - 2).All(group => group.Length == 2) &&
                groups[groups.Length - 1].Length == 3;
            return standard || indian;
        }

        private static string ApplyIndianGrouping(
            decimal number,
            int decimals,
            string thousandSeparator,
            string decimalSeparator)
        {
            string sign = number < 0 ? "-" : string.Empty;
            string[] parts = Math.Abs(number)
                .ToString("F" + decimals, CultureInfo.InvariantCulture)
                .Split('.');
            string integerPart = parts[0];
            string decimalPart = parts.Length > 1 ? parts[1] : string.Empty;

            if (integerPart.Length > 3)
            {
                string lastThree = integerPart.Substring(integerPart.Length - 3);
                string prefix = integerPart.Substring(0, integerPart.Length - 3);
                var groups = new System.Collections.Generic.List<string>();
                while (prefix.Length > 2)
                {
                    groups.Insert(0, prefix.Substring(prefix.Length - 2));
                    prefix = prefix.Substring(0, prefix.Length - 2);
                }
                if (prefix.Length > 0) groups.Insert(0, prefix);
                integerPart = string.Join(thousandSeparator ?? string.Empty, groups) +
                    (groups.Count > 0 ? thousandSeparator : string.Empty) + lastThree;
            }

            return sign + integerPart +
                (decimals > 0 ? (decimalSeparator ?? string.Empty) + decimalPart : string.Empty);
        }
    }
}
