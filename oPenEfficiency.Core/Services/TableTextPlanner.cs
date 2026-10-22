using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace oPenEfficiency.Services
{
    public sealed class TableCellTotal
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public double Value { get; set; }
    }

    public static class FlexibleNumberParser
    {
        public static bool TryParse(string text, out double result)
        {
            result = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;

            string cleaned = text.Trim()
                .Replace("\u00A0", string.Empty)
                .Replace("\u202F", string.Empty)
                .Replace(" ", string.Empty);
            if (cleaned.EndsWith("%", StringComparison.Ordinal))
                cleaned = cleaned.Substring(0, cleaned.Length - 1).Trim();
            cleaned = new string(cleaned
                .Where(character => CharUnicodeInfo.GetUnicodeCategory(character) != UnicodeCategory.CurrencySymbol)
                .ToArray());

            bool parenthesized = cleaned.Length > 1 && cleaned[0] == '(' && cleaned[cleaned.Length - 1] == ')';
            if (parenthesized) cleaned = cleaned.Substring(1, cleaned.Length - 2);
            bool trailingNegative = cleaned.EndsWith("-", StringComparison.Ordinal);
            if (trailingNegative) cleaned = "-" + cleaned.Substring(0, cleaned.Length - 1);

            int lastComma = cleaned.LastIndexOf(',');
            int lastDot = cleaned.LastIndexOf('.');
            if (lastComma >= 0 && lastDot >= 0)
            {
                char decimalSeparator = lastComma > lastDot ? ',' : '.';
                char groupingSeparator = decimalSeparator == ',' ? '.' : ',';
                cleaned = cleaned.Replace(groupingSeparator.ToString(), string.Empty);
                if (decimalSeparator == ',') cleaned = cleaned.Replace(',', '.');
            }
            else if (lastComma >= 0 || lastDot >= 0)
            {
                char separator = lastComma >= 0 ? ',' : '.';
                bool negative = cleaned.StartsWith("-", StringComparison.Ordinal);
                string unsigned = negative ? cleaned.Substring(1) : cleaned;
                string[] groups = unsigned.Split(separator);
                bool grouping = groups.Length > 1 &&
                    groups[0].Length >= 1 && groups[0].Length <= 3 &&
                    groups.Skip(1).All(group => group.Length == 3);
                if (grouping && groups.Length == 2 &&
                    CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator == separator.ToString())
                {
                    grouping = false;
                }
                cleaned = grouping
                    ? (negative ? "-" : string.Empty) + string.Concat(groups)
                    : (negative ? "-" : string.Empty) +
                        string.Concat(groups.Take(groups.Length - 1)) + "." + groups[groups.Length - 1];
            }

            if (parenthesized && !cleaned.StartsWith("-", StringComparison.Ordinal)) cleaned = "-" + cleaned;
            return double.TryParse(cleaned, NumberStyles.Float, CultureInfo.InvariantCulture, out result) &&
                !double.IsNaN(result) &&
                !double.IsInfinity(result);
        }
    }

    public static class TableTextPlanner
    {
        public static IReadOnlyList<string[]> SortRows(
            IReadOnlyList<string[]> rows,
            int sortColumn,
            bool ascending)
        {
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            if (rows.Count == 0) return Array.Empty<string[]>();
            int columnCount = rows[0]?.Length ?? 0;
            if (columnCount == 0 || rows.Any(row => row == null || row.Length != columnCount))
                throw new ArgumentException("Rows must have a consistent positive column count.", nameof(rows));
            if (sortColumn < 0 || sortColumn >= columnCount)
                throw new ArgumentOutOfRangeException(nameof(sortColumn));

            return rows
                .Select((cells, index) => new SortableRow
                {
                    Cells = (string[])cells.Clone(),
                    OriginalIndex = index
                })
                .OrderBy(row => row, new RowComparer(sortColumn, ascending))
                .Select(row => row.Cells)
                .ToList();
        }

        public static IReadOnlyList<TableCellTotal> CalculateTotals(
            string[,] cells,
            bool sumRows,
            bool sumColumns,
            int firstRowNumber = 1,
            int firstColumnNumber = 1)
        {
            if (cells == null) throw new ArgumentNullException(nameof(cells));
            int rows = cells.GetLength(0);
            int columns = cells.GetLength(1);
            bool calculateColumns = sumColumns && rows > 1;
            bool calculateRows = sumRows && columns > 1;
            if (!calculateColumns && !calculateRows) return Array.Empty<TableCellTotal>();

            var totals = new List<TableCellTotal>();
            if (calculateColumns && calculateRows)
            {
                for (int column = 0; column < columns - 1; column++)
                    totals.Add(Total(rows - 1, column, Sum(cells, 0, rows - 2, column, column, firstRowNumber, firstColumnNumber)));
                for (int row = 0; row < rows - 1; row++)
                    totals.Add(Total(row, columns - 1, Sum(cells, row, row, 0, columns - 2, firstRowNumber, firstColumnNumber)));
                totals.Add(Total(rows - 1, columns - 1, Sum(cells, 0, rows - 2, 0, columns - 2, firstRowNumber, firstColumnNumber)));
            }
            else if (calculateColumns)
            {
                for (int column = 0; column < columns; column++)
                    totals.Add(Total(rows - 1, column, Sum(cells, 0, rows - 2, column, column, firstRowNumber, firstColumnNumber)));
            }
            else
            {
                for (int row = 0; row < rows; row++)
                    totals.Add(Total(row, columns - 1, Sum(cells, row, row, 0, columns - 2, firstRowNumber, firstColumnNumber)));
            }
            return totals;
        }

        private static double Sum(
            string[,] cells,
            int firstRow,
            int lastRow,
            int firstColumn,
            int lastColumn,
            int rowOffset,
            int columnOffset)
        {
            double sum = 0;
            for (int row = firstRow; row <= lastRow; row++)
            {
                for (int column = firstColumn; column <= lastColumn; column++)
                {
                    string text = cells[row, column];
                    if (string.IsNullOrWhiteSpace(text)) continue;
                    if (!FlexibleNumberParser.TryParse(text, out double value))
                    {
                        throw new FormatException(
                            $"Cell {ColumnName(columnOffset + column)}{rowOffset + row} contains a non-numeric value ('{text.Trim()}'). No totals were changed.");
                    }
                    sum += value;
                }
            }
            return sum;
        }

        private static TableCellTotal Total(int row, int column, double value)
            => new TableCellTotal { Row = row, Column = column, Value = value };

        private static string ColumnName(int column)
        {
            string result = string.Empty;
            while (column > 0)
            {
                column--;
                result = (char)('A' + column % 26) + result;
                column /= 26;
            }
            return result;
        }

        private sealed class SortableRow
        {
            public string[] Cells { get; set; }
            public int OriginalIndex { get; set; }
        }

        private sealed class RowComparer : IComparer<SortableRow>
        {
            private readonly int _column;
            private readonly bool _ascending;

            public RowComparer(int column, bool ascending)
            {
                _column = column;
                _ascending = ascending;
            }

            public int Compare(SortableRow first, SortableRow second)
            {
                string firstValue = first.Cells[_column];
                string secondValue = second.Cells[_column];
                bool firstNumeric = FlexibleNumberParser.TryParse(firstValue, out double firstNumber);
                bool secondNumeric = FlexibleNumberParser.TryParse(secondValue, out double secondNumber);
                int comparison;
                if (firstNumeric && secondNumeric) comparison = firstNumber.CompareTo(secondNumber);
                else if (firstNumeric != secondNumeric) comparison = firstNumeric ? -1 : 1;
                else comparison = CultureInfo.CurrentCulture.CompareInfo.Compare(
                    firstValue ?? string.Empty,
                    secondValue ?? string.Empty,
                    CompareOptions.IgnoreCase);
                if (!_ascending) comparison = -comparison;
                return comparison != 0 ? comparison : first.OriginalIndex.CompareTo(second.OriginalIndex);
            }
        }
    }
}
