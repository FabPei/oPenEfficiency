using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace oPenEfficiency.Excel.Features.Data
{
    /// <summary>
    /// Pure, culture-agnostic validation of date-like text (no Excel dependency, so it is
    /// fully deterministic and testable). Detects strings that represent a date and decides
        /// whether they are a real calendar date; when not, it returns a human-readable reason
        /// in English such as "November 2026 only has 30 days".
    ///
    /// Two shapes are recognised:
    ///   • purely numeric  – three integer groups split by . - / (ISO "yyyy-mm-dd" is treated
    ///     strictly as year-month-day; otherwise the first two groups are order-ambiguous),
    ///   • with month name – exactly one month name (DE/EN/IT/ES/FR, accent-insensitive) plus
    ///     a day and a year.
    ///
    /// Where day/month order is ambiguous, a value counts as valid if *any* plausible
    /// interpretation is a real date; it is reported as invalid only when none works.
    /// </summary>
    public static class DateValidator
    {
        private static readonly Regex NumericDate =
            new Regex(@"^\s*(\d{1,4})[.\-/](\d{1,4})[.\-/](\d{1,4})\s*$", RegexOptions.Compiled);

        private static readonly char[] TokenSeparators = { ' ', '\t', '.', ',', '-', '/' };

        private static readonly string[] MonthNamesEn =
        {
            "", "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        };

        // Month names (DE/EN/IT/ES/FR), stored ASCII + lower-case; look-ups are accent-stripped.
        private static readonly Dictionary<string, int> Months = BuildMonths();

        /// <summary>True if the text is shaped like a date (numeric or with a month name), regardless of validity.</summary>
        public static bool LooksLikeDate(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            return NumericDate.IsMatch(text) || TryMonthNameDate(text, out _, out _, out _);
        }

        /// <summary>True for text that is shaped like a date but is not a real calendar date.</summary>
        public static bool IsImpossibleDate(string text) => GetInvalidReason(text) != null;

        /// <summary>
        /// Returns an English reason why the text is not a real date, or null if the text is a
        /// valid date or is not a date at all.
        /// </summary>
        public static string GetInvalidReason(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;

            var m = NumericDate.Match(text);
            if (m.Success) return NumericReason(m);

            if (TryMonthNameDate(text, out int month, out int n1, out int n2))
                return MonthNameReason(month, n1, n2);

            return null;
        }

        private static string NumericReason(Match m)
        {
            int a = int.Parse(m.Groups[1].Value);
            int b = int.Parse(m.Groups[2].Value);
            int c = int.Parse(m.Groups[3].Value);

            // ISO: a four-digit first group fixes the order to year-month-day.
            if (m.Groups[1].Value.Length == 4)
                return ReasonForDate(c, b, a);

            // Otherwise the last group is the year; the first two are order-ambiguous.
            int year = ExpandTwoDigitYear(c);
            string r1 = ReasonForDate(a, b, year); // day=a, month=b
            string r2 = ReasonForDate(b, a, year); // day=b, month=a
            if (r1 == null || r2 == null) return null; // a valid reading exists

            // Prefer the interpretation with a valid month (its reason is about the day).
            if (InMonthRange(b)) return r1;
            if (InMonthRange(a)) return r2;
            return r1;
        }

        private static string MonthNameReason(int month, int n1, int n2)
        {
            // month is known; the two numbers are day and year in unknown order.
            string rA = ReasonForDate(n1, month, ExpandTwoDigitYear(n2)); // day=n1, year=n2
            string rB = ReasonForDate(n2, month, ExpandTwoDigitYear(n1)); // day=n2, year=n1
            if (rA == null || rB == null) return null;

            // Prefer the interpretation whose year looks like a year.
            if (n2 > 31) return rA;
            if (n1 > 31) return rB;
            return rA;
        }

        /// <summary>Returns a reason if (day, month, year) is not a real date, else null.</summary>
        private static string ReasonForDate(int day, int month, int year)
        {
            if (month < 1 || month > 12) return $"Month {month} does not exist (valid range is 1 to 12).";
            if (year < 1 || year > 9999) return $"Year {year} is invalid.";
            if (day < 1) return $"Day {day} is invalid.";
            if (day > 31) return $"Day {day} does not exist (no month has more than 31 days).";
            int dim = DateTime.DaysInMonth(year, month);
            if (day > dim) return $"{MonthNamesEn[month]} {year} only has {dim} days (day {day} does not exist).";
            return null;
        }

        /// <summary>
        /// Recognises a date written with a month name. The whole string must consist of
        /// exactly one month name and two numbers (day + year); any other word rejects it,
        /// so prose like "top 30 in Feb 2026" is not treated as a date.
        /// </summary>
        private static bool TryMonthNameDate(string text, out int month, out int num1, out int num2)
        {
            month = num1 = num2 = 0;
            var tokens = text.Split(TokenSeparators, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length != 3) return false;

            int monthCount = 0;
            var numbers = new List<int>(2);
            foreach (var token in tokens)
            {
                if (Months.TryGetValue(Normalize(token), out int mnum)) { month = mnum; monthCount++; }
                else if (int.TryParse(token, out int n)) numbers.Add(n);
                else return false; // unknown alphabetic token -> not a plain date
            }

            if (monthCount != 1 || numbers.Count != 2) return false;
            num1 = numbers[0];
            num2 = numbers[1];
            return true;
        }

        private static bool InMonthRange(int month) => month >= 1 && month <= 12;

        private static int ExpandTwoDigitYear(int year) => year < 100 ? 2000 + year : year;

        /// <summary>Lower-cases and strips diacritics so "Février" and "fevrier" compare equal.</summary>
        private static string Normalize(string token)
        {
            if (string.IsNullOrEmpty(token)) return string.Empty;
            string decomposed = token.ToLowerInvariant().Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(decomposed.Length);
            foreach (char ch in decomposed)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
                    sb.Append(ch);
            }
            return sb.ToString().Normalize(NormalizationForm.FormC);
        }

        private static Dictionary<string, int> BuildMonths()
        {
            var m = new Dictionary<string, int>(StringComparer.Ordinal);
            void Add(int number, params string[] names) { foreach (var n in names) m[n] = number; }

            // German (umlauts written accent-free, since look-ups are normalized: ä -> a)
            Add(1, "januar", "janner", "jan");
            Add(2, "februar", "feb");
            Add(3, "marz", "maerz", "mar", "mrz");
            Add(4, "april", "apr");
            Add(5, "mai");
            Add(6, "juni", "jun");
            Add(7, "juli", "jul");
            Add(8, "august", "aug");
            Add(9, "september", "sep", "sept");
            Add(10, "oktober", "okt");
            Add(11, "november", "nov");
            Add(12, "dezember", "dez");

            // English
            Add(1, "january");
            Add(2, "february");
            Add(3, "march");
            Add(5, "may");
            Add(6, "june");
            Add(7, "july");
            Add(10, "oct", "october");
            Add(12, "dec", "december");

            // Italian
            Add(1, "gennaio", "gen");
            Add(2, "febbraio");
            Add(3, "marzo");
            Add(4, "aprile");
            Add(5, "maggio", "mag");
            Add(6, "giugno", "giu");
            Add(7, "luglio", "lug");
            Add(8, "agosto", "ago");
            Add(9, "settembre", "set");
            Add(10, "ottobre", "ott");
            Add(11, "novembre");
            Add(12, "dicembre", "dic");

            // Spanish
            Add(1, "enero", "ene");
            Add(2, "febrero");
            Add(4, "abril", "abr");
            Add(5, "mayo");
            Add(6, "junio");
            Add(7, "julio");
            Add(9, "septiembre", "setiembre");
            Add(10, "octubre");
            Add(11, "noviembre");
            Add(12, "diciembre");

            // French (accents stripped by Normalize: février -> fevrier, août -> aout)
            Add(1, "janvier", "janv");
            Add(2, "fevrier", "fevr");
            Add(3, "mars");
            Add(4, "avril", "avr");
            Add(6, "juin");
            Add(7, "juillet", "juil");
            Add(8, "aout");
            Add(9, "septembre");
            Add(10, "octobre");
            Add(11, "novembre");
            Add(12, "decembre");
            return m;
        }
    }
}
