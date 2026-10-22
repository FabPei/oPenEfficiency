using System;
using System.Globalization;
using System.Linq;
using System.Text;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free state contract for an Excel cell value link (QA-S9 external-data wave): the
    /// newline-delimited "key=value" AlternativeText codec written below the oPE_ExcelValue
    /// tag line. Numbers are invariant; unparsable fields keep their documented defaults
    /// (decimals/start/length -1, raw null, booleans only on exact "1"). Production
    /// implementation used by ExcelValueLink.
    /// </summary>
    public static class ExcelValueLinkStatePolicy
    {
        public const string Tag = "oPE_ExcelValue";

        public static void Defaults(
            out int decimals, out bool thousands, out bool percent,
            out double? raw, out string id, out int start, out int length, out string text)
        {
            decimals = -1;   // -1 = show Excel's displayed text as-is
            thousands = false;
            percent = false;
            raw = null;
            id = null;
            start = -1;      // 1-based text start inside a text box for inline links
            length = -1;     // current text length for inline links
            text = null;
        }

        public static string BuildTag(
            string path, string sheet, string cell,
            int decimals, bool thousands, bool percent,
            double? raw, string id, int start, int length, string text)
        {
            var sb = new StringBuilder();
            sb.Append(Tag).Append('\n');
            sb.Append("path=").Append(path ?? "").Append('\n');
            sb.Append("sheet=").Append(sheet ?? "").Append('\n');
            sb.Append("cell=").Append(cell ?? "").Append('\n');
            sb.Append("dec=").Append(decimals).Append('\n');
            sb.Append("thou=").Append(thousands ? 1 : 0).Append('\n');
            sb.Append("pct=").Append(percent ? 1 : 0).Append('\n');
            sb.Append("raw=").Append(raw?.ToString(CultureInfo.InvariantCulture) ?? "").Append('\n');
            sb.Append("id=").Append(id ?? "").Append('\n');
            sb.Append("start=").Append(start).Append('\n');
            sb.Append("len=").Append(length).Append('\n');
            sb.Append("text=").Append(text ?? "");
            return sb.ToString();
        }

        public static bool TryParseTag(
            string altText,
            out string path, out string sheet, out string cell,
            out int decimals, out bool thousands, out bool percent,
            out double? raw, out string id, out int start, out int length, out string text)
        {
            path = null; sheet = null; cell = null;
            Defaults(out decimals, out thousands, out percent, out raw, out id, out start, out length, out text);
            if (string.IsNullOrEmpty(altText)) return false;

            string[] lines = altText.Replace("\r", "").Split('\n');
            if (lines.Length == 0 || lines[0].Trim() != Tag) return false;

            foreach (string line in lines.Skip(1))
            {
                int i = line.IndexOf('=');
                if (i < 0) continue;
                string key = line.Substring(0, i);
                string value = line.Substring(i + 1);
                switch (key)
                {
                    case "path": path = value; break;
                    case "sheet": sheet = value; break;
                    case "cell": cell = value; break;
                    case "dec": if (int.TryParse(value, out int d)) decimals = d; break;
                    case "thou": thousands = value == "1"; break;
                    case "pct": percent = value == "1"; break;
                    case "raw":
                        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double r))
                            raw = r;
                        break;
                    case "id": id = value; break;
                    case "start": if (int.TryParse(value, out int s)) start = s; break;
                    case "len": if (int.TryParse(value, out int l)) length = l; break;
                    case "text": text = value; break;
                }
            }
            return true;
        }
    }
}
