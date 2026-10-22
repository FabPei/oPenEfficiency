using System;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using oPenEfficiency.Excel.Services;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;
using Xl = global::Microsoft.Office.Interop.Excel;

namespace oPenEfficiency.Excel.Features.Rounding
{
    /// <summary>
    /// Ribbon-driven actions for consistent rounding: wrapping selected cells with the
    /// OEROUND family in place, removing them again, highlighting "cheated" cells,
    /// and registering the automation add-in that provides the worksheet functions.
    /// </summary>
    public static class OeRoundActions
    {
        private const uint HighlightColor = 0x0000FFFF; // yellow (BGR)

        private static readonly Regex WrapPattern =
            new Regex(@"^\s*OEROUND(?:UP|DOWN|NEAR)?\s*\((.*),[^,()]*\)\s*$",
                      RegexOptions.IgnoreCase | RegexOptions.Singleline);

        /// <summary>Wraps every cell in the selection with <paramref name="func"/>(body, n) in place.</summary>
        public static bool Wrap(ExcelManager manager, string func, double n)
        {
            var range = manager?.SelectedRange;
            var app = manager?.Application;
            if (range == null || app == null) return false;

            string nLiteral = n.ToString(CultureInfo.InvariantCulture);

            foreach (Xl.Range cell in range.Cells)
            {
                try
                {
                    string body = GetInnerBody(cell);
                    if (string.IsNullOrEmpty(body)) continue;
                    cell.Formula = "=" + func + "(" + body + "," + nLiteral + ")";
                }
                catch (Exception ex) { ExceptionLogger.Log(ex, "OeRoundActions.Wrap.Cell"); }
            }

            CreateGroup(app, range);
            RunOptionalExcel(() => app.Calculate(), "OeRoundActions.Wrap.Calculate");
            return true;
        }

        /// <summary>Removes OEROUND wrappers from the selection, restoring the inner formula/value.</summary>
        public static bool Remove(ExcelManager manager)
        {
            var range = manager?.SelectedRange;
            var app = manager?.Application;
            if (range == null || app == null) return false;

            foreach (Xl.Range cell in range.Cells)
            {
                try
                {
                    string formula = cell.Formula as string;
                    if (string.IsNullOrEmpty(formula)) continue;
                    string expr = formula.StartsWith("=") ? formula.Substring(1) : formula;
                    var match = WrapPattern.Match(expr);
                    if (match.Success)
                    {
                        cell.Formula = "=" + match.Groups[1].Value;
                    }
                }
                catch (Exception ex) { ExceptionLogger.Log(ex, "OeRoundActions.Remove.Cell"); }
            }

            RemoveGroupsIntersecting(app, range);
            RunOptionalExcel(() => app.Calculate(), "OeRoundActions.Remove.Calculate");
            return true;
        }

        /// <summary>Highlights every cell that was rounded away from its nearest multiple.</summary>
        public static bool Highlight(ExcelManager manager)
        {
            var app = manager?.Application;
            var wb = app?.ActiveWorkbook;
            if (wb == null) return false;

            int marked = 0;
            foreach (Xl.Name name in wb.Names)
            {
                string nm = null;
                TryReadOptionalExcel(() => name.Name, out nm, "OeRoundActions.Highlight.Name");
                if (nm == null || !nm.Contains(OeRoundEngine.GroupNamePrefix)) continue;

                Xl.Range range = null;
                TryReadOptionalExcel(() => name.RefersToRange, out range, "OeRoundActions.Highlight.Range");
                if (range == null) continue;

                foreach (Xl.Range cell in range.Cells)
                {
                    try
                    {
                        if (OeRoundEngine.IsCheated(cell))
                        {
                            cell.Interior.Color = HighlightColor;
                            marked++;
                        }
                    }
                    catch (Exception ex) { ReportOptionalExcelFailure(ex, "OeRoundActions.Highlight.Cell"); }
                }
            }
            return marked > 0;
        }

        private static string GetInnerBody(Xl.Range cell)
        {
            try
            {
                object formula = cell.Formula;
                string s = formula as string;
                if (s == null)
                {
                    object v = cell.Value2;
                    return v == null ? null : Convert.ToString(v, CultureInfo.InvariantCulture);
                }
                if (s.StartsWith("=")) s = s.Substring(1);
                if (string.IsNullOrWhiteSpace(s)) return null;

                // Avoid nesting: if it is already OEROUND-wrapped, unwrap first.
                var match = WrapPattern.Match(s);
                return match.Success ? match.Groups[1].Value : s;
            }
            catch
            {
                return null;
            }
        }

        private static void CreateGroup(Xl.Application app, Xl.Range range)
        {
            try
            {
                var wb = app.ActiveWorkbook;
                string name = OeRoundEngine.GroupNamePrefix + "N" + Guid.NewGuid().ToString("N");
                wb.Names.Add(name, range);
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "OeRoundActions.CreateGroup"); }
        }

        private static void RemoveGroupsIntersecting(Xl.Application app, Xl.Range range)
        {
            try
            {
                var wb = app.ActiveWorkbook;
                for (int i = wb.Names.Count; i >= 1; i--)
                {
                    Xl.Name name = wb.Names.Item(i);
                    string nm = null;
                    TryReadOptionalExcel(() => name.Name, out nm, "OeRoundActions.RemoveGroups.Name");
                    if (nm == null || !nm.Contains(OeRoundEngine.GroupNamePrefix)) continue;

                    Xl.Range groupRange = null;
                    TryReadOptionalExcel(() => name.RefersToRange, out groupRange, "OeRoundActions.RemoveGroups.Range");
                    if (groupRange == null)
                    {
                        RunOptionalExcel(() => name.Delete(), "OeRoundActions.RemoveGroups.DeleteInvalidName");
                        continue;
                    }

                    Xl.Range overlap = null;
                    TryReadOptionalExcel(() => app.Intersect(groupRange, range), out overlap, "OeRoundActions.RemoveGroups.Intersect");
                    if (overlap != null)
                    {
                        RunOptionalExcel(() => name.Delete(), "OeRoundActions.RemoveGroups.DeleteName");
                    }
                }
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "OeRoundActions.RemoveGroupsIntersecting"); }
        }

        private static bool TryReadOptionalExcel<T>(Func<T> operation, out T value, string context)
        {
            try
            {
                value = operation();
                return true;
            }
            catch (Exception ex)
            {
                value = default(T);
                ReportOptionalExcelFailure(ex, context);
                return false;
            }
        }

        private static void RunOptionalExcel(Action operation, string context)
        {
            try
            {
                operation();
            }
            catch (Exception ex)
            {
                ReportOptionalExcelFailure(ex, context);
            }
        }

        private static void ReportOptionalExcelFailure(Exception exception, string context)
        {
            if (!ExceptionPolicy.IsExpectedComFailure(exception))
            {
                ExceptionLogger.Log(exception, context);
            }
            // A stale workbook/name/cell proxy must not abort processing of the remaining selection.
        }

        // ---- Automation add-in registration (per user) ----

        /// <summary>
        /// Registers the OEROUND function library for the current user and adds it to Excel's
        /// automation add-ins. Excel usually needs a restart before the functions become live.
        /// </summary>
        public static void RegisterForCurrentUser(Xl.Application app)
        {
            try
            {
                var asm = typeof(OeRoundFunctions).Assembly;
                var type = typeof(OeRoundFunctions);
                string guid = type.GUID.ToString("B").ToUpperInvariant();
                string codeBase = new Uri(asm.Location).AbsoluteUri;
                string version = asm.GetName().Version.ToString();

                string clsidPath = $@"Software\Classes\CLSID\{guid}";
                using (var clsid = Registry.CurrentUser.CreateSubKey(clsidPath))
                {
                    clsid.SetValue(null, "oPenEfficiency OEROUND");
                    using (var ips = clsid.CreateSubKey("InprocServer32"))
                    {
                        ips.SetValue(null, "mscoree.dll");
                        ips.SetValue("ThreadingModel", "Both");
                        ips.SetValue("Class", type.FullName);
                        ips.SetValue("Assembly", asm.FullName);
                        ips.SetValue("RuntimeVersion", asm.ImageRuntimeVersion);
                        ips.SetValue("CodeBase", codeBase);
                        using (var ver = ips.CreateSubKey(version))
                        {
                            ver.SetValue("Class", type.FullName);
                            ver.SetValue("Assembly", asm.FullName);
                            ver.SetValue("RuntimeVersion", asm.ImageRuntimeVersion);
                            ver.SetValue("CodeBase", codeBase);
                        }
                    }
                    using (var prog = clsid.CreateSubKey("ProgId")) prog.SetValue(null, OeRoundFunctions.ProgId);
                    clsid.CreateSubKey("Programmable");
                }

                using (var prog = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{OeRoundFunctions.ProgId}"))
                {
                    prog.SetValue(null, "oPenEfficiency OEROUND");
                    using (var c = prog.CreateSubKey("CLSID")) c.SetValue(null, guid);
                }

                AddAutomationAddIn(app);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "OeRoundActions.RegisterForCurrentUser");
            }
        }

        private static void AddAutomationAddIn(Xl.Application app)
        {
            try
            {
                string version = app.Version; // e.g. "16.0"
                string optionsPath = $@"Software\Microsoft\Office\{version}\Excel\Options";
                string entry = "/A \"" + OeRoundFunctions.ProgId + "\"";

                using (var options = Registry.CurrentUser.CreateSubKey(optionsPath))
                {
                    // Already present?
                    foreach (var valueName in options.GetValueNames())
                    {
                        if (!valueName.StartsWith("OPEN", StringComparison.OrdinalIgnoreCase)) continue;
                        if (string.Equals(options.GetValue(valueName) as string, entry, StringComparison.OrdinalIgnoreCase))
                            return;
                    }

                    // Find the next free OPEN / OPENn slot.
                    string name = "OPEN";
                    int i = 0;
                    while (options.GetValue(name) != null)
                    {
                        i++;
                        name = "OPEN" + i;
                    }
                    options.SetValue(name, entry);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "OeRoundActions.AddAutomationAddIn");
            }
        }
    }

    [FeatureMetadata(
        Id = "BtnOeRoundFormula",
        Name = "Formula Rounding",
        Tooltip = "Wrap selection with OEROUND formula",
        Color = "#10B981",
        Description = "Wraps selected cell formulas with consistent OEROUND rounding in place.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Cells)]
    public static class OeRoundFormulaFeature
    {
        public static bool Execute(ExcelManager manager) => OeRoundActions.Wrap(manager, "OEROUND", 1);
    }

    [FeatureMetadata(
        Id = "BtnOeRoundHighlight",
        Name = "Highlight Cheated Cells",
        Tooltip = "Highlight cells rounded away from nearest multiple",
        Color = "#F59E0B",
        Description = "Highlights cells in the workbook that were rounded away from their nearest multiple.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class OeRoundHighlightFeature
    {
        public static bool Execute(ExcelManager manager) => OeRoundActions.Highlight(manager);
    }
}
