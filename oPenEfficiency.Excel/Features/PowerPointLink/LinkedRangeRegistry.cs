using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using WinForms = System.Windows.Forms;
using Xl = Microsoft.Office.Interop.Excel;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Excel.Features.PowerPointLink
{
    /// <summary>
    /// Persistent, format-neutral registry of ranges linked to PowerPoint tables.
    /// Hidden workbook names are used because their references follow row/column moves.
    /// </summary>
    internal static class LinkedRangeRegistry
    {
        private const string NamePrefix = "_oPE_PptTableLink_";
        private const string HighlightFormulaMarker = "oPE_LINK_HIGHLIGHT_";
        private static readonly int[] HighlightColors =
        {
            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 235, 59)),
            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(129, 212, 250)),
            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(165, 214, 167)),
            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 183, 77)),
            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(206, 147, 216)),
            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(239, 154, 154))
        };
        private static readonly string[] HighlightColorNames = { "Yellow", "Blue", "Green", "Orange", "Purple", "Red" };

        public static bool Register(Xl.Workbook workbook, Xl.Range range, bool formatted, out string linkId)
        {
            linkId = null;
            if (workbook == null || range == null) return false;
            try
            {
                linkId = Guid.NewGuid().ToString("N");
                string name = NamePrefix + linkId;
                string sheet = ((Xl.Worksheet)range.Worksheet).Name.Replace("'", "''");
                string address = range.Address[true, true, Xl.XlReferenceStyle.xlA1];
                // The third argument is Visible. Passing the flag in the fifth slot would
                // target ShortcutKey and makes Excel reject it with "must be a letter".
                Xl.Name entry = workbook.Names.Add(
                    name, "='" + sheet + "'!" + address, false);
                entry.Visible = false;
                try { entry.Comment = "oPenEfficiency PowerPoint table link|" + (formatted ? "F" : "V") + "|" + DateTime.UtcNow.ToString("o"); }
                catch (COMException)
                {
                    // The link remains valid when this optional diagnostic comment is unavailable.
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LinkedRangeRegistry.Register");
                linkId = null;
                return false;
            }
        }

        public static void Unregister(Xl.Workbook workbook, string linkId)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(linkId)) return;

            try
            {
                workbook.Names.Item(NamePrefix + linkId).Delete();
            }
            catch (COMException)
            {
                // Registration may already have rolled back or the name may not exist.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LinkedRangeRegistry.Unregister");
            }
        }

        public static string GetSelectionStatus(Xl.Application app)
        {
            try
            {
                Xl.Workbook workbook = app?.ActiveWorkbook;
                Xl.Range selection = app?.Selection as Xl.Range;
                if (workbook == null || selection == null) return "No linked range selected";

                int matches = 0;
                foreach (Xl.Range linked in GetValidRanges(workbook, false))
                {
                    Xl.Range intersection = null;
                    try
                    {
                        intersection = app.Intersect(selection, linked);
                        if (intersection != null) matches++;
                    }
                    catch (COMException)
                    {
                        // Intersect can fail for ranges from incompatible workbook contexts.
                    }
                    finally { Release(intersection); Release(linked); }
                }
                return matches == 0 ? "No linked range selected" :
                    matches == 1 ? "Selection is linked to PowerPoint" :
                    "Selection contains " + matches + " PowerPoint links";
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LinkedRangeRegistry.GetSelectionStatus");
                return "Link status unavailable";
            }
        }

        public static void ShowLinkedRanges(Xl.Application app)
        {
            try
            {
                Xl.Workbook workbook = app?.ActiveWorkbook;
                if (workbook == null) return;
                var ranges = GetValidRanges(workbook, true);
                if (ranges.Count == 0)
                {
                    WinForms.MessageBox.Show("This workbook has no linked PowerPoint table ranges.", "Linked Ranges");
                    return;
                }

                using (var dialog = new WinForms.Form { Text = "Linked PowerPoint Ranges", Width = 500, Height = 330, StartPosition = WinForms.FormStartPosition.CenterScreen })
                using (var list = new WinForms.ListBox { Dock = WinForms.DockStyle.Fill })
                using (var go = new WinForms.Button { Text = "Select Range", Dock = WinForms.DockStyle.Bottom, Height = 36 })
                {
                    for (int i = 0; i < ranges.Count; i++)
                    {
                        Xl.Range range = ranges[i];
                        string sheet = ((Xl.Worksheet)range.Worksheet).Name;
                        list.Items.Add("[" + HighlightColorNames[i % HighlightColorNames.Length] + "]  " +
                            sheet + "!" + range.Address[false, false, Xl.XlReferenceStyle.xlA1]);
                    }
                    if (list.Items.Count > 0) list.SelectedIndex = 0;
                    go.Click += (s, e) => dialog.DialogResult = WinForms.DialogResult.OK;
                    list.DoubleClick += (s, e) => dialog.DialogResult = WinForms.DialogResult.OK;
                    dialog.Controls.Add(list);
                    dialog.Controls.Add(go);
                    if (dialog.ShowDialog() == WinForms.DialogResult.OK && list.SelectedIndex >= 0)
                    {
                        Xl.Range target = ranges[list.SelectedIndex];
                        ((Xl.Worksheet)target.Worksheet).Activate();
                        target.Select();
                    }
                }
                foreach (Xl.Range range in ranges) Release(range);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LinkedRangeRegistry.ShowLinkedRanges");
                WinForms.MessageBox.Show("Linked ranges could not be read.", "Linked Ranges");
            }
        }

        /// <summary>
        /// Adds or removes temporary, identifiable conditional-format rules. These rules
        /// visually color linked ranges without changing their underlying fill formatting.
        /// </summary>
        public static void SetHighlightsVisible(Xl.Application app, bool visible)
        {
            Xl.Workbook workbook = app?.ActiveWorkbook;
            if (workbook == null) return;
            try
            {
                RemoveHighlightRules(workbook);
                if (!visible) return;

                var ranges = GetValidRanges(workbook, true);
                for (int i = 0; i < ranges.Count; i++)
                {
                    Xl.Range range = ranges[i];
                    try
                    {
                        string formula = "=\"" + HighlightFormulaMarker + i + "\"<>\"\"";
                        Xl.FormatCondition rule = (Xl.FormatCondition)range.FormatConditions.Add(
                            Xl.XlFormatConditionType.xlExpression, Type.Missing, formula);
                        rule.Interior.Color = HighlightColors[i % HighlightColors.Length];
                        try { rule.StopIfTrue = false; }
                        catch (COMException)
                        {
                            // Older Excel versions do not expose StopIfTrue for every rule type.
                        }
                    }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "LinkedRangeRegistry.SetHighlightsVisible.Range"); }
                    finally { Release(range); }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LinkedRangeRegistry.SetHighlightsVisible");
                WinForms.MessageBox.Show("Linked ranges could not be highlighted.", "Linked Ranges");
            }
        }

        public static bool AreHighlightsVisible(Xl.Application app)
        {
            Xl.Workbook workbook = app?.ActiveWorkbook;
            if (workbook == null) return false;
            try
            {
                foreach (Xl.Worksheet sheet in workbook.Worksheets)
                {
                    Xl.Range used = null;
                    try
                    {
                        used = sheet.UsedRange;
                        int count = used.FormatConditions.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            Xl.FormatCondition rule = null;
                            try
                            {
                                rule = (Xl.FormatCondition)used.FormatConditions.Item(i);
                                if (IsHighlightRule(rule)) return true;
                            }
                            catch (COMException)
                            {
                                // A rule can disappear while conditional formatting is recalculated.
                            }
                            catch (Exception ex)
                            {
                                ExceptionLogger.Log(ex, "LinkedRangeRegistry.AreHighlightsVisible.Rule");
                            }
                            finally { Release(rule); }
                        }
                    }
                    finally { Release(used); }
                }
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "LinkedRangeRegistry.AreHighlightsVisible"); }
            return false;
        }

        private static void RemoveHighlightRules(Xl.Workbook workbook)
        {
            foreach (Xl.Worksheet sheet in workbook.Worksheets)
            {
                Xl.Range used = null;
                try
                {
                    used = sheet.UsedRange;
                    for (int i = used.FormatConditions.Count; i >= 1; i--)
                    {
                        Xl.FormatCondition rule = null;
                        try
                        {
                            rule = (Xl.FormatCondition)used.FormatConditions.Item(i);
                            if (IsHighlightRule(rule)) rule.Delete();
                        }
                        catch (Exception ex) { ExceptionLogger.Log(ex, "LinkedRangeRegistry.RemoveHighlightRules.Rule"); }
                        finally { Release(rule); }
                    }
                }
                finally { Release(used); }
            }
        }

        private static bool IsHighlightRule(Xl.FormatCondition rule)
        {
            if (rule == null) return false;
            try
            {
                string formula = Convert.ToString(rule.Formula1) ?? "";
                return formula.IndexOf(HighlightFormulaMarker, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch (COMException)
            {
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LinkedRangeRegistry.IsHighlightRule");
                return false;
            }
        }

        private static List<Xl.Range> GetValidRanges(Xl.Workbook workbook, bool deleteStale)
        {
            var result = new List<Xl.Range>();
            var stale = new List<Xl.Name>();
            foreach (Xl.Name name in workbook.Names)
            {
                try
                {
                    string localName = name.Name;
                    int bang = localName.LastIndexOf('!');
                    if (bang >= 0) localName = localName.Substring(bang + 1);
                    if (!localName.StartsWith(NamePrefix, StringComparison.OrdinalIgnoreCase)) continue;
                    Xl.Range range = null;
                    try { range = name.RefersToRange; }
                    catch (COMException)
                    {
                        // A broken workbook name is classified as stale below.
                    }
                    if (range != null) result.Add(range); else if (deleteStale) stale.Add(name);
                }
                catch (COMException)
                {
                    // Invalid or concurrently removed workbook names are skipped.
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "LinkedRangeRegistry.GetValidRanges.Name");
                }
            }
            foreach (Xl.Name name in stale)
            {
                try { name.Delete(); } catch (Exception ex) { ExceptionLogger.Log(ex, "LinkedRangeRegistry.DeleteStale"); }
            }
            return result;
        }

        private static void Release(object value)
        {
            try
            {
                if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
            }
            catch (COMException)
            {
                // The temporary Excel wrapper was already released.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LinkedRangeRegistry.Release");
            }
        }
    }
}
