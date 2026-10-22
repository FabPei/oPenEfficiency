using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using oPenEfficiency.Excel.Features.Auditing;
using oPenEfficiency.Excel.Features.ChartScanner;
using oPenEfficiency.Excel.Features.Data;
using oPenEfficiency.Excel.Features.PowerPointLink;
using oPenEfficiency.Excel.Features.Rounding;
using oPenEfficiency.Excel.Features.View;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Excel.UI
{
    /// <summary>
    /// Excel ribbon, provided as raw ribbon XML via IRibbonExtensibility
    /// (same pattern as the PowerPoint add-in's MainRibbon).
    /// </summary>
    [ComVisible(true)]
    public class ExcelRibbon : IRibbonExtensibility
    {
        private static ExcelRibbon _current;
        private IRibbonUI _ribbon;
        private static ThisAddIn Host => Globals.ThisAddIn;

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("oPenEfficiency.Excel.UI.ExcelRibbon.xml");
        }

        #region Ribbon Callbacks

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _current = this;
            _ribbon = ribbonUI;
        }

        public void OnRemoveDuplicates(IRibbonControl control)
        {
            try
            {
                var manager = Host?.Manager;
                if (manager != null)
                {
                    RemoveDuplicatesFeature.Execute(manager);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.OnRemoveDuplicates");
            }
        }

        public void OnValidateDates(IRibbonControl control)
        {
            try
            {
                var manager = Host?.Manager;
                if (manager != null) DateValidationFeature.Execute(manager);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.OnValidateDates");
            }
        }

        public void OnTraceDependents(IRibbonControl control)
        {
            try
            {
                var manager = Host?.Manager;
                if (manager != null) TraceDependentsFeature.Execute(manager);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.OnTraceDependents");
            }
        }

        public void OnTracePrecedents(IRibbonControl control)
        {
            try
            {
                var manager = Host?.Manager;
                if (manager != null) TracePrecedentsFeature.Execute(manager);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.OnTracePrecedents");
            }
        }

        public void OnClearArrows(IRibbonControl control)
        {
            try
            {
                var manager = Host?.Manager;
                if (manager != null) ClearArrowsFeature.Execute(manager);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.OnClearArrows");
            }
        }

        private double _oePrecision = 1.0;

        public void OnOePrecision(IRibbonControl control, string selectedId, int selectedIndex)
        {
            switch (selectedId)
            {
                case "oep_1": _oePrecision = 1; break;
                case "oep_0_5": _oePrecision = 0.5; break;
                case "oep_0_01": _oePrecision = 0.01; break;
                case "oep_5": _oePrecision = 5; break;
                case "oep_10": _oePrecision = 10; break;
                case "oep_100": _oePrecision = 100; break;
                default: _oePrecision = 1; break;
            }
        }

        private void OeWrap(string func)
        {
            try
            {
                var manager = Host?.Manager;
                if (manager != null) OeRoundActions.Wrap(manager, func, _oePrecision);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.OeWrap." + func);
            }
        }

        public void OnOeRound(IRibbonControl control) => OeWrap("OEROUND");
        public void OnOeRoundUp(IRibbonControl control) => OeWrap("OEROUNDUP");
        public void OnOeRoundDown(IRibbonControl control) => OeWrap("OEROUNDDOWN");
        public void OnOeRoundNear(IRibbonControl control) => OeWrap("OEROUNDNEAR");

        public void OnOeRemove(IRibbonControl control)
        {
            try
            {
                var manager = Host?.Manager;
                if (manager != null) OeRoundActions.Remove(manager);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.OnOeRemove");
            }
        }

        public void OnOeHighlight(IRibbonControl control)
        {
            try
            {
                var manager = Host?.Manager;
                if (manager != null) OeRoundActions.Highlight(manager);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.OnOeHighlight");
            }
        }

        public void OnChartScanner(IRibbonControl control)
        {
            try
            {
                var manager = Host?.Manager;
                if (manager != null) ChartScannerFeature.Execute(manager);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.OnChartScanner");
            }
        }

        public void OnInsertSlideSnapshot(IRibbonControl control)
        {
            try
            {
                var manager = Host?.Manager;
                if (manager != null) InsertSlideSnapshotFeature.Execute(manager);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.OnInsertSlideSnapshot");
            }
        }

        public void OnRefreshSlideSnapshot(IRibbonControl control)
        {
            try
            {
                var manager = Host?.Manager;
                if (manager != null) RefreshSlideSnapshotFeature.Execute(manager);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.OnRefreshSlideSnapshot");
            }
        }

        public void OnLinkRangeToPptTable(IRibbonControl control)
        {
            try
            {
                var manager = Host?.Manager;
                if (manager != null) LinkRangeToPowerPointTableFeature.Execute(manager, false);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.OnLinkRangeToPptTable");
            }
        }

        public void OnLinkRangeValues(IRibbonControl control)
        {
            try
            {
                var manager = Host?.Manager;
                if (manager != null) LinkRangeToPowerPointTableFeature.Execute(manager, false);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.OnLinkRangeValues");
            }
        }

        public void OnLinkRangeFormatted(IRibbonControl control)
        {
            try
            {
                var manager = Host?.Manager;
                if (manager != null) LinkRangeToPowerPointTableFeature.Execute(manager, true);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.OnLinkRangeFormatted");
            }
        }

        public void OnLinkedPptRanges(IRibbonControl control)
        {
            try { LinkedRangeRegistry.ShowLinkedRanges(Host?.Application); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ExcelRibbon.OnLinkedPptRanges"); }
        }

        public void OnToggleLinkedRangeHighlights(IRibbonControl control, bool pressed)
        {
            try
            {
                LinkedRangeRegistry.SetHighlightsVisible(Host?.Application, pressed);
                _ribbon?.InvalidateControl("BtnHighlightLinkedPptRanges");
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ExcelRibbon.OnToggleLinkedRangeHighlights"); }
        }

        public bool GetLinkedRangeHighlightsPressed(IRibbonControl control)
        {
            return LinkedRangeRegistry.AreHighlightsVisible(Host?.Application);
        }

        public string GetPptLinkStatus(IRibbonControl control)
        {
            return LinkedRangeRegistry.GetSelectionStatus(Host?.Application);
        }

        /// <summary>
        /// Ribbon build hint: labels the loaded assembly's write time, so the owner and support
        /// can see which build the running Excel session loads (see StartupBuildLog).
        /// </summary>
        public string GetBuildStampLabel(IRibbonControl control)
        {
            try
            {
                return StartupBuildLog.FormatBuildHint(typeof(ThisAddIn).Assembly);
            }
            catch (Exception ex)
            {
                // The ribbon must render even when the build hint cannot be read.
                ExceptionLogger.Log(ex, "ExcelRibbon.GetBuildStampLabel");
                return string.Empty;
            }
        }

        internal static void InvalidateLinkStatus()
        {
            try { _current?._ribbon?.InvalidateControl("LblPptLinkStatus"); }
            catch { /* Ribbon invalidation is best-effort while Excel is closing or rebuilding the ribbon. */ }
        }

        public bool GetFormulasPressed(IRibbonControl control)
        {
            try
            {
                var manager = Host?.Manager;
                return manager != null && ShowFormulasFeature.IsShowingFormulas(manager);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.GetFormulasPressed");
                return false;
            }
        }

        public void OnToggleFormulas(IRibbonControl control, bool pressed)
        {
            try
            {
                var manager = Host?.Manager;
                if (manager != null)
                {
                    ShowFormulasFeature.SetShowFormulas(manager, pressed);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelRibbon.OnToggleFormulas");
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (var reader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (reader != null)
                        {
                            return reader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
