using System;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Excel.Features.Rounding;
using oPenEfficiency.Excel.Services;
using oPenEfficiency.Excel.UI;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Excel
{
    /// <summary>
    /// VSTO entry point for the Excel add-in.
    /// Kept intentionally small: host-neutral logic lives in oPenEfficiency.Core,
    /// Excel-specific features live under Features/.
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>Central Excel COM access, analogous to PowerPointManager.</summary>
        public ExcelManager Manager { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
            // Build marker: names the assembly file this Excel session actually loaded
            // (diagnostic only, never throws; see %TEMP%\oPenEfficiency\startup_YYYYMMDD.log).
            StartupBuildLog.Record(typeof(ThisAddIn).Assembly, "oPenEfficiency.Excel");

                Manager = new ExcelManager(this.Application);

                // Consistent rounding: register the OEROUND automation add-in and start the
                // consistency engine. Registration may require an Excel restart to take effect.
                OeRoundActions.RegisterForCurrentUser(this.Application);
                OeRoundEngine.Initialize(this.Application);
                this.Application.SheetSelectionChange += Application_SheetSelectionChange;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "Excel.ThisAddIn_Startup");
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try { this.Application.SheetSelectionChange -= Application_SheetSelectionChange; } catch { /* Shutdown continues when Excel already disconnected the event source. */ }
            try { OeRoundEngine.Shutdown(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "Excel.ThisAddIn_Shutdown"); }
        }

        private void Application_SheetSelectionChange(object sheet, Microsoft.Office.Interop.Excel.Range target)
        {
            try { ExcelRibbon.InvalidateLinkStatus(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "Excel.Application_SheetSelectionChange.LinkStatus"); }
        }

        /// <summary>
        /// Provides the ribbon via raw ribbon XML (IRibbonExtensibility),
        /// the same approach used by the PowerPoint add-in.
        /// </summary>
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new ExcelRibbon();
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
