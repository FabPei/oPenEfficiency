using System;
using Xl = global::Microsoft.Office.Interop.Excel;

namespace oPenEfficiency.Excel.Services
{
    /// <summary>
    /// Central Excel COM access point, analogous to PowerPointManager in the PowerPoint add-in.
    /// All Excel-specific features should flow through this wrapper.
    /// </summary>
    public class ExcelManager
    {
        private readonly Xl.Application _app;

        public ExcelManager(Xl.Application app)
        {
            _app = app;
        }

        /// <summary>The live Excel application object.</summary>
        public Xl.Application Application => _app;

        /// <summary>The active workbook, or null if none is open.</summary>
        public Xl.Workbook ActiveWorkbook => _app?.ActiveWorkbook;

        /// <summary>The active worksheet, or null if none is active.</summary>
        public Xl.Worksheet ActiveSheet => _app?.ActiveSheet as Xl.Worksheet;

        /// <summary>
        /// The currently selected cell range, or null if the selection is not a range
        /// (e.g. a chart or shape is selected).
        /// </summary>
        public Xl.Range SelectedRange
        {
            get
            {
                try { return _app?.Selection as Xl.Range; }
                catch { return null; }
            }
        }
    }
}
