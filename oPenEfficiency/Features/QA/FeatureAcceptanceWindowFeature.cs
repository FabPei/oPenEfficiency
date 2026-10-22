using System;
using System.Windows;
using System.Windows.Interop;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: opens the developer/QA feature-acceptance status window (QA-S6).
    /// Dev-only: Execute refuses to open when no repository checkout is reachable
    /// (FeatureAcceptanceStatusService.IsAvailable), so production installs never expose it.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnFeatureAcceptanceWindow",
        Name = "QA Status",
        Tooltip = "Feature acceptance status (dev)",
        IconData = "M4,5H20V19H4V5M6,7V9H18V7H6M6,11V13H18V11H6M6,15V17H14V15H6",
        Color = "#22C55E",
        Description = "Shows the feature acceptance harness status: scenario evidence, severity and age, with commands that launch the existing gate, evidence and runner scripts. Available only where the repository checkout is reachable (dev builds).",
        DetailedHelpText = "### QA Status (dev)\n\nOpens the developer/QA window for the feature acceptance harness. The window reads the acceptance manifests and the local TestResults evidence, derives severity and age, and offers commands that launch the existing eng scripts (Test-Changed gate, Bind-OfficeEvidenceAndReviews, scenario runners) plus shortcuts to the evidence folder and the QA checklist. It contains no gate logic of its own and is hidden on machines without a repository checkout.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None,
        // Dev-only surface: hidden from the production sidebar and therefore from the generated
        // "All Features" preset (the validator only checks user-visible records). The window is
        // reached through the dev-gated ribbon button (MainRibbon getVisible) or the seam.
        UserVisible = false)]
    public static class FeatureAcceptanceWindowFeature
    {
        private static UI.Dialogs.FeatureAcceptanceWindow _window;

        /// <summary>Currently open window, or null. Used by the qa-acceptance-window host probe.</summary>
        public static UI.Dialogs.FeatureAcceptanceWindow CurrentWindow
        {
            get { return _window; }
        }

        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                if (!FeatureAcceptanceStatusService.IsAvailable)
                {
                    // Dev gate: no checkout -> no QA window (production installs take this path).
                    return false;
                }
                if (_window != null && _window.IsLoaded)
                {
                    _window.Activate();
                    return true;
                }
                _window = new UI.Dialogs.FeatureAcceptanceWindow();
                if (manager != null)
                {
                    var app = manager.GetApplication();
                    if (app != null) new WindowInteropHelper(_window).Owner = (IntPtr)app.HWND;
                }
                _window.Closed += (s, e) => _window = null;
                _window.Show();
                _window.Activate();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FeatureAcceptanceWindowFeature.Execute");
                var failedWindow = _window;
                _window = null;
                try { failedWindow?.Close(); }
                catch (Exception closeEx)
                {
                    ExceptionLogger.Log(closeEx, "FeatureAcceptanceWindowFeature.Execute.FailureCleanup");
                }
                return false;
            }
        }
    }
}
