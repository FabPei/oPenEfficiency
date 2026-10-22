using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Interop;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Models;
using oPenEfficiency.UI;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Save agenda layout configuration.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnSaveAgendaLayout",
        Name = "Agenda Layout Designer",
        Tooltip = "Design and save an agenda layout",
        SearchAliases = new[] { "Agenda Layout Creator", "Save Agenda Layout" },
        IconData = "M15,9H5V5H15M12,19A3,3 0 0,1 9,16A3,3 0 0,1 12,13A3,3 0 0,1 15,16A3,3 0 0,1 12,19M17,3H5C3.89,3 3,3.9 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V7L17,3Z",
        Color = "#F43F5E",
        Description = "Captures formatting and placement from selected shapes as a reusable local Agenda Creator layout.",
        DetailedHelpText = "### Agenda Layout Designer\nAssign shapes from the active slide to agenda roles, validate their dimensions and spacing against the actual slide size, and save the result as a local layout. Saving an existing custom name requires confirmation. The built-in Standard Agenda can be inspected and reused but cannot be overwritten or deleted. A clean local template copy is generated without the assigned marker shapes; the source slide is not changed. Custom layout configuration and templates are stored together under the local AgendaLayouts folder. This command saves layout formatting, not agenda topics, durations, or speakers.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class SaveAgendaLayoutFeature
    {
        private static SaveAgendaLayoutWindow _window;

        /// <summary>
        /// Auto-discovery entry - opens the Save Agenda Layout window and returns true
        /// Manual switch in MainSidebar.xaml.cs handles opening the SaveAgendaLayoutWindow dialog.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                if (_window != null && _window.IsLoaded)
                {
                    _window.Activate();
                    return true;
                }

                _window = new SaveAgendaLayoutWindow(manager);
                var app = manager.GetApplication();
                if (app != null) new WindowInteropHelper(_window).Owner = (IntPtr)app.HWND;
                _window.Closed += (s, e) => _window = null;
                _window.Show();
                return true;
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "SaveAgendaLayoutFeature.Execute");
                var failedWindow = _window;
                _window = null;
                try { failedWindow?.Close(); }
                catch (Exception closeEx)
                {
                    oPenEfficiency.Utils.ExceptionLogger.Log(
                        closeEx,
                        "SaveAgendaLayoutFeature.Execute.FailureCleanup");
                }
                throw;
            }
        }
    }
}
