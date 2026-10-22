using oPenEfficiency.UI;
using System;
using System.Windows;
using System.Windows.Interop;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Creates and synchronizes agenda slides and deck structure.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnAgendaWizard",
        Name = "Agenda Creator",
        Tooltip = "Create and update an agenda",
        SearchAliases = new[] { "Agenda Wizard" },
        IconData = "M12,2A3,3 0 0,1 15,5V11A3,3 0 0,1 12,14A3,3 0 0,1 9,11V5A3,3 0 0,1 12,2M12,4A1,1 0 0,0 11,5V11A1,1 0 0,0 12,12A1,1 0 0,0 13,11V5A1,1 0 0,0 12,4M5,13H19V21H5V13M7,15V19H17V15H7Z",
        Color = "#F43F5E",
        Description = "Creates or synchronizes a timed agenda and its tagged dividers, sections, starter slides, and optional backup structure from a reusable local layout.",
        DetailedHelpText = "### Agenda Creator\nDefine agenda topics, hierarchy, durations, speakers, visibility, start time, and target end time, then choose a layout. The built-in Standard Agenda is selected automatically for a clean profile, so an agenda can be created without first designing a custom layout. Create or synchronize a main agenda, chapter divider slides, PowerPoint sections, optional starter slides, highlighted current topics, and a tagged backup slide. Existing tagged agendas are updated in place while chapter content is retained by stable agenda IDs. Use the context menu to prepare a full deck structure, derive an agenda from sections and slide titles, validate the agenda against the deck, or open Agenda Layout Designer. The wizard can import external AgendaCreator.xml and legacy AgendaWizard.xml layout definitions without replacing existing local layouts. Generation uses a PowerPoint Undo transaction and restores the previous agenda when a reported generation error occurs. Presentations or affected structures protected by think-cell are left unchanged. Saved agenda definitions are stored locally; saving an existing name asks before replacement.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class AgendaWizardFeature
    {
        private static AgendaWizard _window;

        public enum LaunchAction
        {
            Open,
            PrepareDeckStructure,
            ExtractFromDeck,
            Validate
        }

        /// <summary>
        /// Wrapper for auto-discovery - returns false as Agenda Wizard requires dialog interaction.
        /// Manual switch in MainSidebar.xaml.cs handles opening the AgendaWizard dialog.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, LaunchAction.Open);
        }

        public static bool Execute(PowerPointManager manager, LaunchAction action)
        {
            if (manager == null) return false;
            try
            {
                if (_window != null && _window.IsLoaded)
                {
                    _window.Activate();
                    ApplyLaunchAction(_window, action);
                    return true;
                }

                _window = new AgendaWizard(manager);
                var app = manager.GetApplication();
                if (app != null) new WindowInteropHelper(_window).Owner = (IntPtr)app.HWND;
                _window.Closed += (s, e) => _window = null;
                _window.Show();
                ApplyLaunchAction(_window, action);
                return true;
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "AgendaWizardFeature.Execute");
                var failedWindow = _window;
                _window = null;
                try { failedWindow?.Close(); }
                catch (Exception closeEx)
                {
                    oPenEfficiency.Utils.ExceptionLogger.Log(
                        closeEx,
                        "AgendaWizardFeature.Execute.FailureCleanup");
                }
                throw;
            }
        }

        private static void ApplyLaunchAction(AgendaWizard window, LaunchAction action)
        {
            switch (action)
            {
                case LaunchAction.PrepareDeckStructure:
                    window.PrepareDeckStructure();
                    break;
                case LaunchAction.ExtractFromDeck:
                    window.ExtractFromDeckCommand();
                    break;
                case LaunchAction.Validate:
                    window.ShowValidationReport();
                    break;
            }
        }
    }
}
