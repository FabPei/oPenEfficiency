using oPenEfficiency.UI;
using System.Windows;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Agenda-Wizard (Platzhalter für spätere Implementierung).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnAgendaWizard",
        Name = "Agenda Wizard",
        Tooltip = "Agenda wizard",
        IconData = "M12,2A3,3 0 0,1 15,5V11A3,3 0 0,1 12,14A3,3 0 0,1 9,11V5A3,3 0 0,1 12,2M12,4A1,1 0 0,0 11,5V11A1,1 0 0,0 12,12A1,1 0 0,0 13,11V5A1,1 0 0,0 12,4M5,13H19V21H5V13M7,15V19H17V15H7Z",
        Color = "#F43F5E",
        Description = "Opens the Agenda Wizard for creating structured presentations.",
        DetailedHelpText = "### Agenda Wizard\nGenerates and manages a dynamic presentation agenda. Define sections, durations, and speakers in a grid. The wizard automatically creates separator slides, calculates time slots per chapter, and synchronizes slide reordering with physical slide moves.",
        MinSelection = 0,
        RequiredType = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionNone)]
    public static class AgendaWizardFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - returns false as Agenda Wizard requires dialog interaction.
        /// Manual switch in MainSidebar.xaml.cs handles opening the AgendaWizard dialog.
        /// </summary>
         public static bool Execute(PowerPointManager manager)
        {
            new AgendaWizard(manager).Show();
            return true;
        }
    }
}
