using System;
using System.Collections.Generic;
using System.Windows;
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
        Name = "Save Agenda Layout",
        Tooltip = "Save agenda layout",
        IconData = "M15,9H5V5H15M12,19A3,3 0 0,1 9,16A3,3 0 0,1 12,13A3,3 0 0,1 15,16A3,3 0 0,1 12,19M17,3H5C3.89,3 3,3.9 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V7L17,3Z",
        Color = "#F43F5E",
        Description = "Saves the current agenda layout configuration.",
        DetailedHelpText = "### Save Agenda Layout\nExports the current Agenda Wizard configuration (all sections, durations, speakers) as a portable file that can be reimported into any other presentation.",
        Keywords = "agenda, save, layout, configuration, export, toc",
        MinSelection = 0,
        RequiredType = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionNone)]
    public static class SaveAgendaLayoutFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - returns false as Save Agenda Layout requires dialog interaction.
        /// Manual switch in MainSidebar.xaml.cs handles opening the SaveAgendaLayoutWindow dialog.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            new SaveAgendaLayoutWindow(manager).Show();
            return true;
        }
    }
}
