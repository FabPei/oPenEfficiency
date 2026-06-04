using System;
using oPenEfficiency.Services.Attributes;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnDocumentAutomation",
        Name = "Document Automation",
        Tooltip = "Run or Design Automation Wizards",
        IconData = "M19,3H5C3.89,3 3,3.89 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5C21,3.89 20.1,3 19,3M19,19H5V5H19V19M17,17H7V15H17V17M17,13H7V11H17V13M17,9H7V7H17V9Z",
        Color = "#8B5CF6",
        Description = "Launch the Document Automation Wizard or Designer.",
        DetailedHelpText = "Document Automation.",
        Keywords = "wizard, designer, macro, template, generate, flow",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class DocumentAutomationFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            var res = MessageBox.Show("Do you want to run the Automation Wizard for this document? (Click 'No' to open the Wizard Designer instead)", "Document Automation", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
            if (res == MessageBoxResult.Yes)
            {
                var window = new oPenEfficiency.UI.AutomationWizardWindow(manager);
                window.ShowDialog();
            }
            else if (res == MessageBoxResult.No)
            {
                var window = new oPenEfficiency.UI.WizardDesignerWindow(manager);
                window.ShowDialog();
            }
            return true;
        }
    }
}
