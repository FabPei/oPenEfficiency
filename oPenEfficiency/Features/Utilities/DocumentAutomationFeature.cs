using System;
using oPenEfficiency.Services.Attributes;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnDocumentAutomation",
        Name = "Build Presentation from Answers",
        SearchAliases = new[] { "Document Automation" },
        Tooltip = "Create a presentation copy from answers and template rules",
        IconData = "M19,3H5C3.89,3 3,3.89 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5C21,3.89 20.1,3 19,3M19,19H5V5H19V19M17,17H7V15H17V17M17,13H7V11H17V13M17,9H7V7H17V9Z",
        Color = "#8B5CF6",
        Description = "Builds a versioned question schema in the presentation and creates automated copies with text replacements and simple slide conditions.",
        DetailedHelpText = "### Build Presentation from Answers\nUse **Design Schema** to add, edit, remove, reorder, validate, and save Text, Yes/No, and Single Choice questions. Place `{{Key}}` tokens in ordinary shapes, nested groups, or table cells; surrounding character formatting is preserved. Add at most one simple Equals or NotEquals condition per slide, based on a Yes/No or Single Choice answer. Use **Run Automation** to review detected placeholders, missing values, slide inclusion decisions, and validation errors. Generation always creates and edits a new `.pptx` copy; the template remains unchanged. Corrupt stored schema parts fall back to the newest valid version, while unsupported newer schema versions are rejected. Concurrent runs may use different targets, but a second run to the same target fails safely. Schemas are limited to 200 questions, 1,000 slide conditions, 100 options per question, 64-character keys, 500-character prompts, 200-character options, and 10,000-character answers. REST, external data sources, image questions, nested conditions, and formulas are not supported in this MVP.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class DocumentAutomationFeature
    {
        public enum Mode
        {
            Ask,
            Run,
            Designer
        }

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, Mode.Ask);
        }

        public static bool Execute(PowerPointManager manager, Mode mode)
        {
            if (manager == null) return false;
            Window window = null;
            try
            {
                var app = manager.GetApplication();
                if (app?.ActivePresentation == null) return false;

                MessageBoxResult res = MessageBoxResult.None;
                if (mode == Mode.Ask)
                {
                    res = MessageBox.Show(
                        "Run document automation?\n\nChoose No to open the schema designer instead.",
                        "Build Presentation from Answers",
                        MessageBoxButton.YesNoCancel,
                        MessageBoxImage.Information);
                    if (res == MessageBoxResult.Cancel) return false;
                }

                bool openRun = mode == Mode.Run || (mode == Mode.Ask && res == MessageBoxResult.Yes);
                window = openRun
                    ? (Window)new oPenEfficiency.UI.AutomationWizardWindow(manager)
                    : new oPenEfficiency.UI.WizardDesignerWindow(manager);
                new System.Windows.Interop.WindowInteropHelper(window).Owner = (IntPtr)app.HWND;
                window.ShowDialog();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DocumentAutomationFeature.Execute");
                try { window?.Close(); }
                catch (Exception closeEx)
                {
                    ExceptionLogger.Log(
                        closeEx,
                        "DocumentAutomationFeature.Execute.FailureCleanup");
                }
                throw;
            }
        }
    }
}
