using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI.Dialogs;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnHtmlSlideImport",
        Name = "Import HTML Slides",
        SearchAliases = new[] { "HTML Slide Import" },
        Tooltip = "Import AI HTML slides as editable PowerPoint slides",
        IconData = "M3,3H21V5H3V3M5,7H19V19H5V7M12,9L16,13H13V17H11V13H8L12,9Z",
        Color = "#0EA5E9",
        Description = "Imports supported AI-generated presentation HTML files and rebuilds them as editable PowerPoint slides.",
        DetailedHelpText = "### Import HTML Slides\nChoose a supported AI-generated HTML presentation, review the detected slides, and select which ones to import. Right-click presets source styling, presentation theme, background-only, or local-resources-only behavior; every mode still shows the detected slides before import. Recognized sections, cards, lists, callouts, and process flows become editable PowerPoint shapes. JavaScript is never executed. Local Resources Only never requests remote files. For stability, HTML is limited to 5 MiB and stylesheet downloads are bounded by size, count, redirects, and time.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class HtmlSlideImportFeature
    {
        public enum ImportPreset
        {
            Review,
            PreserveHtmlStyling,
            PresentationTheme,
            PreserveBackgroundOnly,
            LocalResourcesOnly
        }

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, ImportPreset.Review);
        }

        public static bool Execute(PowerPointManager manager, ImportPreset preset)
        {
            try
            {
                var application = manager?.GetApplication();
                if (application?.ActivePresentation == null)
                {
                    MessageBox.Show(
                        "Please open a PowerPoint presentation before importing HTML slides.",
                        "Import HTML Slides",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return false;
                }

                using (var dialog = new OpenFileDialog())
                {
                    dialog.Title = "Select AI-generated Presentation HTML";
                    dialog.Filter = "HTML files (*.html;*.htm)|*.html;*.htm|All files (*.*)|*.*";
                    dialog.CheckFileExists = true;
                    dialog.Multiselect = false;
                    string presentationDirectory = null;
                    try { presentationDirectory = Path.GetDirectoryName(application.ActivePresentation.FullName); } catch { /* Unsaved presentations fall back to the default file-dialog directory. */ }
                    dialog.InitialDirectory = Directory.Exists(presentationDirectory)
                        ? presentationDirectory
                        : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                    if (dialog.ShowDialog() != DialogResult.OK)
                    {
                        return false;
                    }

                    string html = HtmlImportDocumentPipeline.ReadHtmlFile(dialog.FileName);
                    bool referencesRemoteResources = Regex.IsMatch(
                        html,
                        "(?:src|href)\\s*=\\s*['\"]https?://|@import\\s+(?:url\\()?\\s*['\"]?https?://|url\\(\\s*['\"]?https?://",
                        RegexOptions.IgnoreCase,
                        HtmlImportLimits.RegexTimeout);
                    bool allowRemoteResources = false;
                    if (referencesRemoteResources && preset != ImportPreset.LocalResourcesOnly)
                    {
                        var remoteResult = MessageBox.Show(
                            "This HTML references files on the internet. Allow openEfficiency to download remote styles and images for this import?\n\nChoose No to keep the import local and skip remote resources.",
                            "Import HTML Slides – Remote Resources",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning,
                            MessageBoxDefaultButton.Button2);
                        allowRemoteResources = remoteResult == DialogResult.Yes;
                    }

                    var service = new HtmlPresentationImportService(allowRemoteResources);
                    HtmlImportPreview preview = service.GetImportPreview(dialog.FileName);
                    if (preview == null || preview.Slides.Count == 0)
                    {
                        MessageBox.Show(
                            "No supported slide sections were found in the selected HTML file.",
                            "Import HTML Slides",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return false;
                    }

                    var initialOptions = CreateInitialOptions(preset);
                    var optionsDialog = new HtmlSlideImportOptionsWindow(preview, initialOptions);
                    new System.Windows.Interop.WindowInteropHelper(optionsDialog).Owner = (IntPtr)application.HWND;
                    if (optionsDialog.ShowDialog() != true)
                    {
                        return false;
                    }

                    try { application.StartNewUndoEntry(); } catch { /* Import remains valid when undo grouping is unavailable. */ }
                    int importedSlides = service.ImportFromFile(manager, dialog.FileName, optionsDialog.SelectedOptions);

                    if (importedSlides > 0)
                    {
                        string reportText = service.LastReport != null && service.LastReport.HasWarnings
                            ? "\n\nImport notes:\n" + service.LastReport.ToDisplayText(5)
                            : string.Empty;
                        MessageBox.Show(
                            string.Format("Imported {0} slide(s) from {1}.", importedSlides, Path.GetFileName(dialog.FileName)) + reportText,
                            "Import HTML Slides",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return true;
                    }

                    MessageBox.Show(
                        "No slides were imported. Please select at least one slide.",
                        "Import HTML Slides",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
                }
            }
            catch (HtmlImportLimitException ex)
            {
                ExceptionLogger.Log(ex, "HtmlSlideImportFeature.Execute.ResourceLimit");
                MessageBox.Show(
                    ex.Message + " Reduce or split the source and try again.",
                    "Import HTML Slides",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "HtmlSlideImportFeature.Execute");
                MessageBox.Show(
                    "The HTML presentation could not be imported. Please check the error log for details.",
                    "Import HTML Slides",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
        }

        private static HtmlImportOptions CreateInitialOptions(ImportPreset preset)
        {
            switch (preset)
            {
                case ImportPreset.PreserveHtmlStyling:
                    return new HtmlImportOptions
                    {
                        PreserveSourceStyling = true,
                        PreserveSourceBackground = true,
                        PreserveSourceFonts = true
                    };
                case ImportPreset.PresentationTheme:
                    return new HtmlImportOptions { UsePresentationTheme = true };
                case ImportPreset.PreserveBackgroundOnly:
                    return new HtmlImportOptions { PreserveSourceBackground = true };
                default:
                    return null;
            }
        }

    }
}
