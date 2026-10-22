using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public enum ExportWizardStartFormat
    {
        Pptx,
        Pdf,
        PptxAndPdf,
        Html
    }

    public enum ExportWizardStartScope
    {
        EntirePresentation,
        SelectedSlides
    }

    public partial class ExportWizard : Window
    {
        private PowerPointManager _manager;
        private readonly PowerPoint.Application _application;
        private bool _isManualFilename = false;

        public ExportWizard(
            ExportWizardStartFormat startFormat = ExportWizardStartFormat.Pptx,
            ExportWizardStartScope startScope = ExportWizardStartScope.EntirePresentation,
            PowerPointManager manager = null)
        {
            InitializeComponent();

            _manager = manager;
            var app = _application = manager?.GetApplication();
            if (app != null)
            {
                new System.Windows.Interop.WindowInteropHelper(this).Owner = (IntPtr)app.HWND;
                InitializeData();
            }

            ApplyStartSelection(startFormat, startScope);

            // Hook events for auto-filename
            TxtTheme.TextChanged += (s, e) => UpdateFilename();
            TxtInitials.TextChanged += (s, e) => UpdateFilename();
            TxtVersion.TextChanged += (s, e) => UpdateFilename();
            PickerDate.SelectedDateChanged += (s, e) => UpdateFilename();
            TxtFinalFilename.TextChanged += (s, e) => { if (TxtFinalFilename.IsFocused) _isManualFilename = true; };
        }

        public void ApplyStartSelection(ExportWizardStartFormat startFormat, ExportWizardStartScope startScope)
        {
            FormatPPTX.IsChecked = startFormat == ExportWizardStartFormat.Pptx;
            FormatPDF.IsChecked = startFormat == ExportWizardStartFormat.Pdf;
            FormatBoth.IsChecked = startFormat == ExportWizardStartFormat.PptxAndPdf;
            FormatHTML.IsChecked = startFormat == ExportWizardStartFormat.Html;
            ScopeAll.IsChecked = startScope == ExportWizardStartScope.EntirePresentation;
            ScopeSelected.IsChecked = startScope == ExportWizardStartScope.SelectedSlides;
        }

        private void InitializeData()
        {
            try
            {
                var pres = _application.ActivePresentation;
                int totalSlides = pres.Slides.Count;
                TxtScopeAll.Text = $"Entire Presentation ({totalSlides})";

                int selectedCount = 0;
                try
                {
                    selectedCount = _application.ActiveWindow.Selection.SlideRange.Count;
                }
                catch (Exception ex) { ReportOptionalExportFailure(ex, "ExportWizard.LoadScope.SelectedSlides"); }
                TxtScopeSelected.Text = $"Selected Slides ({selectedCount})";

                bool hasHidden = false;
                foreach (PowerPoint.Slide s in pres.Slides)
                {
                    if (s.SlideShowTransition.Hidden == Office.MsoTriState.msoTrue)
                    {
                        hasHidden = true;
                        break;
                    }
                }
                ChkOnlyVisible.IsEnabled = hasHidden;

                // Load properties
                string presentationTitle = string.Empty;
                try
                {
                    dynamic props = pres.BuiltInDocumentProperties;
                    TxtPropAuthor.Text = LoadProp(props, "Author");
                    presentationTitle = LoadProp(props, "Title");
                    TxtPropTitle.Text = presentationTitle;
                    TxtPropSubject.Text = LoadProp(props, "Subject");
                    TxtPropComments.Text = LoadProp(props, "Comments");
                    TxtPropKeywords.Text = LoadProp(props, "Keywords");
                    TxtPropCategory.Text = LoadProp(props, "Category");
                    TxtPropManager.Text = LoadProp(props, "Manager");

                    // Company is often in BuiltIn but sometimes better accessed via a separate logic if not in standard list
                    TxtPropCompany.Text = LoadProp(props, "Company");
                    TxtPropStatus.Text = LoadProp(props, "Content status");
                    TxtPropHyperlinkBase.Text = LoadProp(props, "Hyperlink base");
                }
                catch (Exception ex) { ReportOptionalExportFailure(ex, "ExportWizard.LoadProperties"); }

                if (string.IsNullOrWhiteSpace(presentationTitle))
                {
                    presentationTitle = Path.GetFileNameWithoutExtension(pres.Name);
                }
                TxtTheme.Text = presentationTitle;

                PickerDate.SelectedDate = DateTime.Now;
                UpdateFilename();
            }
            catch (Exception ex) { ReportOptionalExportFailure(ex, "ExportWizard.Initialize"); }
        }

        private string LoadProp(dynamic props, string name)
        {
            try { return props[name].Value?.ToString() ?? ""; } catch { return ""; }
        }

        private void UpdateFilename()
        {
            if (_isManualFilename) return;

            string dateStr = PickerDate.SelectedDate?.ToString("yyMMdd") ?? DateTime.Now.ToString("yyMMdd");
            string theme = SanitizeFilenameComponent(TxtTheme.Text);
            string initials = TxtInitials.Text.ToUpper();
            string version = "v" + TxtVersion.Text;

            TxtFinalFilename.Text = $"{dateStr}_{theme}_{initials}_{version}";
        }

        private static string SanitizeFilenameComponent(string value)
            => Core.Services.ExportWizardPolicy.SanitizeFilenameComponent(value);

        private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) this.DragMove();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            BtnOK.IsEnabled = false;
            try
            {
                var pres = _application.ActivePresentation;
                List<int> exportSlideIndices = GetExportSlideIndices(pres);

                bool isMail = DestMail.IsChecked == true;
                bool isClipboard = DestClipboard.IsChecked == true;
                
                string baseFolder;
                string overwriteConfirmedBySaveDialog = null;
                string finalName = TxtFinalFilename.Text.Trim();
                if (string.IsNullOrWhiteSpace(finalName) || finalName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                    throw new InvalidOperationException("Enter a valid export filename.");

                if (isMail || isClipboard)
                {
                    baseFolder = Path.Combine(Path.GetTempPath(), "oPE_Export_" + Guid.NewGuid().ToString("N"));
                    Directory.CreateDirectory(baseFolder);
                }
                else
                {
                    // Open Save File Dialog
                    var sfd = new Microsoft.Win32.SaveFileDialog();
                    string initialDir = "";
                    try { initialDir = Path.GetDirectoryName(pres.FullName); }
                    catch (Exception ex) { ReportOptionalExportFailure(ex, "ExportWizard.SelectOutputPath.InitialDirectory"); }
                    if (string.IsNullOrEmpty(initialDir)) initialDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    sfd.InitialDirectory = initialDir;
                    sfd.FileName = TxtFinalFilename.Text;

                    if (FormatPPTX.IsChecked == true) { sfd.DefaultExt = ".pptx"; sfd.Filter = "PowerPoint Presentation|*.pptx"; }
                    else if (FormatHTML.IsChecked == true) { sfd.DefaultExt = ".html"; sfd.Filter = "Editable HTML Presentation|*.html"; }
                    else if (FormatPDF.IsChecked == true) { sfd.DefaultExt = ".pdf"; sfd.Filter = "PDF Document|*.pdf"; }
                    else { sfd.DefaultExt = ".pptx"; sfd.Filter = "PowerPoint Presentation|*.pptx"; }

                    if (sfd.ShowDialog() != true)
                    {
                        BtnOK.IsEnabled = true;
                        return;
                    }

                    baseFolder = Path.GetDirectoryName(sfd.FileName);
                    finalName = Path.GetFileNameWithoutExtension(sfd.FileName);
                    overwriteConfirmedBySaveDialog = Path.GetFullPath(sfd.FileName);
                }

                string pptPath = Path.Combine(baseFolder, finalName + ".pptx");
                string pdfPath = Path.Combine(baseFolder, finalName + ".pdf");
                string htmlPath = Path.Combine(baseFolder, finalName + ".html");

                bool exportPpt = FormatPPTX.IsChecked == true || FormatBoth.IsChecked == true;
                bool exportPdf = FormatPDF.IsChecked == true || FormatBoth.IsChecked == true;
                bool exportHtml = FormatHTML.IsChecked == true;

                if (!isMail && !isClipboard)
                {
                    var existingTargets = new List<string>();
                    if (exportPpt && File.Exists(pptPath) && !PathsEqual(pptPath, overwriteConfirmedBySaveDialog)) existingTargets.Add(pptPath);
                    if (exportPdf && File.Exists(pdfPath) && !PathsEqual(pdfPath, overwriteConfirmedBySaveDialog)) existingTargets.Add(pdfPath);
                    if (exportHtml && File.Exists(htmlPath) && !PathsEqual(htmlPath, overwriteConfirmedBySaveDialog)) existingTargets.Add(htmlPath);
                    if (existingTargets.Count > 0)
                    {
                        var overwrite = MessageBox.Show(
                            "The following export file(s) already exist and will be replaced:\n\n" +
                            string.Join("\n", existingTargets.Select(Path.GetFileName)) +
                            "\n\nContinue?",
                            "Replace Existing Export Files",
                            MessageBoxButton.OKCancel,
                            MessageBoxImage.Warning);
                        if (overwrite != MessageBoxResult.OK)
                        {
                            BtnOK.IsEnabled = true;
                            return;
                        }
                    }
                }

                if (exportPpt)
                {
                    pres.SaveCopyAs(pptPath);
                    ProcessExportCopy(pptPath, exportSlideIndices);
                }

                if (exportPdf)
                {
                    if (exportPpt)
                    {
                        ExportPdfFromCopy(pptPath, pdfPath);
                    }
                    else
                    {
                        ExportPdfFromTemporaryCopy(pres, exportSlideIndices, pdfPath);
                    }
                }

                HtmlExportReport htmlReport = null;
                string htmlPackagePath = null;
                if (exportHtml)
                {
                    htmlReport = new HtmlPresentationExportService().Export(pres, exportSlideIndices, htmlPath);
                    if (isMail || isClipboard)
                    {
                        htmlPackagePath = CreateHtmlPackage(htmlPath, htmlReport.AssetDirectory);
                    }
                }

                if (ChkApplyPropsOnExport.IsChecked == true)
                {
                    ApplyDocumentProperties(pres);
                }

                // Handle Actions
                if (isMail)
                {
                    OpenOutlook(exportHtml ? htmlPackagePath : (exportPpt ? pptPath : pdfPath), (exportPpt && exportPdf) ? pdfPath : null);
                }
                else if (isClipboard)
                {
                    var fileList = new System.Collections.Specialized.StringCollection();
                    if (exportPpt) fileList.Add(pptPath);
                    if (exportPdf) fileList.Add(pdfPath);
                    if (exportHtml) fileList.Add(htmlPackagePath);
                    Clipboard.SetFileDropList(fileList);
                    MessageBox.Show("Files copied to clipboard.", "Success");
                }
                else
                {
                    // Open folder and select file
                    try
                    {
                        string pathToSelect = exportHtml ? htmlPath : (exportPpt ? pptPath : pdfPath);
                        if (File.Exists(pathToSelect))
                        {
                            System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{pathToSelect}\"");
                        }
                        else
                        {
                            System.Diagnostics.Process.Start(baseFolder);
                        }
                    }
                    catch (Exception ex) { ReportOptionalExportFailure(ex, "ExportWizard.RemoveMetadata.Shape"); }
                }

                if (htmlReport != null && htmlReport.Warnings.Count > 0 && !isMail && !isClipboard)
                {
                    MessageBox.Show("The editable HTML was created. " + htmlReport.Warnings.Count + " complex object(s) were exported as PNG fallbacks.", "HTML export", MessageBoxButton.OK, MessageBoxImage.Information);
                }

                this.Close();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExportWizard.BtnOK_Click");
                MessageBox.Show("Error converting/saving: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                BtnOK.IsEnabled = true;
            }
        }

        private List<int> GetExportSlideIndices(Presentation presentation)
        {
            var indices = new List<int>();
            if (ScopeSelected.IsChecked == true)
            {
                try
                {
                    SlideRange selectedSlides = _application.ActiveWindow.Selection.SlideRange;
                    for (int index = 1; index <= selectedSlides.Count; index++)
                    {
                        indices.Add(selectedSlides[index].SlideIndex);
                    }
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "ExportWizard.GetExportSlideIndices.SelectedSlides");
                }

                if (indices.Count == 0)
                    throw new InvalidOperationException("Select at least one slide before exporting the selected-slide scope.");
            }
            else
            {
                indices.AddRange(Enumerable.Range(1, presentation.Slides.Count));
            }

            bool excludeHiddenSlides = ChkOnlyVisible.IsChecked == true || ChkCleanHiddenSlides.IsChecked == true;
            if (excludeHiddenSlides)
            {
                indices = indices
                    .Where(index => presentation.Slides[index].SlideShowTransition.Hidden != Office.MsoTriState.msoTrue)
                    .ToList();
            }

            indices = indices.Distinct().OrderBy(index => index).ToList();
            if (indices.Count == 0)
                throw new InvalidOperationException("The selected export options do not include any slides.");

            return indices;
        }

        private static bool PathsEqual(string first, string second)
            => Core.Services.ExportWizardPolicy.PathsEqual(first, second);

        private static string CreateHtmlPackage(string htmlPath, string assetDirectory)
        {
            string packagePath = Path.ChangeExtension(htmlPath, ".zip");
            if (File.Exists(packagePath)) File.Delete(packagePath);

            using (ZipArchive archive = ZipFile.Open(packagePath, ZipArchiveMode.Create))
            {
                archive.CreateEntryFromFile(htmlPath, Path.GetFileName(htmlPath));
                if (Directory.Exists(assetDirectory))
                {
                    string assetRoot = Path.GetDirectoryName(assetDirectory);
                    foreach (string assetPath in Directory.GetFiles(assetDirectory, "*", SearchOption.AllDirectories))
                    {
                        string archivePath = assetPath.Substring(assetRoot.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                        archive.CreateEntryFromFile(assetPath, archivePath);
                    }
                }
            }

            return packagePath;
        }

        private void ProcessExportCopy(string filePath, ICollection<int> includedSlideIndices)
        {
            Presentation tempPres = null;
            try
            {
                var app = _application;
                tempPres = app.Presentations.Open(filePath, WithWindow: Office.MsoTriState.msoFalse);
                bool changed = false;

                var included = new HashSet<int>(includedSlideIndices);
                for (int index = tempPres.Slides.Count; index >= 1; index--)
                {
                    if (!included.Contains(index))
                    {
                        tempPres.Slides[index].Delete();
                        changed = true;
                    }
                }

                if (ChkCleanMetadata.IsChecked != true)
                {
                    ApplyDocumentProperties(tempPres);
                    changed = true;
                }

                if (ChkCleanComments.IsChecked == true)
                {
                    foreach (PowerPoint.Slide s in tempPres.Slides)
                    {
                        for (int i = s.Comments.Count; i >= 1; i--) s.Comments[i].Delete();
                    }
                    changed = true;
                }

                if (ChkCleanNotes.IsChecked == true)
                {
                    foreach (PowerPoint.Slide s in tempPres.Slides)
                    {
                        if (s.HasNotesPage == Office.MsoTriState.msoTrue)
                        {
                            foreach (PowerPoint.Shape shp in s.NotesPage.Shapes)
                            {
                                if (shp.Type == Office.MsoShapeType.msoPlaceholder && shp.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody)
                                    shp.TextFrame.TextRange.Text = "";
                            }
                        }
                    }
                    changed = true;
                }

                if (ChkCleanDeveloperNotes.IsChecked == true)
                {
                    oPenEfficiency.Services.PrivateCommentsService.DeleteAll(tempPres);
                    changed = true;
                }

                if (ChkCleanMetadata.IsChecked == true)
                {
                    tempPres.RemoveDocumentInformation(PpRemoveDocInfoType.ppRDIAll);
                    changed = true;
                }

                if (changed) tempPres.Save();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExportWizard.ProcessExportCopy");
                throw new InvalidOperationException("The exported presentation copy could not be scoped or cleaned safely.", ex);
            }
            finally
            {
                if (tempPres != null)
                {
                    try { tempPres.Close(); }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "ExportWizard.ProcessExportCopy.Close"); }
                    ExceptionPolicy.RunBestEffortComCleanup(
                        () => Marshal.ReleaseComObject(tempPres),
                        "ExportWizard.ProcessExportCopy.ReleasePresentation");
                }
            }
        }

        private void ExportPdfFromTemporaryCopy(Presentation sourcePresentation, ICollection<int> slideIndices, string pdfPath)
        {
            string tempDirectory = Path.Combine(Path.GetTempPath(), "oPE_PdfExport_" + Guid.NewGuid().ToString("N"));
            string tempPptPath = Path.Combine(tempDirectory, "export-source.pptx");
            Directory.CreateDirectory(tempDirectory);

            try
            {
                sourcePresentation.SaveCopyAs(tempPptPath);
                ProcessExportCopy(tempPptPath, slideIndices);
                ExportPdfFromCopy(tempPptPath, pdfPath);
            }
            finally
            {
                try { Directory.Delete(tempDirectory, true); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ExportWizard.ExportPdfFromTemporaryCopy.Cleanup"); }
            }
        }

        private void ExportPdfFromCopy(string pptPath, string pdfPath)
        {
            Presentation exportPresentation = null;
            try
            {
                exportPresentation = _application.Presentations.Open(
                    pptPath,
                    ReadOnly: Office.MsoTriState.msoTrue,
                    WithWindow: Office.MsoTriState.msoFalse);

                exportPresentation.ExportAsFixedFormat(
                    pdfPath,
                    PpFixedFormatType.ppFixedFormatTypePDF,
                    PpFixedFormatIntent.ppFixedFormatIntentScreen,
                    ChkPdfFrame.IsChecked == true ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse,
                    PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst,
                    PpPrintOutputType.ppPrintOutputSlides,
                    Office.MsoTriState.msoTrue,
                    null,
                    PpPrintRangeType.ppPrintAll,
                    "",
                    ChkCleanMetadata.IsChecked != true,
                    true,
                    true,
                    true,
                    ChkPdfA.IsChecked == true);
            }
            finally
            {
                if (exportPresentation != null)
                {
                    try { exportPresentation.Close(); }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "ExportWizard.ExportPdfFromCopy.Close"); }
                    ExceptionPolicy.RunBestEffortComCleanup(
                        () => Marshal.ReleaseComObject(exportPresentation),
                        "ExportWizard.ExportPdfFromCopy.ReleasePresentation");
                }
            }
        }

        private static void ReportOptionalExportFailure(Exception exception, string context)
        {
            if (!ExceptionPolicy.IsExpectedComFailure(exception))
                ExceptionLogger.Log(exception, context);
            // Optional metadata and selection details do not prevent exporting the remaining content.
        }

        private void ApplyDocumentProperties(Presentation presentation)
        {
            dynamic props = presentation.BuiltInDocumentProperties;
            SetDocumentProperty(props, "Author", TxtPropAuthor.Text);
            SetDocumentProperty(props, "Title", TxtPropTitle.Text);
            SetDocumentProperty(props, "Subject", TxtPropSubject.Text);
            SetDocumentProperty(props, "Comments", TxtPropComments.Text);
            SetDocumentProperty(props, "Keywords", TxtPropKeywords.Text);
            SetDocumentProperty(props, "Category", TxtPropCategory.Text);
            SetDocumentProperty(props, "Manager", TxtPropManager.Text);
            SetDocumentProperty(props, "Company", TxtPropCompany.Text);
            SetDocumentProperty(props, "Content status", TxtPropStatus.Text);
            SetDocumentProperty(props, "Hyperlink base", TxtPropHyperlinkBase.Text);
        }

        private static void SetDocumentProperty(dynamic properties, string name, string value)
        {
            try
            {
                properties[name].Value = value ?? string.Empty;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExportWizard.SetDocumentProperty." + name);
            }
        }

        private void OpenOutlook(string file1, string file2 = null)
        {
            try
            {
                Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                {
                    MessageBox.Show("Outlook not found.");
                    return;
                }

                dynamic outlookApp = Activator.CreateInstance(outlookType);
                dynamic mail = outlookApp.CreateItem(0); // 0 = olMailItem
                
                mail.Subject = Path.GetFileName(file1);
                mail.Attachments.Add(file1);
                if (file2 != null) mail.Attachments.Add(file2);
                
                mail.Display(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Outlook integration failed: " + ex.Message);
            }
        }
    }
}
