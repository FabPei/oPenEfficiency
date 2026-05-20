using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public partial class ExportWizard : Window
    {
        private PowerPointManager _manager;
        private bool _isManualFilename = false;

        public ExportWizard()
        {
            InitializeComponent();
            
            var app = Globals.ThisAddIn?.Application;
            if (app != null)
            {
                _manager = new PowerPointManager(app);
                InitializeData();
            }

            // Hook events for auto-filename
            TxtTheme.TextChanged += (s, e) => UpdateFilename();
            TxtInitials.TextChanged += (s, e) => UpdateFilename();
            TxtVersion.TextChanged += (s, e) => UpdateFilename();
            PickerDate.SelectedDateChanged += (s, e) => UpdateFilename();
            TxtFinalFilename.TextChanged += (s, e) => { if (TxtFinalFilename.IsFocused) _isManualFilename = true; };
        }

        private void InitializeData()
        {
            try
            {
                var pres = Globals.ThisAddIn.Application.ActivePresentation;
                int totalSlides = pres.Slides.Count;
                TxtScopeAll.Text = $"Entire Presentation ({totalSlides})";

                int selectedCount = 0;
                try
                {
                    selectedCount = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.Count;
                }
                catch { }
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
                try
                {
                    dynamic props = pres.BuiltInDocumentProperties;
                    TxtPropAuthor.Text = LoadProp(props, "Author");
                    TxtPropTitle.Text = LoadProp(props, "Title");
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
                catch { }

                PickerDate.SelectedDate = DateTime.Now;
                UpdateFilename();
            }
            catch { }
        }

        private string LoadProp(dynamic props, string name)
        {
            try { return props[name].Value?.ToString() ?? ""; } catch { return ""; }
        }

        private void UpdateFilename()
        {
            if (_isManualFilename) return;

            string dateStr = PickerDate.SelectedDate?.ToString("yyMMdd") ?? DateTime.Now.ToString("yyMMdd");
            string theme = TxtTheme.Text.Replace(" ", "_");
            string initials = TxtInitials.Text.ToUpper();
            string version = "v" + TxtVersion.Text;

            TxtFinalFilename.Text = $"{dateStr}_{theme}_{initials}_{version}";
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) this.DragMove();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var pres = Globals.ThisAddIn.Application.ActivePresentation;
                
                // 1. Update Properties if requested
                if (ChkApplyPropsOnExport.IsChecked == true)
                {
                    try
                    {
                        dynamic props = pres.BuiltInDocumentProperties;
                        props["Author"].Value = TxtPropAuthor.Text;
                        props["Title"].Value = TxtPropTitle.Text;
                        props["Subject"].Value = TxtPropSubject.Text;
                        props["Comments"].Value = TxtPropComments.Text;
                        props["Keywords"].Value = TxtPropKeywords.Text;
                        props["Category"].Value = TxtPropCategory.Text;
                        props["Manager"].Value = TxtPropManager.Text;
                        props["Company"].Value = TxtPropCompany.Text;
                        props["Content status"].Value = TxtPropStatus.Text;
                        props["Hyperlink base"].Value = TxtPropHyperlinkBase.Text;
                    }
                    catch { }
                }

                bool isMail = DestMail.IsChecked == true;
                bool isClipboard = DestClipboard.IsChecked == true;
                
                string baseFolder;
                string finalName = TxtFinalFilename.Text;

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
                    try { initialDir = Path.GetDirectoryName(pres.FullName); } catch { }
                    if (string.IsNullOrEmpty(initialDir)) initialDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    sfd.InitialDirectory = initialDir;
                    sfd.FileName = TxtFinalFilename.Text;

                    if (FormatPPTX.IsChecked == true) { sfd.DefaultExt = ".pptx"; sfd.Filter = "PowerPoint Presentation|*.pptx"; }
                    else if (FormatPDF.IsChecked == true || FormatProtected.IsChecked == true) { sfd.DefaultExt = ".pdf"; sfd.Filter = "PDF Document|*.pdf"; }
                    else { sfd.Filter = "PowerPoint or PDF|*.pptx;*.pdf"; }

                    if (sfd.ShowDialog() != true) return;

                    baseFolder = Path.GetDirectoryName(sfd.FileName);
                    finalName = Path.GetFileNameWithoutExtension(sfd.FileName);
                }

                string pptPath = Path.Combine(baseFolder, finalName + ".pptx");
                string pdfPath = Path.Combine(baseFolder, finalName + ".pdf");

                PpPrintRangeType rangeType = PpPrintRangeType.ppPrintAll;
                var selectedSlideIds = new System.Collections.Generic.List<int>();
                if (ScopeSelected.IsChecked == true)
                {
                    rangeType = PpPrintRangeType.ppPrintSelection;
                    try
                    {
                        foreach (PowerPoint.Slide s in Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange)
                        {
                            selectedSlideIds.Add(s.SlideID);
                        }
                    }
                    catch { }
                }

                bool exportPpt = FormatPPTX.IsChecked == true || FormatBoth.IsChecked == true;
                bool exportPdf = FormatPDF.IsChecked == true || FormatBoth.IsChecked == true || FormatProtected.IsChecked == true;

                if (exportPpt)
                {
                    pres.SaveCopyAs(pptPath);
                    if (ChkCleanComments.IsChecked == true || ChkCleanHiddenSlides.IsChecked == true || 
                        ChkCleanNotes.IsChecked == true || ChkCleanMetadata.IsChecked == true || 
                        ChkCleanDeveloperNotes.IsChecked == true || ScopeSelected.IsChecked == true)
                    {
                        ProcessCleanup(pptPath, selectedSlideIds);
                    }
                }

                if (exportPdf)
                {
                    pres.ExportAsFixedFormat(pdfPath, PpFixedFormatType.ppFixedFormatTypePDF, 
                        PpFixedFormatIntent.ppFixedFormatIntentScreen, 
                        ChkPdfFrame.IsChecked == true ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse, 
                        PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst, 
                        PpPrintOutputType.ppPrintOutputSlides, Office.MsoTriState.msoFalse, 
                        null, rangeType, "", true, true, true, true, 
                        ChkPdfA.IsChecked == true);
                }

                // Handle Actions
                if (isMail)
                {
                    OpenOutlook(exportPpt ? pptPath : pdfPath, (exportPpt && exportPdf) ? pdfPath : null);
                }
                else if (isClipboard)
                {
                    var fileList = new System.Collections.Specialized.StringCollection();
                    if (exportPpt) fileList.Add(pptPath);
                    if (exportPdf) fileList.Add(pdfPath);
                    Clipboard.SetFileDropList(fileList);
                    MessageBox.Show("Files copied to clipboard.", "Success");
                }
                else
                {
                    // Open folder and select file
                    try
                    {
                        string pathToSelect = exportPpt ? pptPath : pdfPath;
                        if (File.Exists(pathToSelect))
                        {
                            System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{pathToSelect}\"");
                        }
                        else
                        {
                            System.Diagnostics.Process.Start(baseFolder);
                        }
                    }
                    catch { }
                }

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error converting/saving: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ProcessCleanup(string filePath, System.Collections.Generic.List<int> keepSlideIds = null)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                // Open saved copy silently
                var tempPres = app.Presentations.Open(filePath, WithWindow: Office.MsoTriState.msoFalse);
                bool changed = false;

                if (keepSlideIds != null && keepSlideIds.Count > 0)
                {
                    for (int i = tempPres.Slides.Count; i >= 1; i--)
                    {
                        if (!keepSlideIds.Contains(tempPres.Slides[i].SlideID))
                        {
                            tempPres.Slides[i].Delete();
                        }
                    }
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

                if (ChkCleanHiddenSlides.IsChecked == true)
                {
                    for (int i = tempPres.Slides.Count; i >= 1; i--)
                    {
                        if (tempPres.Slides[i].SlideShowTransition.Hidden == Office.MsoTriState.msoTrue)
                            tempPres.Slides[i].Delete();
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
                    foreach (PowerPoint.Slide s in tempPres.Slides)
                    {
                        try { s.Tags.Delete("oPE_SlideNotes"); } catch { }
                    }
                    changed = true;
                }

                if (ChkCleanMetadata.IsChecked == true)
                {
                    tempPres.RemoveDocumentInformation(PpRemoveDocInfoType.ppRDIAll);
                    changed = true;
                }

                if (changed) tempPres.Save();
                tempPres.Close();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Export Cleanup failed: " + ex.Message);
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
