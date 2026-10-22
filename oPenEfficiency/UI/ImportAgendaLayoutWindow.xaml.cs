using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using oPenEfficiency.Services;

namespace oPenEfficiency.UI
{
    public partial class ImportAgendaLayoutWindow : Window
    {
        private readonly ExternalAgendaImportService _importService;

        public ImportAgendaLayoutWindow()
        {
            InitializeComponent();
            SourceInitialized += (sender, args) => FitToOwnerWorkArea();
            _importService = new ExternalAgendaImportService();
            // The caller assigns the owner before showing this dialog.
        }

        private void FitToOwnerWorkArea()
        {
            var owner = new System.Windows.Interop.WindowInteropHelper(this).Owner;
            var screen = System.Windows.Forms.Screen.FromHandle(owner);
            var dpi = System.Windows.Media.VisualTreeHelper.GetDpi(this);
            double availableWidth = Math.Max(1, (screen.WorkingArea.Width / dpi.DpiScaleX) - 32);
            double availableHeight = Math.Max(1, (screen.WorkingArea.Height / dpi.DpiScaleY) - 32);
            MinWidth = Math.Min(MinWidth, availableWidth);
            MinHeight = Math.Min(MinHeight, availableHeight);
            MaxWidth = availableWidth;
            MaxHeight = availableHeight;
            Width = Math.Min(Width, availableWidth);
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                BtnClose_Click(sender, e);
                e.Handled = true;
            }
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && e.ButtonState == MouseButtonState.Pressed)
                DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void BtnBrowseFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "Agenda layout XML|AgendaCreator.xml;AgendaWizard.xml|All XML files|*.xml|All files|*.*";
            dialog.Title = "Select AgendaCreator.xml or AgendaWizard.xml file";

            if (dialog.ShowDialog() == true)
            {
                TxtXmlPath.Text = dialog.FileName;
            }
        }

        private void BtnBrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select a folder containing AgendaCreator.xml or AgendaWizard.xml files",
                ShowNewFolderButton = false
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                TxtFolderPath.Text = dialog.SelectedPath;
            }
        }

        private void BtnImportFile_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtXmlPath.Text))
            {
                MessageBox.Show("Please select an XML file first.", "Import", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var imported = _importService.ImportFromXml(TxtXmlPath.Text);

                if (imported.Count > 0)
                {
                    ShowResult($"Successfully imported {imported.Count} layout(s):\n\n" + string.Join("\n", imported));
                }
                else
                {
                    ShowResult("No layouts were imported. The file may be invalid or contain no valid layouts.", true);
                }
            }
            catch (Exception ex)
            {
                ShowResult($"Import failed: {ex.Message}", true);
            }
        }

        private void BtnImportFolder_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtFolderPath.Text))
            {
                MessageBox.Show("Please select a folder first.", "Import", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var imported = _importService.ImportFromFolder(TxtFolderPath.Text);

                if (imported.Count > 0)
                {
                    ShowResult($"Successfully imported {imported.Count} layout(s):\n\n" + string.Join("\n", imported));
                }
                else
                {
                    ShowResult("No layouts were imported. The folder may not contain valid AgendaCreator.xml or AgendaWizard.xml files.", true);
                }
            }
            catch (Exception ex)
            {
                ShowResult($"Import failed: {ex.Message}", true);
            }
        }

        private async void BtnAutoDetect_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                BtnAutoDetect.IsEnabled = false;
                TxtAutoDetectStatus.Text = "Searching for AgendaCreator.xml and AgendaWizard.xml files...";

                var files = _importService.FindAgendaCreatorFiles();

                if (files.Count == 0)
                {
                    TxtAutoDetectStatus.Text = "No supported agenda layout XML files found.";
                    BtnAutoDetect.IsEnabled = true;
                    return;
                }

                TxtAutoDetectStatus.Text = $"Found {files.Count} file(s). Importing...";

                // Import from all found files
                var imported = await System.Threading.Tasks.Task.Run(() =>
                {
                    var allImported = new System.Collections.Generic.List<string>();
                    foreach (var file in files)
                    {
                        var result = _importService.ImportFromXml(file);
                        allImported.AddRange(result);
                    }
                    return allImported;
                });

                if (imported.Count > 0)
                {
                    ShowResult($"Auto-detect found {files.Count} file(s) and imported {imported.Count} layout(s):\n\n" + string.Join("\n", imported));
                    TxtAutoDetectStatus.Text = $"Import complete: {imported.Count} layout(s) imported.";
                }
                else
                {
                    TxtAutoDetectStatus.Text = "No new layouts were imported (may already exist).";
                }

                BtnAutoDetect.IsEnabled = true;
            }
            catch (Exception ex)
            {
                TxtAutoDetectStatus.Text = $"Auto-detect failed: {ex.Message}";
                BtnAutoDetect.IsEnabled = true;
                ShowResult($"Auto-detect failed: {ex.Message}", true);
            }
        }

        private void ShowResult(string message, bool isError = false)
        {
            TxtResult.Text = message;
            ResultBorder.Visibility = Visibility.Visible;

            string brushKey = isError ? "ErrorBrush" : "SuccessBrush";
            ResultBorder.SetResourceReference(System.Windows.Controls.Border.BorderBrushProperty, brushKey);
            TxtResultHeading.SetResourceReference(System.Windows.Controls.TextBlock.ForegroundProperty, brushKey);
        }
    }
}
