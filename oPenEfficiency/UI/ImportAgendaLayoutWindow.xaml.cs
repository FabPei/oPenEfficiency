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
            _importService = new ExternalAgendaImportService();
            // Owner is automatically set to the active window when ShowDialog() is called
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void BtnBrowseFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "AgendaCreator XML|AgendaCreator.xml|All files|*.*";
            dialog.Title = "Select AgendaCreator.xml file";

            if (dialog.ShowDialog() == true)
            {
                TxtXmlPath.Text = dialog.FileName;
            }
        }

        private void BtnBrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select folder containing AgendaCreator.xml files",
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
                    ShowResult("No layouts were imported. The folder may not contain valid AgendaCreator.xml files.", true);
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
                TxtAutoDetectStatus.Text = "Searching for AgendaCreator.xml files...";

                var files = _importService.FindAgendaCreatorFiles();

                if (files.Count == 0)
                {
                    TxtAutoDetectStatus.Text = "No AgendaCreator.xml files found.";
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

            if (isError)
            {
                ResultBorder.Background = (Brush)new BrushConverter().ConvertFrom("#3A1D1E");
                ResultBorder.BorderBrush = (Brush)new BrushConverter().ConvertFrom("#5A2D2E");
            }
        }
    }
}
