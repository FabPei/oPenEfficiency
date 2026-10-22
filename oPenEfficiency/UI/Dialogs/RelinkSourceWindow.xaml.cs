using System.IO;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;

namespace oPenEfficiency.UI
{
    public partial class RelinkSourceWindow : Window
    {
        public RelinkSourceWindow(string workbookPath, string worksheet, string range)
        {
            InitializeComponent();
            WorkbookPathTextBox.Text = workbookPath ?? string.Empty;
            WorksheetTextBox.Text = worksheet ?? string.Empty;
            RangeTextBox.Text = range ?? string.Empty;
        }

        public string WorkbookPath => WorkbookPathTextBox.Text.Trim();
        public string Worksheet => WorksheetTextBox.Text.Trim();
        public string Range => RangeTextBox.Text.Trim();

        private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select the Excel source workbook",
                Filter = "Excel workbooks|*.xlsx;*.xlsm;*.xlsb;*.xls|All files|*.*",
                FileName = WorkbookPath
            };
            if (dialog.ShowDialog(this) == true) WorkbookPathTextBox.Text = dialog.FileName;
        }

        private void Relink_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(WorkbookPath) || !File.Exists(WorkbookPath))
            {
                ValidationText.Text = "Select an existing Excel workbook.";
                WorkbookPathTextBox.Focus();
                return;
            }
            if (string.IsNullOrWhiteSpace(Worksheet))
            {
                ValidationText.Text = "Enter the worksheet name.";
                WorksheetTextBox.Focus();
                return;
            }
            if (string.IsNullOrWhiteSpace(Range))
            {
                ValidationText.Text = "Enter a cell range, for example A1:C4.";
                RangeTextBox.Focus();
                return;
            }

            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
