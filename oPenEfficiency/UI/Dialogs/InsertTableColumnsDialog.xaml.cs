using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;

namespace oPenEfficiency.UI
{
    public partial class InsertTableColumnsDialog : Window
    {
        private readonly bool _rows;
        public int ColumnCount { get; private set; } = 2;
        public bool PreserveWidth => PreserveWidthCheckBox.IsChecked == true;
        public int ItemCount => ColumnCount;
        public bool PreserveDimension => PreserveWidth;

        public InsertTableColumnsDialog(bool rows = false)
        {
            _rows = rows;
            InitializeComponent();
            if (_rows)
            {
                Title = "Insert Multiple Rows";
                DialogTitle.Text = "INSERT MULTIPLE ROWS";
                CountLabel.Text = "Number of rows (2–50)";
                PreserveWidthCheckBox.Content = "Preserve total table height";
            }
            Loaded += (s, e) =>
            {
                ColumnCountTextBox.Focus();
                ColumnCountTextBox.SelectAll();
            };
        }

        private void Insert_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(ColumnCountTextBox.Text, out int count) || count < 2 || count > 50)
            {
                string item = _rows ? "rows" : "columns";
                MessageBox.Show($"Enter a number of {item} between 2 and 50.", $"Invalid {(_rows ? "Row" : "Column")} Count", MessageBoxButton.OK, MessageBoxImage.Information);
                ColumnCountTextBox.Focus();
                ColumnCountTextBox.SelectAll();
                return;
            }

            ColumnCount = count;
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            e.Handled = Regex.IsMatch(e.Text, "[^0-9]+");
        }
    }
}
