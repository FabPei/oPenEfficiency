using System.Windows;
using System.Windows.Input;

namespace oPenEfficiency.UI
{
    public partial class CreateFolderDialog : Window
    {
        public string FolderName => TxtFolderName.Text;

        public CreateFolderDialog()
        {
            InitializeComponent();
            // Owner is automatically set to the active window when ShowDialog() is called
            TxtFolderName.Focus();
        }

        private void BtnCreate_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtFolderName.Text))
            {
                MessageBox.Show("Please enter a folder name.", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void TxtFolderName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                BtnCreate_Click(sender, e);
            }
            else if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }
    }
}
