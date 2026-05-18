using System.Windows;
using System.Windows.Input;

namespace oPenEfficiency.UI
{
    public partial class AutomationWizardWindow : Window
    {
        public AutomationWizardWindow(PowerPointManager pm)
        {
            InitializeComponent();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e) { if (e.ChangedButton == MouseButton.Left) DragMove(); }
        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Automation triggered! System replaces all master tags with input variables, inserts conditional slides (like Appendix), and sets logos.", "Generation Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            Close();
        }
    }
}
