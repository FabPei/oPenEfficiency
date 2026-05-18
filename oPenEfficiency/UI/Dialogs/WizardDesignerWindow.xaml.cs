using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;

namespace oPenEfficiency.UI
{
    public partial class WizardDesignerWindow : Window
    {
        public ObservableCollection<WizardQuestion> Questions { get; set; } = new ObservableCollection<WizardQuestion>();
        public ObservableCollection<string> QuestionTypes { get; set; } = new ObservableCollection<string> { "Text Input", "Yes/No", "Picture", "Single Choice", "Multiple Choice" };

        public WizardDesignerWindow(PowerPointManager pm)
        {
            InitializeComponent();
            DataContext = this;
            Questions.Add(new WizardQuestion { Type = "Text Input", Prompt = "Client Name" });
            Questions.Add(new WizardQuestion { Type = "Yes/No", Prompt = "Include Appendix?" });
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e) { if (e.ChangedButton == MouseButton.Left) DragMove(); }
        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private void AddQuestion_Click(object sender, RoutedEventArgs e)
        {
            Questions.Add(new WizardQuestion { Type = "Text Input", Prompt = "New Question" });
        }

        private void RemoveQuestion_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn && btn.Tag is WizardQuestion q)
            {
                Questions.Remove(q);
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Template questions saved. In a full implementation, this serializes the schema into the Presentation's CustomXMLParts.", "Saved", MessageBoxButton.OK, MessageBoxImage.Information);
            Close();
        }
    }

    public class WizardQuestion
    {
        public string Type { get; set; }
        public string Prompt { get; set; }
    }
}
