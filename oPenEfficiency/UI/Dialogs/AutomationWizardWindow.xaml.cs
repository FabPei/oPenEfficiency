using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using oPenEfficiency.DocumentAutomation;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    public partial class AutomationWizardWindow : Window
    {
        private readonly DocumentAutomationPowerPointService _service;
        private readonly Dictionary<string, Control> _answerControls = new Dictionary<string, Control>(StringComparer.OrdinalIgnoreCase);
        private DocumentAutomationSchema _schema;
        private AutomationExecutionPlan _plan;

        public AutomationWizardWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _service = new DocumentAutomationPowerPointService(manager);
            try
            {
                _schema = _service.LoadSchema();
                if (_schema == null) { PreviewText.Text = "No automation schema is stored. Open the schema designer first."; return; }
                BuildQuestions();
                RefreshPreview();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AutomationWizardWindow.Initialize");
                PreviewText.Text = "The automation schema could not be loaded: " + ex.Message;
            }
        }

        private void BuildQuestions()
        {
            foreach (var question in _schema.Questions)
            {
                var container = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };
                container.Children.Add(new TextBlock { Text = question.Prompt + (question.Required ? " *" : string.Empty), Foreground = (System.Windows.Media.Brush)FindResource("TextSecondaryBrush"), Margin = new Thickness(0, 0, 0, 4) });
                Control control;
                if (question.Type == AutomationQuestionType.YesNo)
                    control = new ComboBox { ItemsSource = new[] { "Yes", "No" }, SelectedIndex = -1 };
                else if (question.Type == AutomationQuestionType.SingleChoice)
                    control = new ComboBox { ItemsSource = question.Options, SelectedIndex = -1 };
                else
                    control = new TextBox { Style = (Style)FindResource("StandardTextBox") };
                control.ToolTip = "Value for {{" + question.Key + "}}";
                container.Children.Add(control);
                QuestionsPanel.Children.Add(container);
                _answerControls[question.Key] = control;
            }
        }

        private Dictionary<string, string> ReadValues()
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in _answerControls)
                values[pair.Key] = pair.Value is TextBox textBox ? textBox.Text : (pair.Value as ComboBox)?.SelectedItem as string;
            return values;
        }

        private void RefreshPreview()
        {
            _plan = _service.BuildPlan(_schema, ReadValues());
            var lines = new List<string>();
            lines.Add("Recognized placeholders: " + (_plan.PlaceholderKeys.Count == 0 ? "None" : string.Join(", ", _plan.PlaceholderKeys.Select(key => "{{" + key + "}}"))));
            lines.Add("Missing values: " + (_plan.MissingValueKeys.Count == 0 ? "None" : string.Join(", ", _plan.MissingValueKeys)));
            lines.Add("Slides to include: " + FormatSlides(true));
            lines.Add("Slides to exclude: " + FormatSlides(false));
            lines.Add("Validation errors: " + (_plan.ValidationIssues.Count == 0 ? "None" : Environment.NewLine + "• " + string.Join(Environment.NewLine + "• ", _plan.ValidationIssues.Select(issue => issue.Message))));
            PreviewText.Text = string.Join(Environment.NewLine + Environment.NewLine, lines);
            GenerateButton.IsEnabled = _plan.CanExecute;
        }

        private string FormatSlides(bool include)
        {
            var decisions = _plan.SlideDecisions.Where(decision => decision.Include == include).ToList();
            return decisions.Count == 0 ? "None" : string.Join(", ", decisions.Select(decision => decision.SlideNumber + " (" + decision.Reason + ")"));
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                RefreshPreview();
                if (!_plan.CanExecute) { MessageBox.Show("Resolve the preview issues before generating a presentation.", "Build Presentation from Answers", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
                var source = _service.ActivePresentation;
                string baseName = string.IsNullOrWhiteSpace(source?.Name) ? "Automated Presentation" : Path.GetFileNameWithoutExtension(source.Name) + " - Automated";
                var dialog = new SaveFileDialog { Title = "Create automated presentation copy", Filter = "PowerPoint Presentation (*.pptx)|*.pptx", DefaultExt = ".pptx", AddExtension = true, OverwritePrompt = true, FileName = baseName };
                if (dialog.ShowDialog(this) != true) return;
                _service.ExecuteToNewPresentation(_schema, ReadValues(), dialog.FileName);
                MessageBox.Show("The new presentation was created successfully. The template was not changed.", "Build Presentation from Answers", MessageBoxButton.OK, MessageBoxImage.Information);
                DialogResult = true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AutomationWizardWindow.Generate");
                MessageBox.Show("No completed output was reported. " + ex.Message, "Build Presentation from Answers", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RefreshPreview_Click(object sender, RoutedEventArgs e) { try { RefreshPreview(); } catch (Exception ex) { ExceptionLogger.Log(ex, "AutomationWizardWindow.RefreshPreview"); PreviewText.Text = ex.Message; GenerateButton.IsEnabled = false; } }
        private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) { if (e.ChangedButton == MouseButton.Left) DragMove(); }
        private void Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}
