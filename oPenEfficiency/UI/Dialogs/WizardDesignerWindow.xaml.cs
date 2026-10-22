using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using oPenEfficiency.DocumentAutomation;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    public partial class WizardDesignerWindow : Window
    {
        private readonly DocumentAutomationPowerPointService _service;
        public ObservableCollection<DesignerQuestion> Questions { get; } = new ObservableCollection<DesignerQuestion>();
        public ObservableCollection<DesignerCondition> Conditions { get; } = new ObservableCollection<DesignerCondition>();
        public IList<AutomationQuestionType> QuestionTypes { get; } = Enum.GetValues(typeof(AutomationQuestionType)).Cast<AutomationQuestionType>().ToList();
        public IList<AutomationConditionOperator> ConditionOperators { get; } = Enum.GetValues(typeof(AutomationConditionOperator)).Cast<AutomationConditionOperator>().ToList();

        public WizardDesignerWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _service = new DocumentAutomationPowerPointService(manager);
            DataContext = this;
            try { LoadSchema(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "WizardDesignerWindow.LoadSchema"); StatusText.Text = ex.Message; }
        }

        private void LoadSchema()
        {
            var schema = _service.LoadSchema();
            if (schema == null) { StatusText.Text = "No schema is stored yet. Add the first question."; return; }
            foreach (var question in schema.Questions) Questions.Add(DesignerQuestion.FromModel(question));
            var slides = _service.CaptureSlides().ToDictionary(slide => slide.SlideId, slide => slide.SlideNumber);
            foreach (var condition in schema.SlideConditions)
                Conditions.Add(new DesignerCondition { SlideNumber = (slides.TryGetValue(condition.SlideId, out int number) ? number : 0).ToString(), QuestionKey = condition.QuestionKey, Operator = condition.Operator, ExpectedValue = condition.ExpectedValue });
            StatusText.Text = $"Loaded schema version {schema.SchemaVersion} with {Questions.Count} question(s).";
        }

        private DocumentAutomationSchema BuildSchema()
        {
            var slides = _service.CaptureSlides();
            var schema = new DocumentAutomationSchema { Questions = Questions.Select(question => question.ToModel()).ToList() };
            foreach (var condition in Conditions)
            {
                int.TryParse(condition.SlideNumber, out int slideNumber);
                var slide = slides.FirstOrDefault(item => item.SlideNumber == slideNumber);
                schema.SlideConditions.Add(new AutomationSlideCondition { SlideId = slide?.SlideId ?? 0, QuestionKey = condition.QuestionKey?.Trim(), Operator = condition.Operator, ExpectedValue = condition.ExpectedValue?.Trim() });
            }
            return schema;
        }

        private void AddQuestion_Click(object sender, RoutedEventArgs e) => Questions.Add(new DesignerQuestion { Type = AutomationQuestionType.Text, Prompt = "New question", Key = NextKey() });
        private string NextKey() { int i = Questions.Count + 1; while (Questions.Any(q => string.Equals(q.Key, "Question" + i, StringComparison.OrdinalIgnoreCase))) i++; return "Question" + i; }
        private void RemoveQuestion_Click(object sender, RoutedEventArgs e) { if ((sender as Button)?.Tag is DesignerQuestion question) Questions.Remove(question); }
        private void MoveQuestionUp_Click(object sender, RoutedEventArgs e) => MoveQuestion((sender as Button)?.Tag as DesignerQuestion, -1);
        private void MoveQuestionDown_Click(object sender, RoutedEventArgs e) => MoveQuestion((sender as Button)?.Tag as DesignerQuestion, 1);
        private void MoveQuestion(DesignerQuestion question, int delta) { if (question == null) return; int oldIndex = Questions.IndexOf(question), newIndex = oldIndex + delta; if (newIndex >= 0 && newIndex < Questions.Count) Questions.Move(oldIndex, newIndex); }
        private void AddCondition_Click(object sender, RoutedEventArgs e) => Conditions.Add(new DesignerCondition { SlideNumber = "1", Operator = AutomationConditionOperator.Equals });
        private void RemoveCondition_Click(object sender, RoutedEventArgs e) { if ((sender as Button)?.Tag is DesignerCondition condition) Conditions.Remove(condition); }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var schema = BuildSchema();
                var issues = new DocumentAutomationValidator().Validate(schema, _service.CaptureSlides().Select(slide => slide.SlideId));
                if (issues.Count > 0) { StatusText.Text = string.Join("  ", issues.Select(issue => issue.Message)); return; }
                _service.SaveSchema(schema);
                MessageBox.Show("The versioned automation schema was saved in this presentation.", "Build Presentation from Answers", MessageBoxButton.OK, MessageBoxImage.Information);
                DialogResult = true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "WizardDesignerWindow.Save");
                StatusText.Text = "The schema was not saved: " + ex.Message;
            }
        }

        private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) { if (e.ChangedButton == MouseButton.Left) DragMove(); }
        private void Close_Click(object sender, RoutedEventArgs e) => Close();
    }

    public sealed class DesignerQuestion
    {
        public string Key { get; set; }
        public string Prompt { get; set; }
        public AutomationQuestionType Type { get; set; }
        public bool Required { get; set; }
        public string OptionsText { get; set; }
        public AutomationQuestion ToModel() => new AutomationQuestion { Key = Key?.Trim(), Prompt = Prompt?.Trim(), Type = Type, Required = Required, Options = (OptionsText ?? string.Empty).Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(option => option.Trim()).ToList() };
        public static DesignerQuestion FromModel(AutomationQuestion question) => new DesignerQuestion { Key = question.Key, Prompt = question.Prompt, Type = question.Type, Required = question.Required, OptionsText = string.Join("; ", question.Options ?? new List<string>()) };
    }

    public sealed class DesignerCondition
    {
        public string SlideNumber { get; set; }
        public string QuestionKey { get; set; }
        public AutomationConditionOperator Operator { get; set; }
        public string ExpectedValue { get; set; }
    }
}
