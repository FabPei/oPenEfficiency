using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace oPenEfficiency.DocumentAutomation
{
    public enum AutomationQuestionType { Text, YesNo, SingleChoice }
    public enum AutomationConditionOperator { Equals, NotEquals }

    [DataContract]
    public sealed class DocumentAutomationSchema
    {
        public const int CurrentVersion = 1;
        [DataMember(Order = 1)] public int SchemaVersion { get; set; } = CurrentVersion;
        [DataMember(Order = 2)] public string SchemaId { get; set; } = Guid.NewGuid().ToString("N");
        [DataMember(Order = 3)] public string Title { get; set; } = "Build Presentation from Answers";
        [DataMember(Order = 4)] public List<AutomationQuestion> Questions { get; set; } = new List<AutomationQuestion>();
        [DataMember(Order = 5)] public List<AutomationSlideCondition> SlideConditions { get; set; } = new List<AutomationSlideCondition>();
    }

    [DataContract]
    public sealed class AutomationQuestion
    {
        [DataMember(Order = 1)] public string Key { get; set; }
        [DataMember(Order = 2)] public string Prompt { get; set; }
        [DataMember(Order = 3)] public AutomationQuestionType Type { get; set; }
        [DataMember(Order = 4)] public bool Required { get; set; }
        [DataMember(Order = 5)] public List<string> Options { get; set; } = new List<string>();
    }

    [DataContract]
    public sealed class AutomationSlideCondition
    {
        [DataMember(Order = 1)] public int SlideId { get; set; }
        [DataMember(Order = 2)] public string QuestionKey { get; set; }
        [DataMember(Order = 3)] public AutomationConditionOperator Operator { get; set; }
        [DataMember(Order = 4)] public string ExpectedValue { get; set; }
    }

    public sealed class AutomationSlideSnapshot
    {
        public int SlideId { get; set; }
        public int SlideNumber { get; set; }
        public IList<string> TextFragments { get; set; } = new List<string>();
    }

    public sealed class AutomationValidationIssue
    {
        public string Code { get; set; }
        public string Message { get; set; }
        public override string ToString() => Message;
    }

    public sealed class AutomationSlideDecision
    {
        public int SlideId { get; set; }
        public int SlideNumber { get; set; }
        public bool Include { get; set; }
        public string Reason { get; set; }
    }

    public sealed class AutomationExecutionPlan
    {
        public IList<AutomationValidationIssue> ValidationIssues { get; set; } = new List<AutomationValidationIssue>();
        public IList<string> PlaceholderKeys { get; set; } = new List<string>();
        public IList<string> MissingValueKeys { get; set; } = new List<string>();
        public IList<AutomationSlideDecision> SlideDecisions { get; set; } = new List<AutomationSlideDecision>();
        public bool CanExecute => ValidationIssues.Count == 0 && MissingValueKeys.Count == 0;
    }

    public static class DocumentAutomationSchemaSerializer
    {
        public static string Serialize(DocumentAutomationSchema schema)
        {
            if (schema == null) throw new ArgumentNullException(nameof(schema));
            using (var stream = new MemoryStream())
            {
                new DataContractJsonSerializer(typeof(DocumentAutomationSchema)).WriteObject(stream, schema);
                return Encoding.UTF8.GetString(stream.ToArray());
            }
        }

        public static DocumentAutomationSchema Deserialize(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return null;
            DocumentAutomationSchema schema;
            try
            {
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                    schema = (DocumentAutomationSchema)new DataContractJsonSerializer(typeof(DocumentAutomationSchema)).ReadObject(stream);
            }
            catch (Exception ex) when (ex is SerializationException || ex is XmlException)
            {
                throw new InvalidDataException("The document automation schema is not valid JSON.", ex);
            }
            if (schema == null) throw new InvalidDataException("The document automation schema is empty.");
            if (schema.SchemaVersion > DocumentAutomationSchema.CurrentVersion)
                throw new NotSupportedException($"Document automation schema version {schema.SchemaVersion} is newer than supported version {DocumentAutomationSchema.CurrentVersion}.");
            if (schema.SchemaVersion < 1)
                throw new InvalidDataException("The document automation schema version is invalid.");
            schema.Questions = schema.Questions ?? new List<AutomationQuestion>();
            schema.SlideConditions = schema.SlideConditions ?? new List<AutomationSlideCondition>();
            foreach (var question in schema.Questions.Where(question => question != null))
                question.Options = question.Options ?? new List<string>();
            return schema;
        }
    }

    public static class DocumentAutomationSchemaRecovery
    {
        public static DocumentAutomationSchema DeserializeNewestValid(IEnumerable<string> candidatesNewestFirst)
        {
            if (candidatesNewestFirst == null) return null;
            InvalidDataException lastInvalid = null;
            foreach (string candidate in candidatesNewestFirst)
            {
                try
                {
                    DocumentAutomationSchema schema = DocumentAutomationSchemaSerializer.Deserialize(candidate);
                    if (schema != null) return schema;
                    lastInvalid = new InvalidDataException("A stored document automation schema is empty.");
                }
                catch (NotSupportedException) { throw; }
                catch (InvalidDataException ex) { lastInvalid = ex; }
            }
            if (lastInvalid != null) throw new InvalidDataException("No valid stored document automation schema could be recovered.", lastInvalid);
            return null;
        }
    }

    public sealed class DocumentAutomationValidator
    {
        public const int MaximumQuestions = 200;
        public const int MaximumConditions = 1000;
        public const int MaximumKeyLength = 64;
        public const int MaximumPromptLength = 500;
        public const int MaximumOptionsPerQuestion = 100;
        public const int MaximumOptionLength = 200;
        public const int MaximumAnswerLength = 10000;
        private static readonly Regex KeyPattern = new Regex("^[A-Za-z][A-Za-z0-9_]*$", RegexOptions.CultureInvariant);

        public IList<AutomationValidationIssue> Validate(DocumentAutomationSchema schema, IEnumerable<int> knownSlideIds = null)
        {
            var issues = new List<AutomationValidationIssue>();
            if (schema == null) { Add(issues, "SchemaMissing", "No document automation schema is available."); return issues; }
            if (schema.SchemaVersion != DocumentAutomationSchema.CurrentVersion)
                Add(issues, "SchemaVersion", $"Schema version {schema.SchemaVersion} is not supported.");
            var questions = schema.Questions ?? new List<AutomationQuestion>();
            if (questions.Count > MaximumQuestions) Add(issues, "TooManyQuestions", $"A schema can contain at most {MaximumQuestions} questions.");
            var conditions = schema.SlideConditions ?? new List<AutomationSlideCondition>();
            if (conditions.Count > MaximumConditions) Add(issues, "TooManyConditions", $"A schema can contain at most {MaximumConditions} slide conditions.");
            var duplicateKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var seenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var question in questions)
            {
                if (question == null) { Add(issues, "QuestionMissing", "A question entry is empty."); continue; }
                string key = (question.Key ?? string.Empty).Trim();
                if (key.Length > MaximumKeyLength) Add(issues, "KeyTooLong", $"Question key '{key.Substring(0, MaximumKeyLength)}…' exceeds {MaximumKeyLength} characters.");
                if (!Enum.IsDefined(typeof(AutomationQuestionType), question.Type)) Add(issues, "InvalidQuestionType", $"Question '{key}' has an unsupported type.");
                if (!KeyPattern.IsMatch(key)) Add(issues, "InvalidKey", $"Question key '{key}' must start with a letter and contain only letters, numbers, or underscores.");
                if (!seenKeys.Add(key) && duplicateKeys.Add(key)) Add(issues, "DuplicateKey", $"Question key '{key}' is duplicated.");
                if (string.IsNullOrWhiteSpace(question.Prompt)) Add(issues, "PromptMissing", $"Question '{key}' needs a prompt.");
                else if (question.Prompt.Length > MaximumPromptLength) Add(issues, "PromptTooLong", $"Question '{key}' exceeds the {MaximumPromptLength}-character prompt limit.");
                if (question.Type == AutomationQuestionType.SingleChoice)
                {
                    var options = (question.Options ?? new List<string>()).Where(o => !string.IsNullOrWhiteSpace(o)).Select(o => o.Trim()).ToList();
                    if (options.Count < 2) Add(issues, "OptionsMissing", $"Single-choice question '{key}' needs at least two options.");
                    if (options.Count > MaximumOptionsPerQuestion) Add(issues, "TooManyOptions", $"Single-choice question '{key}' exceeds {MaximumOptionsPerQuestion} options.");
                    if (options.Any(option => option.Length > MaximumOptionLength)) Add(issues, "OptionTooLong", $"Single-choice question '{key}' has an option longer than {MaximumOptionLength} characters.");
                    if (options.Count != options.Distinct(StringComparer.OrdinalIgnoreCase).Count()) Add(issues, "DuplicateOption", $"Single-choice question '{key}' has duplicate options.");
                }
            }
            var known = knownSlideIds == null ? null : new HashSet<int>(knownSlideIds);
            var conditionedSlides = new HashSet<int>();
            foreach (var condition in conditions)
            {
                if (condition == null) { Add(issues, "ConditionMissing", "A slide condition entry is empty."); continue; }
                if (!Enum.IsDefined(typeof(AutomationConditionOperator), condition.Operator)) Add(issues, "InvalidConditionOperator", $"Condition for slide ID {condition.SlideId} has an unsupported operator.");
                if (condition.SlideId <= 0 || (known != null && !known.Contains(condition.SlideId))) Add(issues, "UnknownSlide", $"Condition references unknown slide ID {condition.SlideId}.");
                if (!conditionedSlides.Add(condition.SlideId)) Add(issues, "NestedCondition", $"Slide ID {condition.SlideId} has more than one condition; the MVP supports one simple condition per slide.");
                var question = questions.FirstOrDefault(q => q != null && string.Equals(q.Key, condition.QuestionKey, StringComparison.OrdinalIgnoreCase));
                if (question == null) Add(issues, "UnknownConditionKey", $"Slide ID {condition.SlideId} references unknown question '{condition.QuestionKey}'.");
                else if (question.Type == AutomationQuestionType.Text) Add(issues, "UnsupportedConditionType", $"Slide conditions can only use Yes/No or single-choice questions; '{question.Key}' is text.");
                if (string.IsNullOrWhiteSpace(condition.ExpectedValue)) Add(issues, "ExpectedValueMissing", $"Condition for slide ID {condition.SlideId} needs an expected value.");
                else if (question?.Type == AutomationQuestionType.YesNo && !AutomationValue.IsYesNo(condition.ExpectedValue)) Add(issues, "InvalidYesNo", $"Condition for '{question.Key}' must expect Yes or No.");
                else if (question?.Type == AutomationQuestionType.SingleChoice && !(question.Options ?? new List<string>()).Contains(condition.ExpectedValue, StringComparer.OrdinalIgnoreCase)) Add(issues, "InvalidChoice", $"Condition value '{condition.ExpectedValue}' is not an option for '{question.Key}'.");
            }
            return issues;
        }

        private static void Add(ICollection<AutomationValidationIssue> issues, string code, string message) => issues.Add(new AutomationValidationIssue { Code = code, Message = message });
    }

    public static class AutomationValue
    {
        public static bool IsYesNo(string value) => TryNormalizeYesNo(value, out _);
        public static bool TryNormalizeYesNo(string value, out string normalized)
        {
            normalized = null;
            if (string.Equals(value?.Trim(), "Yes", StringComparison.OrdinalIgnoreCase) || string.Equals(value?.Trim(), "True", StringComparison.OrdinalIgnoreCase)) { normalized = "Yes"; return true; }
            if (string.Equals(value?.Trim(), "No", StringComparison.OrdinalIgnoreCase) || string.Equals(value?.Trim(), "False", StringComparison.OrdinalIgnoreCase)) { normalized = "No"; return true; }
            return false;
        }
        public static bool EqualsNormalized(string left, string right)
        {
            if (TryNormalizeYesNo(left, out string a) && TryNormalizeYesNo(right, out string b)) return a == b;
            return string.Equals(left?.Trim(), right?.Trim(), StringComparison.OrdinalIgnoreCase);
        }
    }

    public static class PlaceholderEngine
    {
        private static readonly Regex PlaceholderPattern = new Regex(@"\{\{\s*([A-Za-z][A-Za-z0-9_]*)\s*\}\}", RegexOptions.CultureInvariant);
        public static IList<string> Detect(IEnumerable<string> texts) => (texts ?? Enumerable.Empty<string>()).Where(t => t != null).SelectMany(t => PlaceholderPattern.Matches(t).Cast<Match>()).Select(m => m.Groups[1].Value).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(k => k, StringComparer.OrdinalIgnoreCase).ToList();
        public static string Replace(string text, IReadOnlyDictionary<string, string> values)
        {
            if (text == null) return null;
            return PlaceholderPattern.Replace(text, match => TryGet(values, match.Groups[1].Value, out string value) ? value ?? string.Empty : match.Value);
        }
        public static IList<PlaceholderReplacement> PlanReplacements(string text, IReadOnlyDictionary<string, string> values)
        {
            if (string.IsNullOrEmpty(text)) return new List<PlaceholderReplacement>();
            return PlaceholderPattern.Matches(text).Cast<Match>()
                .Where(match => TryGet(values, match.Groups[1].Value, out string _))
                .Select(match =>
                {
                    TryGet(values, match.Groups[1].Value, out string value);
                    return new PlaceholderReplacement { Start = match.Index, Length = match.Length, Key = match.Groups[1].Value, Value = value ?? string.Empty };
                })
                .OrderByDescending(replacement => replacement.Start)
                .ToList();
        }
        public static IList<string> Missing(IEnumerable<string> keys, IReadOnlyDictionary<string, string> values) => (keys ?? Enumerable.Empty<string>()).Where(key => !TryGet(values, key, out string value) || string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(k => k, StringComparer.OrdinalIgnoreCase).ToList();
        private static bool TryGet(IReadOnlyDictionary<string, string> values, string key, out string value)
        {
            value = null; if (values == null) return false;
            if (values.TryGetValue(key, out value)) return true;
            foreach (var pair in values) if (string.Equals(pair.Key, key, StringComparison.OrdinalIgnoreCase)) { value = pair.Value; return true; }
            return false;
        }
    }

    public sealed class PlaceholderReplacement
    {
        public int Start { get; set; }
        public int Length { get; set; }
        public string Key { get; set; }
        public string Value { get; set; }
    }

    public sealed class DocumentAutomationPlanner
    {
        public AutomationExecutionPlan Build(DocumentAutomationSchema schema, IEnumerable<AutomationSlideSnapshot> slides, IReadOnlyDictionary<string, string> values)
        {
            var snapshots = (slides ?? Enumerable.Empty<AutomationSlideSnapshot>()).ToList();
            var plan = new AutomationExecutionPlan();
            plan.ValidationIssues = new DocumentAutomationValidator().Validate(schema, snapshots.Select(s => s.SlideId));
            if (schema == null) return plan;
            plan.PlaceholderKeys = PlaceholderEngine.Detect(snapshots.SelectMany(s => s.TextFragments ?? new List<string>()));
            var questions = (schema.Questions ?? new List<AutomationQuestion>()).Where(question => question != null).ToList();
            var conditions = (schema.SlideConditions ?? new List<AutomationSlideCondition>()).Where(condition => condition != null).ToList();
            foreach (string placeholderKey in plan.PlaceholderKeys.Where(key => !questions.Any(question => string.Equals(question.Key, key, StringComparison.OrdinalIgnoreCase))))
                plan.ValidationIssues.Add(new AutomationValidationIssue { Code = "UnknownPlaceholder", Message = $"Placeholder '{{{{{placeholderKey}}}}}' has no matching question." });
            var requiredKeys = questions.Where(q => q.Required).Select(q => q.Key)
                .Concat(conditions.Select(condition => condition.QuestionKey));
            plan.MissingValueKeys = PlaceholderEngine.Missing(plan.PlaceholderKeys.Concat(requiredKeys), values);
            foreach (var question in questions)
            {
                string answer = null;
                bool hasAnswer = values != null && values.Any(pair => string.Equals(pair.Key, question.Key, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(answer = pair.Value));
                if (!hasAnswer) continue;
                if (answer.Length > DocumentAutomationValidator.MaximumAnswerLength)
                    plan.ValidationIssues.Add(new AutomationValidationIssue { Code = "AnswerTooLong", Message = $"Answer for '{question.Key}' exceeds {DocumentAutomationValidator.MaximumAnswerLength} characters." });
                if (question.Type == AutomationQuestionType.YesNo && !AutomationValue.IsYesNo(answer))
                    plan.ValidationIssues.Add(new AutomationValidationIssue { Code = "InvalidAnswer", Message = $"Answer for '{question.Key}' must be Yes or No." });
                if (question.Type == AutomationQuestionType.SingleChoice && !(question.Options ?? new List<string>()).Contains(answer, StringComparer.OrdinalIgnoreCase))
                    plan.ValidationIssues.Add(new AutomationValidationIssue { Code = "InvalidAnswer", Message = $"Answer '{answer}' is not an option for '{question.Key}'." });
            }
            foreach (var slide in snapshots.OrderBy(s => s.SlideNumber))
            {
                var condition = conditions.FirstOrDefault(c => c.SlideId == slide.SlideId);
                bool include = true; string reason = "No condition";
                if (condition != null)
                {
                    string actual = null; bool hasValue = values != null && values.Any(p => string.Equals(p.Key, condition.QuestionKey, StringComparison.OrdinalIgnoreCase) && (actual = p.Value) != null);
                    bool equals = hasValue && AutomationValue.EqualsNormalized(actual, condition.ExpectedValue);
                    include = condition.Operator == AutomationConditionOperator.Equals ? equals : !equals;
                    reason = $"{condition.QuestionKey} {condition.Operator} {condition.ExpectedValue}";
                }
                plan.SlideDecisions.Add(new AutomationSlideDecision { SlideId = slide.SlideId, SlideNumber = slide.SlideNumber, Include = include, Reason = reason });
            }
            return plan;
        }
    }
}
