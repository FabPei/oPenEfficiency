using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.DocumentAutomation;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class DocumentAutomationHardeningTests
    {
        [TestMethod]
        public void RecoveryUsesNewestValidSchema()
        {
            string older = JsonSchema("Older");
            string newer = JsonSchema("Newer");
            Assert.AreEqual("Newer", DocumentAutomationSchemaRecovery.DeserializeNewestValid(new[] { newer, older }).Title);
        }

        [TestMethod]
        public void RecoverySkipsCorruptNewestCandidate()
        {
            Assert.AreEqual("Older", DocumentAutomationSchemaRecovery.DeserializeNewestValid(new[] { "{truncated", JsonSchema("Older") }).Title);
        }

        [TestMethod]
        public void RecoveryRejectsUnsupportedNewestVersionInsteadOfSilentlyDowngrading()
        {
            string future = JsonSchema("Future").Replace("\"SchemaVersion\":1", "\"SchemaVersion\":2");
            Assert.ThrowsException<NotSupportedException>(() => DocumentAutomationSchemaRecovery.DeserializeNewestValid(new[] { future, JsonSchema("Older") }));
        }

        [TestMethod]
        public void RecoveryRejectsAStoreWithNoValidCandidate()
        {
            Assert.ThrowsException<InvalidDataException>(() => DocumentAutomationSchemaRecovery.DeserializeNewestValid(new[] { "{", "not-json" }));
        }

        [TestMethod]
        public void SerializerRejectsTruncatedAndMalformedJson()
        {
            Assert.ThrowsException<InvalidDataException>(() => DocumentAutomationSchemaSerializer.Deserialize("{\"SchemaVersion\":"));
            Assert.ThrowsException<InvalidDataException>(() => DocumentAutomationSchemaSerializer.Deserialize("not-json"));
        }

        [TestMethod]
        public void NullEntriesAndUnknownEnumsBecomeValidationIssues()
        {
            var schema = DocumentAutomationSchemaSerializer.Deserialize("{\"SchemaVersion\":1,\"Questions\":[null,{\"Key\":\"Type\",\"Prompt\":\"Type\",\"Type\":99}],\"SlideConditions\":[null]}");
            var issues = new DocumentAutomationValidator().Validate(schema);
            Assert.IsTrue(issues.Any(issue => issue.Code == "QuestionMissing"));
            Assert.IsTrue(issues.Any(issue => issue.Code == "InvalidQuestionType"));
            Assert.IsTrue(issues.Any(issue => issue.Code == "ConditionMissing"));
        }

        [TestMethod]
        public void SchemaSizeLimitsAreValidatedWithoutThrowing()
        {
            var schema = new DocumentAutomationSchema();
            for (int index = 0; index <= DocumentAutomationValidator.MaximumQuestions; index++)
                schema.Questions.Add(new AutomationQuestion { Key = "Q" + index, Prompt = new string('P', DocumentAutomationValidator.MaximumPromptLength + 1), Type = AutomationQuestionType.Text });
            schema.Questions[0].Key = new string('K', DocumentAutomationValidator.MaximumKeyLength + 1);
            schema.Questions.Add(new AutomationQuestion
            {
                Key = "Choice",
                Prompt = "Choice",
                Type = AutomationQuestionType.SingleChoice,
                Options = Enumerable.Range(0, DocumentAutomationValidator.MaximumOptionsPerQuestion + 1).Select(index => index == 0 ? new string('O', DocumentAutomationValidator.MaximumOptionLength + 1) : "Option" + index).ToList()
            });
            for (int index = 0; index <= DocumentAutomationValidator.MaximumConditions; index++)
                schema.SlideConditions.Add(new AutomationSlideCondition { SlideId = index + 1, QuestionKey = "Missing", ExpectedValue = "Yes" });

            var codes = new HashSet<string>(new DocumentAutomationValidator().Validate(schema).Select(issue => issue.Code));
            CollectionAssert.IsSubsetOf(new[] { "TooManyQuestions", "TooManyConditions", "KeyTooLong", "PromptTooLong", "TooManyOptions", "OptionTooLong" }, codes.ToList());
        }

        [TestMethod]
        public void OversizedAnswerIsReportedByPlanning()
        {
            var schema = new DocumentAutomationSchema { Questions = new List<AutomationQuestion> { new AutomationQuestion { Key = "Value", Prompt = "Value", Type = AutomationQuestionType.Text } } };
            var values = new Dictionary<string, string> { ["Value"] = new string('x', DocumentAutomationValidator.MaximumAnswerLength + 1) };
            var plan = new DocumentAutomationPlanner().Build(schema, new AutomationSlideSnapshot[0], values);
            Assert.IsTrue(plan.ValidationIssues.Any(issue => issue.Code == "AnswerTooLong"));
        }

        [TestMethod]
        public void UnicodeSchemaAndValuesRoundTripWithoutLoss()
        {
            var schema = new DocumentAutomationSchema
            {
                Title = "Präsentation 東京",
                Questions = new List<AutomationQuestion>
                {
                    new AutomationQuestion { Key = "Region", Prompt = "Région auswählen", Type = AutomationQuestionType.SingleChoice, Options = new List<string> { "München", "東京" } }
                }
            };
            var clone = DocumentAutomationSchemaSerializer.Deserialize(DocumentAutomationSchemaSerializer.Serialize(schema));
            Assert.AreEqual("Präsentation 東京", clone.Title);
            Assert.AreEqual("東京", clone.Questions[0].Options[1]);
            Assert.AreEqual("Ort: 東京", PlaceholderEngine.Replace("Ort: {{Region}}", new Dictionary<string, string> { ["Region"] = "東京" }));
        }

        [TestMethod]
        public void ReplacementPlanIsReverseOrderedAndHandlesAdjacentTokensEmptyValuesAndNewlines()
        {
            const string source = "A {{One}}\r\n{{Two}}{{One}} Z";
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["One"] = "", ["Two"] = "東京" };
            IList<PlaceholderReplacement> replacements = PlaceholderEngine.PlanReplacements(source, values);
            Assert.AreEqual(3, replacements.Count);
            Assert.IsTrue(replacements[0].Start > replacements[1].Start && replacements[1].Start > replacements[2].Start);
            string result = source;
            foreach (var replacement in replacements) result = result.Remove(replacement.Start, replacement.Length).Insert(replacement.Start, replacement.Value);
            Assert.AreEqual("A \r\n東京 Z", result);
        }

        [TestMethod]
        public void OneBasedTraversalUsesOfficeIndicesAndReverseDeletionOrder()
        {
            var forward = new List<int>();
            var reverse = new List<int>();
            OneBasedCollectionTraversal.Forward(4, forward.Add);
            OneBasedCollectionTraversal.Reverse(4, reverse.Add);
            CollectionAssert.AreEqual(new[] { 1, 2, 3, 4 }, forward);
            CollectionAssert.AreEqual(new[] { 4, 3, 2, 1 }, reverse);
        }

        [TestMethod]
        public void NestedGroupTraversalVisitsLeavesInStableOneBasedOrder()
        {
            var root = Node.Group(Node.Leaf("A"), Node.Group(Node.Leaf("B"), Node.Leaf("C")));
            var visited = new List<string>();
            OneBasedCollectionTraversal.VisitLeaves(root, node => node.Children.Count, (node, oneBasedIndex) => node.Children[oneBasedIndex - 1], node => visited.Add(node.Name));
            CollectionAssert.AreEqual(new[] { "A", "B", "C" }, visited);
        }

        [TestMethod]
        public void TraversalPropagatesAStaleObjectFailureInsteadOfSkippingIt()
        {
            var visited = new List<int>();
            InvalidOperationException error = Assert.ThrowsException<InvalidOperationException>(() =>
                OneBasedCollectionTraversal.Forward(3, index => { if (index == 2) throw new InvalidOperationException("stale COM object"); visited.Add(index); }));
            Assert.AreEqual("stale COM object", error.Message);
            CollectionAssert.AreEqual(new[] { 1 }, visited);
        }

        private static string JsonSchema(string title)
        {
            return DocumentAutomationSchemaSerializer.Serialize(new DocumentAutomationSchema { Title = title });
        }

        private sealed class Node
        {
            public string Name { get; private set; }
            public List<Node> Children { get; private set; } = new List<Node>();
            public static Node Leaf(string name) => new Node { Name = name };
            public static Node Group(params Node[] children) => new Node { Children = children.ToList() };
        }
    }
}
