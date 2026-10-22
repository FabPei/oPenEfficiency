using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.DocumentAutomation;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public class Program
    {
        [TestMethod] public void ComProbeReturnsValues() => ProbeReturnsValue();
        [TestMethod] public void ComProbeSuppressesExpectedFailuresOnly() => ProbeSuppressesOnlyExpectedComFailures();
        [TestMethod] public void CleanupLogsUnexpectedFailures() => CleanupSuppressesExpectedAndLogsUnexpectedFailures();
        [TestMethod] public void OptionalComFailuresAreAccepted() => OptionalFailureReporterAcceptsExpectedComFailures();
        [TestMethod] public void InvalidExceptionContractsFailFast() => InvalidContractsFailFast();
        [TestMethod] public void HtmlFixtureAndResourcesAreParsed() => HtmlPipelineParsesFixturesAndResolvesResources();
        [TestMethod] public void HtmlColumnsAndGridArePlanned() => HtmlLayoutPlansColumnsAndGrid();
        [TestMethod] public void HtmlRendererKindsAreSeparated() => HtmlRenderRoutingSeparatesSupportedKinds();
        [TestMethod] public void AgendaPreviewAndValidationAreStable() => AgendaPlanningBuildsStablePreviewAndValidation();
        [TestMethod] public void AgendaInfrastructureIsLimitedToVisibleChapters() => AgendaInfrastructurePolicyIsStable();
        [TestMethod] public void AgendaXmlImportAcceptsSupportedNames() => AgendaImportFileNamesAreStable();
        [TestMethod] public void AgendaProtectionPolicyBlocksOnlyMutatingPaths() => AgendaProtectionPolicyIsStable();
        [TestMethod] public void IconSearchAndCacheAreDeterministic() => IconSearchAndManifestCacheRemainDeterministic();
        [TestMethod] public void TableDimensionsScaleToRequestedTotal() => TableDimensionScalingIsStable();
        [TestMethod] public void DirectionalGuidesNeverReverseAlignmentDirection() => DirectionalGuideSelectionIsStable();
        [TestMethod] public void DocumentAutomationJsonRoundTripIsVersioned() => DocumentAutomationJsonRoundTrip();
        [TestMethod] public void DocumentAutomationValidationRejectsBadKeysAndRules() => DocumentAutomationValidation();
        [TestMethod] public void DocumentAutomationPlaceholdersAreDetectedAndReplaced() => DocumentAutomationPlaceholders();
        [TestMethod] public void DocumentAutomationConditionsAndPlanAreDeterministic() => DocumentAutomationPlanning();

        [DataTestMethod]
        [DataRow("p", null, HtmlNodeRendererKind.Text)]
        [DataRow("ol", null, HtmlNodeRendererKind.List)]
        [DataRow("img", null, HtmlNodeRendererKind.Image)]
        [DataRow("div", "grid", HtmlNodeRendererKind.Grid)]
        [DataRow("div", "flex", HtmlNodeRendererKind.Flex)]
        public void HtmlRendererRoutingIsDataDriven(string tag, string display, HtmlNodeRendererKind expected)
        {
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(
                expected,
                new HtmlRenderRouter().GetNodeRenderer(tag, display));
        }

        [DataTestMethod]
        [DataRow("", 0)]
        [DataRow("<p>no slide</p>", 0)]
        [DataRow("<section class='slide'><h2>One</h2></section>", 1)]
        [DataRow("<article role='slide'><h1>Fallback</h1></article>", 1)]
        public void HtmlParserHandlesBoundaryDocuments(string html, int expectedSlides)
        {
            HtmlImportDocument document = new HtmlImportDocumentPipeline(false).Parse(html, Path.GetTempPath());
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(expectedSlides, document.Slides.Count);
            if (expectedSlides == 0 && html.Length > 0)
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(document.Warnings.Count > 0);
        }

        [TestMethod]
        public void AgendaValidationCoversHierarchyAndInvalidDurations()
        {
            var items = new List<AgendaPlanItem>
            {
                new AgendaPlanItem { Topic = "", Level = 1, Duration = 0, ShouldGenerate = true },
                new AgendaPlanItem { Topic = "Jump", Level = 3, Duration = 10, ShouldGenerate = true }
            };
            AgendaPlanResult result = new AgendaPlanningService().Build(items, TimeSpan.Zero, 1, "HH:mm", null);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(result.Issues.Any(issue => issue.Contains("no name")));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(result.Issues.Any(issue => issue.Contains("positive duration")));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(result.Issues.Any(issue => issue.Contains("first topic")));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(result.Issues.Any(issue => issue.Contains("skips")));
        }

        [TestMethod]
        public void AgendaDurationFittingPreservesBreaksAndHiddenItems()
        {
            var items = new List<AgendaPlanItem>
            {
                new AgendaPlanItem { Topic = "Opening", Duration = 30, ShouldGenerate = true },
                new AgendaPlanItem { Topic = "Break", Duration = 10, IsBreak = true, ShouldGenerate = true },
                new AgendaPlanItem { Topic = "Workshop", Duration = 60, ShouldGenerate = true },
                new AgendaPlanItem { Topic = "Hidden", Duration = 99, ShouldGenerate = false }
            };

            bool fitted = new AgendaPlanningService().TryFitDurations(
                items,
                TimeSpan.FromHours(9),
                new TimeSpan(10, 10, 0));

            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(fitted);
            CollectionAssert.AreEqual(new[] { 20, 10, 40, 99 }, items.Select(item => item.Duration).ToArray());
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(
                70,
                items.Where(item => item.ShouldGenerate).Sum(item => item.Duration));
        }

        [TestMethod]
        public void AgendaDurationFittingHandlesOvernightWindows()
        {
            var items = new List<AgendaPlanItem>
            {
                new AgendaPlanItem { Topic = "Late", Duration = 10, ShouldGenerate = true },
                new AgendaPlanItem { Topic = "Later", Duration = 20, ShouldGenerate = true }
            };

            bool fitted = new AgendaPlanningService().TryFitDurations(
                items,
                new TimeSpan(23, 30, 0),
                new TimeSpan(0, 30, 0));

            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(fitted);
            CollectionAssert.AreEqual(new[] { 20, 40 }, items.Select(item => item.Duration).ToArray());
        }

        [TestMethod]
        public void AgendaDurationFittingRejectsImpossibleWindowsWithoutMutation()
        {
            var items = new List<AgendaPlanItem>
            {
                new AgendaPlanItem { Topic = "One", Duration = 5, ShouldGenerate = true },
                new AgendaPlanItem { Topic = "Break", Duration = 10, IsBreak = true, ShouldGenerate = true },
                new AgendaPlanItem { Topic = "Two", Duration = 15, ShouldGenerate = true }
            };
            int[] originalDurations = items.Select(item => item.Duration).ToArray();

            bool fitted = new AgendaPlanningService().TryFitDurations(
                items,
                TimeSpan.FromHours(9),
                new TimeSpan(9, 11, 0));

            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(fitted);
            CollectionAssert.AreEqual(originalDurations, items.Select(item => item.Duration).ToArray());
        }

        [TestMethod]
        public void IconManifestParserHandlesEmptyAndRoundTripData()
        {
            var parser = new IconManifestParser();
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, parser.Parse(null).Icons.Count);
            var source = new IconCollectionManifest { Icons = new List<IconCatalogItem> { new IconCatalogItem { Id = "one", Name = "One" } } };
            IconCollectionManifest clone = parser.Clone(source);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("one", clone.Icons.Single().Id);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreNotSame(source, clone);
        }

        private static void TableDimensionScalingIsStable()
        {
            float[] scaled = TableDimensionScaler.ScaleToTotal(
                new[] { 40f, 20f, 40f, 20f },
                90f);

            Assert(scaled.Length == 4, "The scaled dimension count changed.");
            Assert(Math.Abs(scaled.Sum() - 90f) < 0.001f,
                "Scaled dimensions do not match the requested total.");
            Assert(Math.Abs(scaled[0] / scaled[1] - 2f) < 0.001f &&
                   Math.Abs(scaled[2] / scaled[3] - 2f) < 0.001f,
                "Scaled dimensions did not preserve their proportions.");

            AssertThrows<ArgumentException>(
                () => TableDimensionScaler.ScaleToTotal(Array.Empty<float>(), 100f),
                "Empty table dimensions were accepted.");
            AssertThrows<ArgumentOutOfRangeException>(
                () => TableDimensionScaler.ScaleToTotal(new[] { 20f, 0f }, 100f),
                "A non-positive table dimension was accepted.");
        }

        private static void AgendaInfrastructurePolicyIsStable()
        {
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(
                AgendaWorkflowPolicy.ShouldCreateChapterInfrastructure(0, false, true));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(
                AgendaWorkflowPolicy.ShouldCreateChapterInfrastructure(1, false, true));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(
                AgendaWorkflowPolicy.ShouldCreateChapterInfrastructure(0, true, true));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(
                AgendaWorkflowPolicy.ShouldCreateChapterInfrastructure(0, false, false));
        }

        private static void AgendaImportFileNamesAreStable()
        {
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(
                AgendaWorkflowPolicy.IsSupportedExternalLayoutFile(@"C:\\Layouts\\AgendaCreator.xml"));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(
                AgendaWorkflowPolicy.IsSupportedExternalLayoutFile(@"C:\\Layouts\\agendawizard.XML"));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(
                AgendaWorkflowPolicy.IsSupportedExternalLayoutFile(@"C:\\Layouts\\layout.xml"));
        }

        private static void AgendaProtectionPolicyIsStable()
        {
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNotNull(
                AgendaWorkflowPolicy.GetMutationProtectionError(true, false, false, false));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNotNull(
                AgendaWorkflowPolicy.GetMutationProtectionError(false, true, true, false));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNotNull(
                AgendaWorkflowPolicy.GetMutationProtectionError(false, true, false, true));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNull(
                AgendaWorkflowPolicy.GetMutationProtectionError(false, true, false, false));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNull(
                AgendaWorkflowPolicy.GetMutationProtectionError(false, false, true, true));
        }

        private static void DirectionalGuideSelectionIsStable()
        {
            var guides = new[] { 0f, 480f, 720f };

            Assert(DirectionalGuideSelector.NearestAtOrBefore(guides, 470f, 0f) == 0f,
                "Align Left moved a left edge rightward to the slide center.");
            Assert(DirectionalGuideSelector.NearestAtOrBefore(guides, 500f, 0f) == 480f,
                "Align Left did not choose the nearest guide to the left.");
            Assert(DirectionalGuideSelector.NearestAtOrAfter(guides, 490f, 960f) == 720f,
                "Align Right did not choose the nearest guide to the right.");
            Assert(DirectionalGuideSelector.NearestAtOrAfter(guides, 730f, 960f) == 960f,
                "Align Right did not fall back to the right slide edge.");
        }

        private static void DocumentAutomationJsonRoundTrip()
        {
            var schema = SampleAutomationSchema();
            string json = DocumentAutomationSchemaSerializer.Serialize(schema);
            var clone = DocumentAutomationSchemaSerializer.Deserialize(json);
            Assert(clone.SchemaVersion == 1 && clone.Questions.Count == 3 && clone.SlideConditions.Single().SlideId == 20, "Document automation JSON did not round-trip.");
            AssertThrows<NotSupportedException>(() => DocumentAutomationSchemaSerializer.Deserialize(json.Replace("\"SchemaVersion\":1", "\"SchemaVersion\":2")), "A newer schema version was accepted.");
        }

        private static void DocumentAutomationValidation()
        {
            var schema = SampleAutomationSchema();
            schema.Questions.Add(new AutomationQuestion { Key = "ClientName", Prompt = "Duplicate", Type = AutomationQuestionType.Text });
            schema.Questions.Add(new AutomationQuestion { Key = "9 bad", Prompt = "Bad", Type = AutomationQuestionType.Text });
            schema.SlideConditions.Add(new AutomationSlideCondition { SlideId = 20, QuestionKey = "ClientName", ExpectedValue = "x" });
            var issues = new DocumentAutomationValidator().Validate(schema, new[] { 10, 20 });
            Assert(issues.Any(i => i.Code == "DuplicateKey") && issues.Any(i => i.Code == "InvalidKey") && issues.Any(i => i.Code == "NestedCondition"), "Schema validation missed invalid or duplicate keys or complex conditions.");
        }

        private static void DocumentAutomationPlaceholders()
        {
            var keys = PlaceholderEngine.Detect(new[] { "Hello {{ClientName}}", "{{ ClientName }} / {{Scope}}" });
            Assert(keys.SequenceEqual(new[] { "ClientName", "Scope" }), "Placeholder detection was not stable.");
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["clientname"] = "Acme" };
            Assert(PlaceholderEngine.Replace("Hello {{ClientName}} — {{Scope}}", values) == "Hello Acme — {{Scope}}", "Placeholder replacement changed missing values.");
            Assert(PlaceholderEngine.Missing(keys, values).SequenceEqual(new[] { "Scope" }), "Missing placeholder values were not reported.");
        }

        private static void DocumentAutomationPlanning()
        {
            var slides = new[]
            {
                new AutomationSlideSnapshot { SlideId = 10, SlideNumber = 1, TextFragments = new[] { "{{ClientName}}" } },
                new AutomationSlideSnapshot { SlideId = 20, SlideNumber = 2, TextFragments = new[] { "{{Scope}}" } }
            };
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["ClientName"] = "Acme", ["IncludeAppendix"] = "No", ["Scope"] = "Audit" };
            var plan = new DocumentAutomationPlanner().Build(SampleAutomationSchema(), slides, values);
            Assert(plan.CanExecute && plan.SlideDecisions[0].Include && !plan.SlideDecisions[1].Include, "Execution planning did not apply a Yes/No condition.");
            values.Remove("Scope");
            Assert(new DocumentAutomationPlanner().Build(SampleAutomationSchema(), slides, values).MissingValueKeys.Single() == "Scope", "Execution planning missed a required placeholder value.");
        }

        private static DocumentAutomationSchema SampleAutomationSchema()
        {
            return new DocumentAutomationSchema
            {
                Questions = new List<AutomationQuestion>
                {
                    new AutomationQuestion { Key = "ClientName", Prompt = "Client name", Type = AutomationQuestionType.Text, Required = true },
                    new AutomationQuestion { Key = "IncludeAppendix", Prompt = "Include appendix?", Type = AutomationQuestionType.YesNo, Required = true },
                    new AutomationQuestion { Key = "Scope", Prompt = "Scope", Type = AutomationQuestionType.SingleChoice, Options = new List<string> { "Audit", "Implementation" } }
                },
                SlideConditions = new List<AutomationSlideCondition>
                {
                    new AutomationSlideCondition { SlideId = 20, QuestionKey = "IncludeAppendix", Operator = AutomationConditionOperator.Equals, ExpectedValue = "Yes" }
                }
            };
        }

        private static void ProbeReturnsValue()
        {
            Assert(ExceptionPolicy.TryComProbe(() => 42, out int value) && value == 42,
                "A successful COM probe did not preserve its value.");
        }

        private static void ProbeSuppressesOnlyExpectedComFailures()
        {
            bool succeeded = ExceptionPolicy.TryComProbe<int>(
                () => throw new COMException("Expected probe failure"),
                out int value);
            Assert(!succeeded && value == 0,
                "An expected COM probe failure was not converted to a negative result.");

            AssertThrows<InvalidOperationException>(
                () => ExceptionPolicy.TryComProbe<int>(
                    () => throw new InvalidOperationException("Unexpected"),
                    out _),
                "A non-COM probe failure was silently swallowed.");
        }

        private static void CleanupSuppressesExpectedAndLogsUnexpectedFailures()
        {
            int logs = 0;
            ExceptionPolicy.RunBestEffort(
                () => throw new COMException("Expected cleanup failure"),
                "ExpectedCleanup",
                ExceptionPolicy.IsExpectedComFailure,
                (exception, context) => logs++);
            Assert(logs == 0, "Expected COM cleanup was logged as an unexpected failure.");

            ExceptionPolicy.RunBestEffort(
                () => throw new InvalidOperationException("Unexpected cleanup failure"),
                "UnexpectedCleanup",
                ExceptionPolicy.IsExpectedComFailure,
                (exception, context) =>
                {
                    Assert(context == "UnexpectedCleanup", "Cleanup context was not preserved.");
                    logs++;
                });
            Assert(logs == 1, "Unexpected cleanup was not reported exactly once.");
        }

        private static void InvalidContractsFailFast()
        {
            AssertThrows<ArgumentNullException>(
                () => ExceptionPolicy.TryComProbe<int>(null, out _),
                "A null probe was accepted.");
            AssertThrows<ArgumentException>(
                () => ExceptionPolicy.RunBestEffort(
                    () => { },
                    string.Empty,
                    ExceptionPolicy.IsExpectedComFailure,
                    (exception, context) => { }),
                "Cleanup without diagnostic context was accepted.");
            AssertThrows<ArgumentNullException>(
                () => ExceptionPolicy.ReportOptionalComFailure(null, "Probe"),
                "A null optional failure was accepted.");
            AssertThrows<ArgumentException>(
                () => ExceptionPolicy.ReportOptionalComFailure(new COMException(), ""),
                "An optional failure without diagnostic context was accepted.");
        }

        private static void OptionalFailureReporterAcceptsExpectedComFailures()
        {
            ExceptionPolicy.ReportOptionalComFailure(
                new COMException("Expected optional COM failure"),
                "OptionalProbe");
        }

        private static void HtmlPipelineParsesFixturesAndResolvesResources()
        {
            string root = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fixtures", "html");
            HtmlImportDocument document = new HtmlImportDocumentPipeline(false).Load(Path.Combine(root, "basic-deck.html"));
            Assert(document.Slides.Count == 1 && document.Slides[0].Title == "Hello", "HTML fixture did not produce the expected slide model.");
            Assert(document.Slides[0].FooterRight == "Right" && document.Slides[0].StyleSheet.Contains("nested.css") == false, "CSS imports or footer parsing changed.");
            Assert(HtmlResourceLoader.ResolveReference("images/a.svg", root) == Path.Combine(root, "images", "a.svg"), "Relative resource resolution changed.");
        }

        private static void HtmlLayoutPlansColumnsAndGrid()
        {
            var engine = new HtmlLayoutEngine();
            var columns = engine.Columns(new HtmlLayoutRect(10, 20, 220, 100), 2, 20);
            Assert(columns.Count == 2 && columns[0].Width == 100 && columns[1].Left == 130, "Column planning is not deterministic.");
            var grid = engine.Grid(new HtmlLayoutRect(0, 0, 210, 110), 4, 2, 10, 10);
            Assert(grid.Count == 4 && grid[3].Left == 110 && grid[3].Top == 60, "Grid planning is not deterministic.");
        }

        private static void HtmlRenderRoutingSeparatesSupportedKinds()
        {
            var router = new HtmlRenderRouter();
            Assert(router.GetNodeRenderer("p", null) == HtmlNodeRendererKind.Text, "Text renderer routing changed.");
            Assert(router.GetNodeRenderer("ol", null) == HtmlNodeRendererKind.List, "List renderer routing changed.");
            Assert(router.GetNodeRenderer("img", null) == HtmlNodeRendererKind.Image, "Image renderer routing changed.");
            Assert(router.GetNodeRenderer("div", "grid") == HtmlNodeRendererKind.Grid, "Grid renderer routing changed.");
            Assert(router.GetNodeRenderer("div", "flex") == HtmlNodeRendererKind.Flex, "Flex renderer routing changed.");
            Assert(router.GetSlideRenderer(new HtmlImportSlideModel { IsGptDeck=true }) == HtmlSlideRendererKind.GptDeck, "GPT deck renderer routing changed.");
        }

        private static void AgendaPlanningBuildsStablePreviewAndValidation()
        {
            var items = new List<AgendaPlanItem>
            {
                new AgendaPlanItem { Topic="Opening", Level=0, Duration=30, ShouldGenerate=true },
                new AgendaPlanItem { Topic="Detail", Level=1, Duration=20, ShouldGenerate=true },
                new AgendaPlanItem { Topic="Detail", Level=1, Duration=20, ShouldGenerate=true }
            };
            AgendaPlanResult result = new AgendaPlanningService().Build(items, TimeSpan.FromHours(9), 1, "HH:mm", TimeSpan.FromHours(10));
            Assert(items[0].DisplayIndex == "1" && items[1].DisplayIndex == "1.1", "Agenda numbering changed.");
            Assert(items[0].TimeSlot == "09:00 - 09:30" && result.Preview.Count == 3, "Agenda preview timing changed.");
            Assert(result.Issues.Any(issue => issue.Contains("duplicated")) && result.Issues.Any(issue => issue.Contains("10 minutes")), "Agenda validation missed duplicate or budget issues.");
        }

        private static void IconSearchAndManifestCacheRemainDeterministic()
        {
            var icons = new[]
            {
                new IconCatalogItem { Id="z", Name="Zulu", Style="outline", Tags=new List<string>{"arrow"} },
                new IconCatalogItem { Id="a", Name="Alpha", Style="filled", Tags=new List<string>{"arrow"}, IsFavorite=true }
            };
            var results = new IconSearchIndex(icons).Search("arrow");
            Assert(results.Select(icon => icon.Name).SequenceEqual(new[]{"Alpha","Zulu"}), "Icon search order changed.");
            Assert(new IconSearchIndex(icons).Search(favoritesOnly:true).Single().Id == "a", "Favorite filtering changed.");

            int loads=0; var cache=new IconManifestCache();
            Func<string,IconCollectionManifest> loader = path => { loads++; return new IconCollectionManifest { Version=loads.ToString() }; };
            cache.Get("manifest.json", new DateTime(2026,1,1), loader); cache.Get("manifest.json", new DateTime(2026,1,1), loader);
            cache.Get("manifest.json", new DateTime(2026,1,2), loader);
            Assert(loads == 2, "Manifest cache did not invalidate on write-time change.");
        }

        private static void AssertThrows<TException>(Action action, string message)
            where TException : Exception
        {
            try
            {
                action();
            }
            catch (TException)
            {
                return;
            }

            throw new InvalidOperationException(message);
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }
    }
}
