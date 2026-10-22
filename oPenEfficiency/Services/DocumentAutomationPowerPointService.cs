using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.DocumentAutomation;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    /// <summary>PowerPoint boundary for Build Presentation from Answers. Planning remains host-neutral in Core.</summary>
    public sealed class DocumentAutomationPowerPointService
    {
        public const string StoreId = "oPE_DocumentAutomation_v1";
        public const string SlideConditionTag = "oPE_DocumentAutomationCondition";
        private readonly PowerPointManager _manager;

        public DocumentAutomationPowerPointService(PowerPointManager manager)
        {
            _manager = manager ?? throw new ArgumentNullException(nameof(manager));
        }

        public Presentation ActivePresentation => _manager.GetApplication()?.ActivePresentation;

        public DocumentAutomationSchema LoadSchema()
        {
            var presentation = ActivePresentation;
            if (presentation == null) return null;
            return DocumentAutomationSchemaRecovery.DeserializeNewestValid(DataStoreManager.LoadDataCandidates(presentation, StoreId));
        }

        public void SaveSchema(DocumentAutomationSchema schema)
        {
            var presentation = ActivePresentation ?? throw new InvalidOperationException("No active presentation is available.");
            var slideIds = presentation.Slides.Cast<Slide>().Select(slide => slide.SlideID).ToList();
            var issues = new DocumentAutomationValidator().Validate(schema, slideIds);
            if (issues.Count > 0) throw new InvalidDataException(string.Join(Environment.NewLine, issues.Select(issue => issue.Message)));
            DataStoreManager.SaveData(presentation, StoreId, DocumentAutomationSchemaSerializer.Serialize(schema));
            SynchronizeConditionTags(presentation, schema);
        }

        public IList<AutomationSlideSnapshot> CaptureSlides(Presentation presentation = null)
        {
            presentation = presentation ?? ActivePresentation;
            var result = new List<AutomationSlideSnapshot>();
            if (presentation == null) return result;
            OneBasedCollectionTraversal.Forward(presentation.Slides.Count, slideIndex =>
            {
                Slide slide = presentation.Slides[slideIndex];
                var snapshot = new AutomationSlideSnapshot { SlideId = slide.SlideID, SlideNumber = slide.SlideIndex };
                OneBasedCollectionTraversal.Forward(slide.Shapes.Count, shapeIndex => CaptureShapeText(slide.Shapes[shapeIndex], snapshot.TextFragments));
                result.Add(snapshot);
            });
            return result;
        }

        public AutomationExecutionPlan BuildPlan(DocumentAutomationSchema schema, IReadOnlyDictionary<string, string> values)
        {
            return new DocumentAutomationPlanner().Build(schema, CaptureSlides(), values);
        }

        public Presentation ExecuteToNewPresentation(DocumentAutomationSchema schema, IReadOnlyDictionary<string, string> values, string outputPath)
        {
            var source = ActivePresentation ?? throw new InvalidOperationException("No active presentation is available.");
            if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
            outputPath = Path.GetFullPath(outputPath);
            if (!string.IsNullOrWhiteSpace(source.FullName) && string.Equals(Path.GetFullPath(source.FullName), outputPath, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The output must not overwrite the source presentation.");

            // Complete validation and planning before creating any file.
            AutomationExecutionPlan plan = BuildPlan(schema, values);
            if (!plan.CanExecute)
                throw new InvalidOperationException("The automation preview contains validation errors or missing values.");

            var transaction = new DocumentAutomationOutputTransaction<Presentation>(cleanupFailureReporter: ExceptionLogger.Log);
            return transaction.Execute(outputPath, new PowerPointOutputOperations(_manager, source, plan, values));
        }

        private static void ApplyPlan(Presentation presentation, AutomationExecutionPlan plan, IReadOnlyDictionary<string, string> values)
        {
            var excludedIds = new HashSet<int>(plan.SlideDecisions.Where(decision => !decision.Include).Select(decision => decision.SlideId));
            OneBasedCollectionTraversal.Reverse(presentation.Slides.Count, slideIndex =>
            {
                if (excludedIds.Contains(presentation.Slides[slideIndex].SlideID)) presentation.Slides[slideIndex].Delete();
            });

            OneBasedCollectionTraversal.Forward(presentation.Slides.Count, slideIndex =>
            {
                Slide slide = presentation.Slides[slideIndex];
                OneBasedCollectionTraversal.Forward(slide.Shapes.Count, shapeIndex => ReplaceShapeText(slide.Shapes[shapeIndex], values));
            });
        }

        private sealed class PowerPointOutputOperations : IDocumentAutomationOutputOperations<Presentation>
        {
            private readonly PowerPointManager _manager;
            private readonly Presentation _source;
            private readonly AutomationExecutionPlan _plan;
            private readonly IReadOnlyDictionary<string, string> _values;

            public PowerPointOutputOperations(PowerPointManager manager, Presentation source, AutomationExecutionPlan plan, IReadOnlyDictionary<string, string> values)
            {
                _manager = manager;
                _source = source;
                _plan = plan;
                _values = values;
            }

            public void CreateWorkingCopy(string temporaryPath)
            {
                _source.SaveCopyAs(temporaryPath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Office.MsoTriState.msoFalse);
            }

            public void ApplyAndSaveWorkingCopy(string temporaryPath)
            {
                Presentation workingCopy = null;
                try
                {
                    workingCopy = _manager.GetApplication().Presentations.Open(temporaryPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse);
                    ApplyPlan(workingCopy, _plan, _values);
                    workingCopy.Save();
                    workingCopy.Close();
                    workingCopy = null;
                }
                finally
                {
                    if (workingCopy != null)
                        ExceptionPolicy.RunBestEffortComCleanup(() =>
                        {
                            workingCopy.Saved = Office.MsoTriState.msoTrue;
                            workingCopy.Close();
                        }, "DocumentAutomationPowerPointService.CloseFailedCopy");
                }
            }

            public Presentation OpenCompletedOutput(string outputPath)
            {
                return _manager.GetApplication().Presentations.Open(outputPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue);
            }
        }

        private static void CaptureShapeText(Shape shape, IList<string> texts)
        {
            if (shape == null) return;
            OneBasedCollectionTraversal.VisitLeaves(
                shape,
                item => item.Type == Office.MsoShapeType.msoGroup ? item.GroupItems.Count : 0,
                (item, index) => item.GroupItems[index],
                item => CaptureLeafText(item, texts));
        }

        private static void CaptureLeafText(Shape shape, IList<string> texts)
        {
            if (shape.HasTable == Office.MsoTriState.msoTrue)
            {
                OneBasedCollectionTraversal.Forward(shape.Table.Rows.Count, row =>
                    OneBasedCollectionTraversal.Forward(shape.Table.Columns.Count, column =>
                        texts.Add(shape.Table.Cell(row, column).Shape.TextFrame.TextRange.Text ?? string.Empty)));
            }
            else if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                texts.Add(shape.TextFrame.TextRange.Text ?? string.Empty);
        }

        private static void ReplaceShapeText(Shape shape, IReadOnlyDictionary<string, string> values)
        {
            if (shape == null) return;
            OneBasedCollectionTraversal.VisitLeaves(
                shape,
                item => item.Type == Office.MsoShapeType.msoGroup ? item.GroupItems.Count : 0,
                (item, index) => item.GroupItems[index],
                item => ReplaceLeafText(item, values));
        }

        private static void ReplaceLeafText(Shape shape, IReadOnlyDictionary<string, string> values)
        {
            if (shape.HasTable == Office.MsoTriState.msoTrue)
            {
                OneBasedCollectionTraversal.Forward(shape.Table.Rows.Count, row =>
                    OneBasedCollectionTraversal.Forward(shape.Table.Columns.Count, column =>
                        ReplaceTextRange(shape.Table.Cell(row, column).Shape.TextFrame.TextRange, values)));
            }
            else if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                ReplaceTextRange(shape.TextFrame.TextRange, values);
        }

        private static void ReplaceTextRange(TextRange range, IReadOnlyDictionary<string, string> values)
        {
            foreach (PlaceholderReplacement replacement in PlaceholderEngine.PlanReplacements(range.Text ?? string.Empty, values))
            {
                // Replace from the end so earlier character offsets remain valid and surrounding formatting is retained.
                range.Characters(replacement.Start + 1, replacement.Length).Text = replacement.Value;
            }
        }

        private static void SynchronizeConditionTags(Presentation presentation, DocumentAutomationSchema schema)
        {
            var conditions = (schema.SlideConditions ?? new List<AutomationSlideCondition>()).ToDictionary(condition => condition.SlideId);
            OneBasedCollectionTraversal.Forward(presentation.Slides.Count, slideIndex =>
            {
                Slide slide = presentation.Slides[slideIndex];
                ExceptionPolicy.TryComProbe(() => { slide.Tags.Delete(SlideConditionTag); return true; }, out bool _);
                if (conditions.TryGetValue(slide.SlideID, out AutomationSlideCondition condition))
                    slide.Tags.Add(SlideConditionTag, condition.QuestionKey + "|" + condition.Operator + "|" + condition.ExpectedValue);
            });
        }
    }
}
