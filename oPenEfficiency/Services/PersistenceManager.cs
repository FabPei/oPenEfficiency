using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Serialization;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Handles the serialization and persistence of the Gantt project data
    /// directly into the tags of the GanttChart_Project group shape, allowing
    /// for easy copy-pasting of charts across slides and presentations.
    /// </summary>
    public class PersistenceManager
    {
        public const string GANTT_DATA_TAG = "oPE_GanttData";
        private const string LegacyGanttDataTag = "GANTT_DATA";
        private readonly PowerPointManager _pptManager;

        public PersistenceManager(PowerPointManager pptManager)
        {
            _pptManager = pptManager;
        }

        /// <summary>
        /// Saves the Gantt project into the tags of the newly rendered group shape.
        /// </summary>
        public void SaveProject(GanttProject project, Shape groupShape)
        {
            if (project == null || groupShape == null) return;

            try
            {
                PowerPointShapeManager.SetChartId(groupShape, project.ChartId, true);
                var serializer = new XmlSerializer(typeof(GanttProject));
                string xmlData = string.Empty;

                using (var stringWriter = new StringWriter())
                {
                    serializer.Serialize(stringWriter, project);
                    xmlData = stringWriter.ToString();
                }

                // Compress string by converting to Base64 to avoid tag string parsing issues with raw XML
                var bytes = System.Text.Encoding.UTF8.GetBytes(xmlData);
                string b64 = Convert.ToBase64String(bytes);

                groupShape.Tags.Add(GANTT_DATA_TAG, b64);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PersistenceManager.SaveProject");
            }
        }

        /// <summary>
        /// Loads the Gantt project from a selected shape, or falls back to finding it on the current slide.
        /// Returns a new empty project if no data is found.
        /// </summary>
        public GanttProject LoadProject(Shape selectedShape)
        {
            // Try loading from selected shape
            if (selectedShape != null)
            {
                GanttProject proj = TryLoadFromShape(selectedShape);
                if (proj != null) return PrepareLoadedProject(proj, selectedShape, true);

                // If they selected an inner shape, try to find its parent group
                Shape parentGroup = null;
                try
                {
                    parentGroup = selectedShape.ParentGroup;
                    if (parentGroup != null)
                    {
                        proj = TryLoadFromShape(parentGroup);
                        if (proj != null) return PrepareLoadedProject(proj, parentGroup, true);
                    }
                }
                catch { /* Throws if no parent group exists */ }
                finally
                {
                    if (parentGroup != null) Marshal.ReleaseComObject(parentGroup);
                }
            }

            // Fallback: Check the entire slide for the GanttChart_Project group
            var slide = _pptManager.GetCurrentSlide();
            if (slide != null)
            {
                Shapes shapes = null;
                try
                {
                    shapes = slide.Shapes;
                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        Shape shape = null;
                        try
                        {
                            shape = shapes[i];
                            if (PowerPointShapeManager.HasGanttProjectTag(shape))
                            {
                                GanttProject proj = TryLoadFromShape(shape);
                                if (proj != null) return PrepareLoadedProject(proj, shape, false);
                            }
                        }
                        finally
                        {
                            if (shape != null) Marshal.ReleaseComObject(shape);
                        }
                    }
                }
                catch (COMException)
                {
                    // A stale slide or shape collection falls back to a new project.
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "PersistenceManager.LoadProject.ScanSlide");
                }
                finally
                {
                    if (shapes != null) Marshal.ReleaseComObject(shapes);
                }
            }

            return GanttProject.CreateDefault();
        }

        private GanttProject TryLoadFromShape(Shape shape)
        {
            try
            {
                string b64 = shape.Tags[GANTT_DATA_TAG];
                if (string.IsNullOrEmpty(b64))
                {
                    b64 = shape.Tags[LegacyGanttDataTag];
                }
                if (!string.IsNullOrEmpty(b64))
                {
                    var bytes = Convert.FromBase64String(b64);
                    string xmlData = System.Text.Encoding.UTF8.GetString(bytes);

                    var serializer = new XmlSerializer(typeof(GanttProject));
                    using (var stringReader = new StringReader(xmlData))
                    {
                        var project = (GanttProject)serializer.Deserialize(stringReader);
                        if (project != null)
                        {
                            var tasks = project.Tasks;
                            project.Tasks = null;
                            project.Tasks = tasks;
                            return project;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PersistenceManager.TryLoadFromShape");
            }
            return null;
        }

        private GanttProject PrepareLoadedProject(GanttProject project, Shape projectShape, bool resolveCopiedIdentity)
        {
            if (project == null || projectShape == null) return project;

            Guid taggedChartId;
            if (Guid.TryParse(PowerPointShapeManager.GetChartId(projectShape), out taggedChartId) && taggedChartId != Guid.Empty)
            {
                project.ChartId = taggedChartId;
            }

            if (resolveCopiedIdentity && HasDuplicateChartIdentity(projectShape, project.ChartId))
            {
                project.ChartId = Guid.NewGuid();
            }

            PowerPointShapeManager.SetChartId(projectShape, project.ChartId, true);
            return project;
        }

        private bool HasDuplicateChartIdentity(Shape selectedProjectShape, Guid chartId)
        {
            Slide slide = _pptManager.GetCurrentSlide();
            if (slide == null || selectedProjectShape == null || chartId == Guid.Empty) return false;

            Shapes shapes = null;
            int matches = 0;
            try
            {
                shapes = slide.Shapes;
                string chartIdText = chartId.ToString();
                for (int i = 1; i <= shapes.Count; i++)
                {
                    Shape shape = null;
                    try
                    {
                        shape = shapes[i];
                        if (PowerPointShapeManager.HasGanttProjectTag(shape) &&
                            PowerPointShapeManager.GetChartId(shape) == chartIdText)
                        {
                            matches++;
                            if (matches > 1) return true;
                        }
                    }
                    finally
                    {
                        if (shape != null) Marshal.ReleaseComObject(shape);
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PersistenceManager.HasDuplicateChartIdentity");
            }
            finally
            {
                if (shapes != null) Marshal.ReleaseComObject(shapes);
            }

            return false;
        }
    }
}
