using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency
{
    public class AgendaService
    {
        private readonly Application _application;
        private const string AgendaSlideTag = "oPE_AgendaSlide";

        public AgendaService(Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public bool CreateOrUpdateAgenda()
        {
            try
            {
                var pres = _application.ActivePresentation;
                dynamic sections = pres.SectionProperties;

                if (sections.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show("No sections found in this presentation.");
                    return false;
                }

                // 1. Collect all section names
                List<string> sectionNames = new List<string>();
                for (int i = 1; i <= sections.Count; i++)
                {
                    sectionNames.Add(sections.Name(i));
                }

                // Check presentation level Think-Cell tag
                bool isPresThinkCellProtected = false;
                try 
                {
                    isPresThinkCellProtected = !string.IsNullOrEmpty(pres.Tags["THINKCELLPRESENTATIONDONOTDELETE"]);
                } 
                catch { }

                // 2. Iterate through sections to create/update slides
                // We go backwards to avoid index shifting issues if we were deleting, 
                // but since we insert at specific section start, forwards is fine too.
                for (int i = 1; i <= sections.Count; i++)
                {
                    string currentSectionName = sections.Name(i);
                    int insertIndex = sections.FirstSlideIndex(i);

                    Slide agendaSlide = FindExistingAgendaSlideInSection(pres, i);

                    if (isPresThinkCellProtected || IsThinkCellProtected(agendaSlide))
                    {
                        System.Diagnostics.Debug.WriteLine($"Skipping Think-Cell protected section '{currentSectionName}'");
                        continue;
                    }

                    if (agendaSlide == null)
                    {
                        // Create new slide at the start of the section
                        agendaSlide = pres.Slides.Add(insertIndex, PpSlideLayout.ppLayoutText);
                        agendaSlide.Tags.Add(AgendaSlideTag, "True");
                    }

                    UpdateAgendaSlideContent(agendaSlide, sectionNames, currentSectionName);
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Agenda Creator Error: {ex.Message}");
                return false;
            }
        }

        private bool IsThinkCellProtected(Slide slide)
        {
            if (slide == null) return false;

            try
            {
                if (!string.IsNullOrEmpty(slide.Tags["THINKCELLSHAPEDONOTDELETE"]))
                {
                    return true;
                }

                foreach (Shape shape in slide.Shapes)
                {
                    if (!string.IsNullOrEmpty(shape.Tags["THINKCELLSHAPEDONOTDELETE"]))
                    {
                        return true;
                    }
                }
            }
            catch { }

            return false;
        }

        private Slide FindExistingAgendaSlideInSection(Presentation pres, int sectionIndex)
        {
            dynamic sections = pres.SectionProperties;
            int start = sections.FirstSlideIndex(sectionIndex);
            
            // Look at the first slide of the section
            if (pres.Slides.Count >= start)
            {
                Slide firstSlide = pres.Slides[start];
                if (firstSlide.Tags[AgendaSlideTag] == "True")
                {
                    return firstSlide;
                }
            }
            return null;
        }

        private void UpdateAgendaSlideContent(Slide slide, List<string> allSections, string activeSection)
        {
            // Set Title
            if (slide.Shapes.HasTitle == Office.MsoTriState.msoTrue)
            {
                slide.Shapes.Title.TextFrame.TextRange.Text = "Agenda";
            }

            // Find the main text body (usually the first placeholder that isn't the title)
            Shape body = null;
            foreach (Shape s in slide.Shapes)
            {
                if (s.Type == Office.MsoShapeType.msoPlaceholder && s != slide.Shapes.Title)
                {
                    body = s;
                    break;
                }
            }

            if (body == null) return;

            var textRange = body.TextFrame.TextRange;
            textRange.Text = ""; // Clear existing

            for (int i = 0; i < allSections.Count; i++)
            {
                string name = allSections[i];
                var line = textRange.InsertAfter(name + (i < allSections.Count - 1 ? "\r" : ""));
                
                if (name == activeSection)
                {
                    line.Font.Bold = Office.MsoTriState.msoTrue;
                    line.Font.Color.RGB = 0x4CAF50; // Green highlight
                    // You could also add a bullet point or indentation here
                }
                else
                {
                    line.Font.Bold = Office.MsoTriState.msoFalse;
                    line.Font.Color.RGB = 0x888888; // Dimmed for non-active
                }
            }
        }
    }
}
