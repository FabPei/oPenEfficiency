using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI.Dialogs;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnConsolidateMasters",
        Name = "Consolidate Slide Masters",
        Tooltip = "Consolidate Slide Masters",
        IconData = "M3,4H21V8H3V4 M3,10H21V14H3V10 M3,16H21V20H3V16 M17,2H17.01 M17,8H17.01 M17,14H17.01 M17,20H17.01 M21,2V22H3V2H21Z M12,2A2,2 0 0,1 14,4A2,2 0 0,1 12,6A2,2 0 0,1 10,4A2,2 0 0,1 12,2M12,8A2,2 0 0,1 14,10A2,2 0 0,1 12,12A2,2 0 0,1 10,10A2,2 0 0,1 12,8M12,14A2,2 0 0,1 14,16A2,2 0 0,1 12,18A2,2 0 0,1 10,16A2,2 0 0,1 12,14Z",
        Color = "#F43F5E",
        MinSelection = 0,
        RequiredType = PowerPoint.PpSelectionType.ppSelectionNone,
        Description = "A hard cleanup function that forces all slides to use the first slide master. Any other masters are deleted. Re-mapped slides get a warning sticker.",
        DetailedHelpText = "### Consolidate Slide Masters\nDeletes all redundant Slide Masters accumulated from copy-pasting slides, keeping only the primary master. Orphaned slides are automatically remapped and visually flagged for review.",
        Keywords = "cleanup, merge, remove, unused, templates, primary")
    ]
    public static class ConsolidateMastersFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var pres = manager.GetApplication().ActivePresentation;
                if (pres.Designs.Count <= 1)
                {
                    MessageBox.Show("Presentation already has only one Slide Master.", "Consolidate Masters", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return true;
                }

                var result = MessageBox.Show("This action will force all slides onto the primary Master and permanently delete all other Slide Masters.\n\nMigrated slides will receive a warning sticker for you to review.\n\nDo you want to proceed?", "Consolidate Slide Masters", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result != DialogResult.Yes) return false;

                // 1. Identify primary design (Design 1)
                PowerPoint.Design primaryDesign = pres.Designs[1];
                string primaryDesignName = primaryDesign.Name;

                int mappedCount = 0;

                // 2. Map all slides to the primary design
                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    try
                    {
                        PowerPoint.Design slideDesign = slide.Design;
                        if (slideDesign.Name != primaryDesignName)
                        {
                            PowerPoint.CustomLayout oldLayout = slide.CustomLayout;
                            PowerPoint.CustomLayout newLayout = FindBestLayoutMatch(primaryDesign, oldLayout);
                            
                            // Re-map slide
                            slide.CustomLayout = newLayout;
                            slide.Design = primaryDesign;
                            
                            // Mark for review
                            AddWarningSticker(slide, pres);
                            mappedCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log("Failed to map slide: " + ex.Message, "ConsolidateMastersFeature");
                    }
                }

                // 3. Delete the now unused Designs (Must iterate backwards)
                int deletedCount = 0;
                for (int i = pres.Designs.Count; i >= 2; i--)
                {
                    try
                    {
                        pres.Designs[i].Delete();
                        deletedCount++;
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log("Failed to delete design: " + ex.Message, "ConsolidateMastersFeature");
                    }
                }

                MessageBox.Show($"Successfully consolidated masters.\n\nDeleted Masters: {deletedCount}\nRe-mapped Slides: {mappedCount}\n\nPlease review the marked slides for any visual shifts.", "Consolidate Masters", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex.Message, "ConsolidateMastersFeature");
                return false;
            }
        }

        private static PowerPoint.CustomLayout FindBestLayoutMatch(PowerPoint.Design targetDesign, PowerPoint.CustomLayout sourceLayout)
        {
            try
            {
                string sourceName = sourceLayout.Name;

                // Pass 1: Try to match by exact Name
                foreach (PowerPoint.CustomLayout targetLayout in targetDesign.SlideMaster.CustomLayouts)
                {
                    if (targetLayout.Name == sourceName)
                    {
                        return targetLayout;
                    }
                }

                // Removed Pass 2 LayoutType matching because CustomLayout doesn't expose Type natively in this interop version.

                // Fallback: Just return the first layout (usually Title or Blank)
                if (targetDesign.SlideMaster.CustomLayouts.Count > 0)
                {
                    return targetDesign.SlideMaster.CustomLayouts[1];
                }
            }
            catch (Exception)
            {
                // Ignore matching errors and let it fallback
            }

            return null; // Should not happen ideally
        }

        private static void AddWarningSticker(PowerPoint.Slide slide, PowerPoint.Presentation pres)
        {
            try
            {
                // Create a prominent red warning box at top right
                float slideWidth = pres.PageSetup.SlideWidth;
                var shape = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, slideWidth - 160, 10, 150, 25);
                shape.Fill.ForeColor.RGB = 0x0000FF; // Red
                shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                shape.TextFrame.TextRange.Text = "Master Re-mapped: Check";
                shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF; // White
                shape.TextFrame.TextRange.Font.Size = 10;
                shape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                shape.TextFrame.MarginLeft = 2;
                shape.TextFrame.MarginRight = 2;
                shape.TextFrame.MarginTop = 2;
                shape.TextFrame.MarginBottom = 2;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log("Failed to add warning sticker: " + ex.Message, "ConsolidateMastersFeature");
            }
        }
    }
}
