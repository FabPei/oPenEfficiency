using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnConsolidateMasters",
        Name = "Consolidate Slide Masters",
        SearchAliases = new[] { "Consolidate Masters" },
        Tooltip = "Merge duplicate slide masters and layouts",
        IconData = "M3,4H21V8H3V4 M3,10H21V14H3V10 M3,16H21V20H3V16 M17,2H17.01 M17,8H17.01 M17,14H17.01 M17,20H17.01 M21,2V22H3V2H21Z M12,2A2,2 0 0,1 14,4A2,2 0 0,1 12,6A2,2 0 0,1 10,4A2,2 0 0,1 12,2M12,8A2,2 0 0,1 14,10A2,2 0 0,1 12,12A2,2 0 0,1 10,10A2,2 0 0,1 12,8M12,14A2,2 0 0,1 14,16A2,2 0 0,1 12,18A2,2 0 0,1 10,16A2,2 0 0,1 12,14Z",
        Color = "#F43F5E",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None,
        Description = "Backs up the presentation, remaps slides to the primary master, flags migrated slides for review, and removes redundant masters.",
        DetailedHelpText = "### Consolidate Slide Masters\nCreates a timestamped backup copy before changing the open presentation. Slides that use another master are mapped to an equally named layout on the primary master where possible, with Blank or the first layout as a fallback. Migrated slides receive a tagged review marker. If slide mapping fails, completed mappings are rolled back and no masters are deleted. If master deletion fails, the command reports the incomplete result and the backup path.")]
    public static class ConsolidateMastersFeature
    {
        private const string ReviewMarkerTag = "oPE_MasterConsolidationWarning";

        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;

            try
            {
                var app = manager.GetApplication();
                var pres = app.ActivePresentation;
                if (pres == null) return false;

                if (pres.Designs.Count <= 1)
                {
                    MessageBox.Show("Presentation already has only one Slide Master.", "Consolidate Masters", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return true;
                }

                var result = MessageBox.Show(
                    $"Remap all slides to the primary master and delete {pres.Designs.Count - 1} other master(s)?\n\n" +
                    "A timestamped backup copy will be created first. Migrated slides will receive a review marker.",
                    "Consolidate Slide Masters",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (result != DialogResult.Yes) return false;

                string backupPath;
                try
                {
                    backupPath = CreateBackupCopy(pres);
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "ConsolidateMastersFeature.CreateBackupCopy");
                    MessageBox.Show(
                        $"No changes were made because the backup copy could not be created.\n\n{ex.Message}",
                        "Consolidate Masters",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return false;
                }

                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ConsolidateMastersFeature.StartNewUndoEntry"); }

                PowerPoint.Design primaryDesign = pres.Designs[1];
                var mappings = new List<SlideMapping>();

                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    PowerPoint.Design sourceDesign = slide.Design;
                    if (IsSameComObject(sourceDesign, primaryDesign)) continue;

                    PowerPoint.CustomLayout sourceLayout = slide.CustomLayout;
                    PowerPoint.CustomLayout targetLayout = FindBestLayoutMatch(primaryDesign, sourceLayout);
                    if (targetLayout == null || !TryMapSlide(slide, primaryDesign, targetLayout, sourceDesign, sourceLayout, mappings))
                    {
                        RollBackMappings(mappings);
                        MessageBox.Show(
                            $"Master consolidation stopped while mapping slide {slide.SlideIndex}. Completed mappings were rolled back and no masters were deleted.\n\nBackup: {backupPath}",
                            "Consolidate Masters",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return false;
                    }
                }

                foreach (var mapping in mappings) AddWarningSticker(mapping.Slide, pres);

                int deletedCount = 0;
                var deletionFailures = new List<string>();
                for (int i = pres.Designs.Count; i >= 2; i--)
                {
                    try
                    {
                        pres.Designs[i].Delete();
                        deletedCount++;
                    }
                    catch (Exception ex)
                    {
                        deletionFailures.Add($"Master {i}: {ex.Message}");
                        ExceptionLogger.Log(ex, "ConsolidateMastersFeature.DeleteDesign");
                    }
                }

                if (deletionFailures.Count > 0)
                {
                    MessageBox.Show(
                        $"Slide mapping completed, but {deletionFailures.Count} master(s) could not be deleted.\n\n" +
                        $"Deleted masters: {deletedCount}\nRe-mapped slides: {mappings.Count}\nBackup: {backupPath}",
                        "Consolidate Masters — Incomplete",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
                }

                MessageBox.Show(
                    $"Consolidated slide masters.\n\nDeleted masters: {deletedCount}\nRe-mapped slides: {mappings.Count}\nBackup: {backupPath}\n\nReview the marked slides for visual shifts.",
                    "Consolidate Masters",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ConsolidateMastersFeature.Execute");
                MessageBox.Show($"Master consolidation failed.\n\n{ex.Message}", "Consolidate Masters", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private static bool TryMapSlide(
            PowerPoint.Slide slide,
            PowerPoint.Design targetDesign,
            PowerPoint.CustomLayout targetLayout,
            PowerPoint.Design sourceDesign,
            PowerPoint.CustomLayout sourceLayout,
            List<SlideMapping> mappings)
        {
            try
            {
                slide.Design = targetDesign;
                slide.CustomLayout = targetLayout;
                mappings.Add(new SlideMapping(slide, sourceDesign, sourceLayout));
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ConsolidateMastersFeature.TryMapSlide");
                try
                {
                    slide.Design = sourceDesign;
                    slide.CustomLayout = sourceLayout;
                }
                catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "ConsolidateMastersFeature.TryMapSlide.RollbackCurrent"); }
                return false;
            }
        }

        private static void RollBackMappings(List<SlideMapping> mappings)
        {
            for (int i = mappings.Count - 1; i >= 0; i--)
            {
                try
                {
                    mappings[i].Slide.Design = mappings[i].SourceDesign;
                    mappings[i].Slide.CustomLayout = mappings[i].SourceLayout;
                }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ConsolidateMastersFeature.RollBackMappings"); }
            }
        }

        private static PowerPoint.CustomLayout FindBestLayoutMatch(PowerPoint.Design targetDesign, PowerPoint.CustomLayout sourceLayout)
        {
            try
            {
                string sourceName = sourceLayout?.Name ?? string.Empty;
                foreach (PowerPoint.CustomLayout targetLayout in targetDesign.SlideMaster.CustomLayouts)
                {
                    if (string.Equals(targetLayout.Name, sourceName, StringComparison.OrdinalIgnoreCase)) return targetLayout;
                }

                foreach (PowerPoint.CustomLayout targetLayout in targetDesign.SlideMaster.CustomLayouts)
                {
                    if (targetLayout.Name.IndexOf("blank", StringComparison.OrdinalIgnoreCase) >= 0) return targetLayout;
                }

                if (targetDesign.SlideMaster.CustomLayouts.Count > 0) return targetDesign.SlideMaster.CustomLayouts[1];
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ConsolidateMastersFeature.FindBestLayoutMatch");
            }
            return null;
        }

        private static string CreateBackupCopy(PowerPoint.Presentation presentation)
        {
            string directory = string.IsNullOrWhiteSpace(presentation.Path)
                ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                : presentation.Path;
            string extension = Path.GetExtension(presentation.Name);
            if (string.IsNullOrWhiteSpace(extension)) extension = ".pptx";
            string baseName = Path.GetFileNameWithoutExtension(presentation.Name);
            if (string.IsNullOrWhiteSpace(baseName)) baseName = "Presentation";

            string backupPath = Path.Combine(directory, $"{baseName}_before_master_consolidation_{DateTime.Now:yyyyMMdd_HHmmss}{extension}");
            if (File.Exists(backupPath))
                backupPath = Path.Combine(directory, $"{baseName}_before_master_consolidation_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}{extension}");

            presentation.SaveCopyAs(backupPath);
            return backupPath;
        }

        private static bool IsSameComObject(object first, object second)
        {
            IntPtr firstPointer = IntPtr.Zero;
            IntPtr secondPointer = IntPtr.Zero;
            try
            {
                firstPointer = Marshal.GetIUnknownForObject(first);
                secondPointer = Marshal.GetIUnknownForObject(second);
                return firstPointer == secondPointer;
            }
            finally
            {
                if (firstPointer != IntPtr.Zero) Marshal.Release(firstPointer);
                if (secondPointer != IntPtr.Zero) Marshal.Release(secondPointer);
            }
        }

        private static void AddWarningSticker(PowerPoint.Slide slide, PowerPoint.Presentation pres)
        {
            try
            {
                float slideWidth = pres.PageSetup.SlideWidth;
                var shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, slideWidth - 160, 10, 150, 25);
                shape.Name = "oPE_MasterConsolidationWarning";
                shape.Tags.Add(ReviewMarkerTag, "True");
                shape.Fill.ForeColor.RGB = 0x0000FF;
                shape.Line.Visible = MsoTriState.msoFalse;
                shape.TextFrame.TextRange.Text = "Master Re-mapped: Check";
                shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF;
                shape.TextFrame.TextRange.Font.Size = 10;
                shape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                shape.TextFrame.MarginLeft = 2;
                shape.TextFrame.MarginRight = 2;
                shape.TextFrame.MarginTop = 2;
                shape.TextFrame.MarginBottom = 2;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ConsolidateMastersFeature.AddWarningSticker");
            }
        }

        private sealed class SlideMapping
        {
            public SlideMapping(PowerPoint.Slide slide, PowerPoint.Design sourceDesign, PowerPoint.CustomLayout sourceLayout)
            {
                Slide = slide;
                SourceDesign = sourceDesign;
                SourceLayout = sourceLayout;
            }

            public PowerPoint.Slide Slide { get; }
            public PowerPoint.Design SourceDesign { get; }
            public PowerPoint.CustomLayout SourceLayout { get; }
        }
    }
}
