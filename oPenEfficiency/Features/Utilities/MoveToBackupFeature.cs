using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnMoveToBackup",
        Name = "Move Slides to Backup",
        SearchAliases = new[] { "Move to Backup" },
        Tooltip = "Move selected slides to the end behind the backup slide",
        IconData = "M14,3V5H17.59L7.76,14.83L9.17,16.24L19,6.41V10H21V3M19,19H5V5H12V3H5C3.89,3 3,3.9 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V12H19V19Z",
        Color = "#F59E0B",
        Description = "Moves selected slides, in their existing order, to the end of a presentation that contains a designated Backup slide.",
        DetailedHelpText = "### Move Slides to Backup\nMoves the selected slides to the end of the presentation, behind the Backup slide created by Agenda Creator. Right-click can move, copy, or locate the designated Backup slide. This feature never creates the Backup slide. Move and copy stop without changes when no designated Backup slide exists, exclude the Backup slide itself, require confirmation, and start a PowerPoint Undo entry where supported.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Slides)]
    public static class MoveToBackupFeature
    {
        public enum BackupAction { Move, Copy, Locate }

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, BackupAction.Move);
        }

        public static bool Execute(PowerPointManager manager, BackupAction action)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;
                var selection = app.ActiveWindow.Selection;

                var pres = app.ActivePresentation;
                Slide backupSlide = null;
                for (int i = pres.Slides.Count; i >= 1; i--)
                {
                    try
                    {
                        if (oPenEfficiency.Core.Services.BackupSlidePolicy.IsDesignatedBackupSlide(
                                pres.Slides[i].Tags[oPenEfficiency.Core.Services.BackupSlidePolicy.OpeBackupSlideTag],
                                pres.Slides[i].Tags[oPenEfficiency.Core.Services.BackupSlidePolicy.Ee4pBackupSlideTag]))
                        {
                            backupSlide = pres.Slides[i];
                            break;
                        }
                    }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "MoveToBackupFeature.FindBackupSlide"); }
                }

                if (backupSlide == null)
                {
                    MessageBox.Show(
                        "No designated Backup slide was found. Create one through Agenda Creator before moving slides to backup.",
                        "Move Slides to Backup",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return false;
                }

                if (action == BackupAction.Locate)
                {
                    app.ActiveWindow.View.GotoSlide(backupSlide.SlideIndex);
                    return true;
                }

                if (selection == null || selection.Type != PpSelectionType.ppSelectionSlides) return false;

                int backupSlideId = backupSlide.SlideID;
                var slidesToMove = new List<Slide>();
                foreach (Slide slide in selection.SlideRange)
                {
                    if (slide.SlideID != backupSlideId) slidesToMove.Add(slide);
                }
                slidesToMove = slidesToMove.OrderBy(slide => slide.SlideIndex).ToList();
                if (slidesToMove.Count == 0) return false;

                var result = MessageBox.Show(
                    oPenEfficiency.Core.Services.BackupSlidePolicy.BuildConfirmationText(
                        action == BackupAction.Copy, slidesToMove.Count, backupSlide.SlideIndex),
                    "Move Slides to Backup",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (result != DialogResult.Yes) return false;

                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "MoveToBackupFeature.StartNewUndoEntry"); }

                foreach (Slide slide in slidesToMove)
                {
                    if (action == BackupAction.Copy)
                    {
                        var duplicate = slide.Duplicate();
                        duplicate[1].MoveTo(pres.Slides.Count);
                    }
                    else
                    {
                        slide.MoveTo(pres.Slides.Count);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MoveToBackupFeature");
                return false;
            }
        }
    }
}
