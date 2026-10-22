using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointApplication = Microsoft.Office.Interop.PowerPoint.Application;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnStickyNoteManager",
        Name = "Sticky Note Manager",
        Tooltip = "Manage sticky notes",
        IconData = "M19,3H5C3.89,3 3,3.89 3,5V19C3,20.11 3.89,21 5,21H19C20.11,21 21,20.11 21,19V5C21,3.89 20.11,3 19,3M19,19H5V5H19V19M7,7H17V9H7V7M7,11H17V13H7V11M7,15H17V17H7V15Z",
        Color = "#F43F5E",
        Description = "Moves, restores, deletes, or creates sticky notes from PowerPoint comments.",
        DetailedHelpText = "### Sticky Note Manager\nA normal click moves notes off the current slide without deleting them. The context menu provides restore and confirmed delete actions for the current or selected slides, copies PowerPoint comments to tagged sticky notes while retaining the originals, and marks selected shapes as sticky notes.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class StickyNoteManagerFeature
    {
        private const string OriginalLeftTag = "oPE_StickyNoteOrigLeft";
        private const string OriginalTopTag = "oPE_StickyNoteOrigTop";
        private const string LegacyOriginalLeftTag = "OrigLeft";
        private const string LegacyOriginalTopTag = "OrigTop";
        private const string SourceCommentTag = "oPE_SourceComment";

        public enum ActionType { MoveOff, MoveOn, Delete, ConvertComments, PrefixName }
        public enum Scope { ThisSlide, SelectedSlides }

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, ActionType.MoveOff, Scope.ThisSlide);
        }

        public static bool Execute(PowerPointManager manager, ActionType action, Scope scope)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app?.ActiveWindow == null || app.ActivePresentation == null) return false;

                if (action == ActionType.PrefixName) return MarkSelectedShapes(app);

                var slides = GetSlides(manager, app, scope);
                if (slides.Count == 0) return false;

                if (action == ActionType.Delete)
                {
                    int count = slides.Sum(CountStickyNotes);
                    if (count == 0) return true;
                    string target = scope == Scope.ThisSlide ? "the current slide" : $"{slides.Count} selected slide(s)";
                    if (MessageBox.Show($"Delete {count} sticky note(s) from {target}? This cannot reliably be undone.",
                        "Delete Sticky Notes", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                        return false;
                }

                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "StickyNoteManagerFeature.StartUndo"); }

                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
                foreach (var slide in slides)
                {
                    if (action == ActionType.Delete) DeleteNotes(slide);
                    else if (action == ActionType.MoveOff) MoveNotesOff(slide, slideWidth, slideHeight);
                    else if (action == ActionType.MoveOn) MoveNotesOn(slide, slideWidth, slideHeight);
                    else if (action == ActionType.ConvertComments) CopyCommentsToNotes(app, slide, slideWidth);
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StickyNoteManagerFeature.Execute");
                return false;
            }
        }

        private static List<Slide> GetSlides(PowerPointManager manager, PowerPointApplication app, Scope scope)
        {
            if (scope == Scope.SelectedSlides)
                return manager.GetSelectedSlidesForOperation()?.Where(s => s != null).ToList() ?? new List<Slide>();
            var slide = app.ActiveWindow.View.Slide as Slide;
            return slide == null ? new List<Slide>() : new List<Slide> { slide };
        }

        private static bool MarkSelectedShapes(PowerPointApplication app)
        {
            var selection = app.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count == 0) return false;
            try { app.StartNewUndoEntry(); } catch { /* Note updates remain valid when undo grouping is unavailable. */ }
            foreach (Shape shape in selection.ShapeRange)
            {
                StickyNoteFeature.MarkStickyNote(shape, "Custom");
                if (!shape.Name.StartsWith("StickyNote_", StringComparison.OrdinalIgnoreCase))
                    shape.Name = "StickyNote_" + shape.Name;
            }
            return true;
        }

        private static int CountStickyNotes(Slide slide)
        {
            int count = 0;
            foreach (Shape shape in slide.Shapes) if (StickyNoteFeature.IsStickyNote(shape)) count++;
            return count;
        }

        private static void DeleteNotes(Slide slide)
        {
            for (int i = slide.Shapes.Count; i >= 1; i--)
                if (StickyNoteFeature.IsStickyNote(slide.Shapes[i])) slide.Shapes[i].Delete();
        }

        private static void MoveNotesOff(Slide slide, float slideWidth, float slideHeight)
        {
            foreach (Shape shape in slide.Shapes)
            {
                if (!StickyNoteFeature.IsStickyNote(shape) || !IsOnSlide(shape, slideWidth, slideHeight)) continue;
                shape.Tags.Add(OriginalLeftTag, shape.Left.ToString(CultureInfo.InvariantCulture));
                shape.Tags.Add(OriginalTopTag, shape.Top.ToString(CultureInfo.InvariantCulture));

                float left = shape.Left;
                float right = slideWidth - (shape.Left + shape.Width);
                float top = shape.Top;
                float bottom = slideHeight - (shape.Top + shape.Height);
                float nearest = Math.Min(Math.Min(left, right), Math.Min(top, bottom));
                if (nearest == left) shape.Left = -shape.Width - 50;
                else if (nearest == right) shape.Left = slideWidth + 50;
                else if (nearest == top) shape.Top = -shape.Height - 50;
                else shape.Top = slideHeight + 50;
            }
        }

        private static void MoveNotesOn(Slide slide, float slideWidth, float slideHeight)
        {
            foreach (Shape shape in slide.Shapes)
            {
                if (!StickyNoteFeature.IsStickyNote(shape) || IsOnSlide(shape, slideWidth, slideHeight)) continue;
                string left = GetTag(shape, OriginalLeftTag) ?? GetTag(shape, LegacyOriginalLeftTag);
                string top = GetTag(shape, OriginalTopTag) ?? GetTag(shape, LegacyOriginalTopTag);
                if (float.TryParse(left, NumberStyles.Any, CultureInfo.InvariantCulture, out float oldLeft) &&
                    float.TryParse(top, NumberStyles.Any, CultureInfo.InvariantCulture, out float oldTop))
                {
                    shape.Left = oldLeft;
                    shape.Top = oldTop;
                }
                else
                {
                    shape.Left = slideWidth - shape.Width - 50;
                    shape.Top = 50;
                }
                TryDeleteTag(shape, OriginalLeftTag);
                TryDeleteTag(shape, OriginalTopTag);
            }
        }

        private static bool IsOnSlide(Shape shape, float width, float height)
        {
            return shape.Left < width && shape.Left + shape.Width > 0 && shape.Top < height && shape.Top + shape.Height > 0;
        }

        private static void CopyCommentsToNotes(PowerPointApplication app, Slide slide, float slideWidth)
        {
            if (slide.Comments == null || slide.Comments.Count == 0) return;
            var existingKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Shape shape in slide.Shapes)
            {
                string key = GetTag(shape, SourceCommentTag);
                if (!string.IsNullOrEmpty(key)) existingKeys.Add(key);
            }

            float top = 50;
            for (int i = 1; i <= slide.Comments.Count; i++)
            {
                var comment = slide.Comments[i];
                string key = HashComment(comment.Author, comment.Text);
                if (existingKeys.Contains(key)) continue;
                string text = $"{comment.Author}:\n{comment.Text}";
                const float width = 184.25f;
                const float height = 62.36f;
                var shape = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeFoldedCorner,
                    slideWidth - width - 50, top, width, height);
                shape.Name = "StickyNote_Comment_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                StickyNoteFeature.MarkStickyNote(shape, "PowerPointComment");
                shape.Tags.Add(SourceCommentTag, key);
                shape.Fill.ForeColor.RGB = 0x88D2F8;
                shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                shape.TextFrame.TextRange.Text = text;
                shape.TextFrame.TextRange.Font.Size = 10;
                top += height + 10;
            }
        }

        private static string HashComment(string author, string text)
        {
            using (var sha = SHA256.Create())
                return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes((author ?? "") + "\n" + (text ?? ""))));
        }

        private static string GetTag(Shape shape, string name)
        {
            try
            {
                for (int i = 1; i <= shape.Tags.Count; i++)
                    if (string.Equals(shape.Tags.Name(i), name, StringComparison.OrdinalIgnoreCase)) return shape.Tags.Value(i);
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "StickyNoteManagerFeature.GetTag"); }
            return null;
        }

        private static void TryDeleteTag(Shape shape, string name)
        {
            try { shape.Tags.Delete(name); }
            catch { /* Deleting a missing sticky-note tag is idempotent. */ }
        }
    }
}
