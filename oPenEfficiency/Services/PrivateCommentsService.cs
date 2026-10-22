using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Stores oPenEfficiency-only comments in the presentation package. These comments are
    /// deliberately not PowerPoint comments and therefore are not shown in PowerPoint's review UI.
    /// They are metadata, not a security boundary: anyone inspecting the .pptx package can read them.
    /// </summary>
    public static class PrivateCommentsService
    {
        private const string StoreId = "private-comments-v2";
        private const string LegacyTagName = "oPE_SlideNotes";

        public static PrivateComment GetForSlide(Presentation presentation, Slide slide)
        {
            var comments = GetCommentsForSlide(presentation, slide);
            return comments.FirstOrDefault(comment => !comment.IsDone) ?? comments.FirstOrDefault();
        }

        public static List<PrivateComment> GetCommentsForSlide(Presentation presentation, Slide slide)
        {
            if (presentation == null || slide == null) return new List<PrivateComment>();

            var store = Load(presentation);
            MigrateLegacyComment(presentation, slide, store);
            return store.Comments.Where(comment => comment.SlideId == slide.SlideID)
                .OrderBy(comment => comment.IsDone)
                .ThenByDescending(comment => comment.UpdatedAtUtc)
                .ToList();
        }

        public static List<PrivateComment> GetAllComments(Presentation presentation)
        {
            if (presentation == null) return new List<PrivateComment>();
            var existingSlideIds = new HashSet<int>(presentation.Slides.Cast<Slide>().Select(slide => slide.SlideID));
            return Load(presentation).Comments.Where(comment => existingSlideIds.Contains(comment.SlideId))
                .OrderBy(comment => comment.IsDone)
                .ThenBy(comment => comment.SlideId)
                .ThenByDescending(comment => comment.UpdatedAtUtc)
                .ToList();
        }

        public static PrivateComment CreateComment(Slide slide, Shape anchor = null)
        {
            if (slide == null) return null;
            var now = DateTime.UtcNow;
            return new PrivateComment
            {
                Id = Guid.NewGuid().ToString("N"),
                SlideId = slide.SlideID,
                Author = Environment.UserName,
                CreatedAtUtc = now,
                UpdatedAtUtc = now,
                Status = PrivateCommentStatus.Open,
                AnchorShapeId = anchor?.Id ?? 0,
                AnchorShapeName = anchor?.Name
            };
        }

        public static int GetOpenCount(Presentation presentation, Slide slide)
        {
            return GetCommentsForSlide(presentation, slide).Count(comment => !comment.IsDone);
        }

        private static void MigrateLegacyComment(Presentation presentation, Slide slide, PrivateCommentStore store)
        {
            if (store.Comments.Any(existingComment => existingComment.SlideId == slide.SlideID)) return;

            // One-time, lossless migration of the former tag-based scratchpad.
            string legacyText = GetTagValue(slide, LegacyTagName);
            if (string.IsNullOrWhiteSpace(legacyText)) return;

            var comment = CreateComment(slide);
            comment.Text = legacyText;
            store.Comments.Add(comment);
            Save(presentation, store);
            // Remove the legacy source only after the new store was written successfully.
            TryDeleteTag(slide, LegacyTagName);
        }

        public static void SaveForSlide(Presentation presentation, Slide slide, PrivateComment comment)
        {
            if (presentation == null || slide == null || comment == null) return;

            var store = Load(presentation);
            var existing = store.Comments.FirstOrDefault(c => c.Id == comment.Id);
            comment.SlideId = slide.SlideID;
            comment.UpdatedAtUtc = DateTime.UtcNow;
            comment.IsResolved = string.Equals(comment.Status, PrivateCommentStatus.Done, StringComparison.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(comment.Text))
            {
                if (existing != null) store.Comments.Remove(existing);
            }
            else if (existing == null)
            {
                store.Comments.Add(comment);
            }
            else
            {
                existing.Text = comment.Text;
                existing.Status = comment.Status;
                existing.IsResolved = comment.IsResolved;
                existing.AnchorShapeId = comment.AnchorShapeId;
                existing.AnchorShapeName = comment.AnchorShapeName;
                existing.UpdatedAtUtc = comment.UpdatedAtUtc;
            }

            Save(presentation, store);
        }

        public static void DeleteForSlides(Presentation presentation, IEnumerable<Slide> slides)
        {
            if (presentation == null || slides == null) return;
            var slideIds = new HashSet<int>(slides.Where(s => s != null).Select(s => s.SlideID));
            var store = Load(presentation);
            store.Comments.RemoveAll(c => slideIds.Contains(c.SlideId));
            foreach (var slide in slides.Where(s => s != null)) TryDeleteTag(slide, LegacyTagName);
            Save(presentation, store);
        }

        public static void DeleteAll(Presentation presentation)
        {
            if (presentation == null) return;
            DataStoreManager.DeleteData(presentation, StoreId);
            foreach (Slide slide in presentation.Slides) TryDeleteTag(slide, LegacyTagName);
        }

        private static PrivateCommentStore Load(Presentation presentation)
        {
            return JsonService.Deserialize<PrivateCommentStore>(DataStoreManager.LoadData(presentation, StoreId))
                ?? new PrivateCommentStore { Comments = new List<PrivateComment>() };
        }

        private static void Save(Presentation presentation, PrivateCommentStore store)
        {
            DataStoreManager.SaveData(presentation, StoreId, JsonService.Serialize(store));
        }

        public static void DeleteComment(Presentation presentation, string commentId)
        {
            if (presentation == null || string.IsNullOrEmpty(commentId)) return;
            var store = Load(presentation);
            store.Comments.RemoveAll(comment => comment.Id == commentId);
            Save(presentation, store);
        }

        private static string GetTagValue(Slide slide, string name)
        {
            try { return slide.Tags[name]; } catch { return null; }
        }

        private static void TryDeleteTag(Slide slide, string name)
        {
            try { slide.Tags.Delete(name); } catch { /* Deleting a missing private-comment tag is idempotent. */ }
        }
    }

    [DataContract]
    public class PrivateCommentStore
    {
        [DataMember] public List<PrivateComment> Comments { get; set; } = new List<PrivateComment>();
    }

    [DataContract]
    public class PrivateComment
    {
        [DataMember] public string Id { get; set; }
        [DataMember] public int SlideId { get; set; }
        [DataMember] public string Text { get; set; }
        [DataMember] public string Author { get; set; }
        [DataMember] public DateTime CreatedAtUtc { get; set; }
        [DataMember] public DateTime UpdatedAtUtc { get; set; }
        [DataMember] public string Status { get; set; } = PrivateCommentStatus.Open;
        [DataMember] public int AnchorShapeId { get; set; }
        [DataMember] public string AnchorShapeName { get; set; }

        // Compatibility with existing presentations that stored a boolean resolution flag.
        [DataMember] public bool IsResolved { get; set; }
        public bool IsDone
        {
            get
            {
                return IsResolved || string.Equals(Status, PrivateCommentStatus.Done, StringComparison.OrdinalIgnoreCase);
            }
        }

        public string Summary => $"[{(IsDone ? "Done" : "Open")}] {Text}";
    }

    public static class PrivateCommentStatus
    {
        public const string Open = "Open";
        public const string Done = "Done";
    }
}
