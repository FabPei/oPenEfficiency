using System;
using System.Collections.ObjectModel;
using System.Windows.Media.Imaging;

namespace oPenEfficiency.Features
{
    public class PresentationFile
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string FolderName { get; set; }
        public string RootFolder { get; set; }
        public string SubFolder { get; set; }
        public BitmapImage Thumbnail { get; set; }
    }

    public class ShapeFile
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string FolderName { get; set; }
        public string RootFolder { get; set; }
        public string SubFolder { get; set; }
    }

    public class ImageFile
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string FolderName { get; set; }
        public string RootFolder { get; set; }
        public string SubFolder { get; set; }
    }

    public class ShapeItem
    {
        public string IdStr { get; set; }
        public int UniqueId { get; set; }
        public string Title { get; set; }
        public int SlideIndex { get; set; }
        public int OriginalShapeId { get; set; }
    }

    public class FolderTreeNode
    {
        public string Name { get; set; }
        public string FullPath { get; set; }
        public int FileCount { get; set; }
        public int TotalFileCount { get; set; }
        public bool IsFile { get; set; }
        public bool IsFolder => !IsFile;
        public bool IsCached { get; set; }
        public string FilePath { get; set; }
        public ObservableCollection<FolderTreeNode> Children { get; set; } = new ObservableCollection<FolderTreeNode>();

        public string DisplayName
        {
            get
            {
                if (IsFile) return Name;
                return TotalFileCount > 0 ? $"{Name}  ({TotalFileCount})" : Name;
            }
        }

        public string Icon
        {
            get
            {
                if (IsFolder) return "\uD83D\uDCC1"; // Folder 📁
                if (IsCached) return "\uD83D\uDCC4"; // Document 📄
                return "\u2601"; // Cloud ☁ (Pending cache)
            }
        }

        public string IconColor
        {
            get
            {
                if (IsFolder) return "#2563EB"; // Accent Blue for folders
                if (IsCached) return "#10B981"; // Success Green for cached files
                return "#F59E0B"; // Warning Orange for uncached files
            }
        }

        public string StatusTooltip => IsFile ? (IsCached ? "Ready (Cached)" : "Requires generation (Will take a moment)") : null;
    }
}
