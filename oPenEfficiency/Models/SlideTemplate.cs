using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;

namespace oPenEfficiency.Models
{
    public class SlideTemplate : INotifyPropertyChanged
    {
        private string _title;
        public string Title
        {
            get { return _title; }
            set { _title = value; OnPropertyChanged("Title"); }
        }

        private System.Windows.Media.ImageSource _thumbnailSource;
        public System.Windows.Media.ImageSource ThumbnailSource
        {
            get { return _thumbnailSource; }
            set { _thumbnailSource = value; OnPropertyChanged("ThumbnailSource"); }
        }

        private string _category;
        public string Category
        {
            get { return _category; }
            set { _category = value; OnPropertyChanged("Category"); }
        }

        private string _thumbnailPath;
        public string ThumbnailPath
        {
            get { return _thumbnailPath; }
            set { _thumbnailPath = value; OnPropertyChanged("ThumbnailPath"); }
        }

        private string _filePath;
        public string FilePath
        {
            get { return _filePath; }
            set { _filePath = value; OnPropertyChanged("FilePath"); }
        }
        
        private string _keywords;
        public string Keywords
        {
            get { return _keywords; }
            set { _keywords = value; OnPropertyChanged("Keywords"); }
        }

        private bool _isFavorite;
        public bool IsFavorite
        {
            get { return _isFavorite; }
            set { _isFavorite = value; OnPropertyChanged("IsFavorite"); }
        }

        /// <summary>1-based slide index within the source .pptx file. -1 means "all slides".</summary>
        public int SlideIndex { get; set; } = -1;

        /// <summary>Total slides in the source presentation (used for subtitle display).</summary>
        public int PresentationSlideCount { get; set; } = 1;

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
