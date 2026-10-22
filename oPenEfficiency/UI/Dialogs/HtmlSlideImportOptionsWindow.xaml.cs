using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using oPenEfficiency.Services;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class HtmlSlideImportOptionsWindow : Window
    {
        public ObservableCollection<HtmlSlideImportOptionItem> Slides { get; } = new ObservableCollection<HtmlSlideImportOptionItem>();

        public HtmlImportOptions SelectedOptions { get; private set; }

        public HtmlSlideImportOptionsWindow(HtmlImportPreview preview, HtmlImportOptions initialOptions = null)
        {
            InitializeComponent();
            DataContext = this;

            if (preview != null)
            {
                foreach (var slide in preview.Slides)
                {
                    Slides.Add(new HtmlSlideImportOptionItem
                    {
                        Number = slide.Number,
                        Title = slide.Title,
                        IsSelected = true
                    });
                }
            }

            if (initialOptions != null)
            {
                ChkPreserveStyling.IsChecked = initialOptions.PreserveSourceStyling;
                ChkPreserveBackground.IsChecked = initialOptions.PreserveSourceBackground;
                ChkPreserveFonts.IsChecked = initialOptions.PreserveSourceFonts;
                ChkUsePresentationTheme.IsChecked = initialOptions.UsePresentationTheme;
            }
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var slide in Slides)
            {
                slide.IsSelected = true;
            }
        }

        private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            foreach (var slide in Slides)
            {
                slide.IsSelected = false;
            }
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            SelectedOptions = new HtmlImportOptions
            {
                PreserveSourceStyling = ChkPreserveStyling.IsChecked == true,
                PreserveSourceBackground = ChkPreserveBackground.IsChecked == true,
                PreserveSourceFonts = ChkPreserveFonts.IsChecked == true,
                UsePresentationTheme = ChkUsePresentationTheme.IsChecked == true
            };

            foreach (var slide in Slides.Where(s => s.IsSelected))
            {
                SelectedOptions.SelectedSlideNumbers.Add(slide.Number);
            }

            if (SelectedOptions.SelectedSlideNumbers.Count == 0)
            {
                MessageBox.Show(
                    "Please select at least one slide to import.",
                    "Import HTML Slides",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }

    public class HtmlSlideImportOptionItem : INotifyPropertyChanged
    {
        private bool _isSelected;

        public int Number { get; set; }
        public string Title { get; set; }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value) return;
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public string DisplayText => $"{Number}. {(string.IsNullOrWhiteSpace(Title) ? "(Untitled slide)" : Title)}";

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
