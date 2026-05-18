using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Features.Utilities;

namespace oPenEfficiency.UI.Toolbars
{
    public partial class SlideNotesWindow : Window
    {
        private readonly PowerPointManager _manager;
        private Microsoft.Office.Interop.PowerPoint.Application _app;
        private ToggleButton _toggleBtn;
        private DispatcherTimer _debounceTimer;
        private Slide _currentSlide;
        private bool _isUpdating = false;

        private const string TagName = "oPE_SlideNotes";

        public SlideNotesWindow(PowerPointManager manager, ToggleButton toggleBtn)
        {
            InitializeComponent();
            _manager = manager;
            _toggleBtn = toggleBtn;
            _app = _manager.GetApplication();

            if (_app != null)
            {
                try { _app.WindowSelectionChange += App_WindowSelectionChange; } catch { }
            }

            _debounceTimer = new DispatcherTimer();
            _debounceTimer.Interval = TimeSpan.FromMilliseconds(500);
            _debounceTimer.Tick += DebounceTimer_Tick;

            this.Closed += SlideNotesWindow_Closed;

            LoadNotesForCurrentSlide();
        }

        private void App_WindowSelectionChange(Selection Sel)
        {
            // Dispatch to UI thread to update TextBox safely
            this.Dispatcher.BeginInvoke(new Action(() =>
            {
                LoadNotesForCurrentSlide();
            }));
        }

        private void LoadNotesForCurrentSlide()
        {
            if (_app == null || _app.ActiveWindow == null) return;
            
            try
            {
                var view = _app.ActiveWindow.View;
                if (view != null && view.Slide != null)
                {
                    var slide = view.Slide as Slide;
                    if (slide != null)
                    {
                        if (_currentSlide != null && _currentSlide.SlideID == slide.SlideID) return; // Same slide

                        // Before switching slides, save any pending changes
                        if (_debounceTimer.IsEnabled)
                        {
                            _debounceTimer.Stop();
                            SaveNotesToCurrentSlide();
                        }

                        _currentSlide = slide;
                        string notes = "";
                        try
                        {
                            foreach (string tag in slide.Tags)
                            {
                                if (tag == TagName)
                                {
                                    notes = slide.Tags[tag];
                                    break;
                                }
                            }
                        }
                        catch { }

                        _isUpdating = true;
                        NotesTextBox.Text = notes;
                        _isUpdating = false;
                    }
                }
            }
            catch { }
        }

        private void NotesTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdating) return;

            // Restart debounce timer
            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        private void DebounceTimer_Tick(object sender, EventArgs e)
        {
            _debounceTimer.Stop();
            SaveNotesToCurrentSlide();
        }

        private void SaveNotesToCurrentSlide()
        {
            if (_currentSlide == null) return;

            try
            {
                string notes = NotesTextBox.Text;
                if (string.IsNullOrEmpty(notes))
                {
                    try { _currentSlide.Tags.Delete(TagName); } catch { }
                }
                else
                {
                    _currentSlide.Tags.Add(TagName, notes);
                }
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "SlideNotesWindow.SaveNotesToCurrentSlide");
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        public void ClearNotesIfSlideMatches(int slideId)
        {
            if (_currentSlide != null && _currentSlide.SlideID == slideId)
            {
                _isUpdating = true;
                NotesTextBox.Text = "";
                _isUpdating = false;
                _debounceTimer.Stop(); // Don't re-save empty back just in case
            }
        }
        
        public void ForceRefresh()
        {
            _currentSlide = null; // force reload
            LoadNotesForCurrentSlide();
        }

        private void SlideNotesWindow_Closed(object sender, EventArgs e)
        {
            if (_app != null)
            {
                try { _app.WindowSelectionChange -= App_WindowSelectionChange; } catch { }
            }

            _debounceTimer.Stop();
            SaveNotesToCurrentSlide(); // Final save before closing

            if (_toggleBtn != null) _toggleBtn.IsChecked = false;
            SlideNotesFeature.SetWindowNull();
        }
    }
}
