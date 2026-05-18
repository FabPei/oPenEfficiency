using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using oPenEfficiency.Services;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public partial class StyleCheckWindow : Window
    {
        private PowerPointManager _powerPointManager;

        public StyleCheckWindow(PowerPointManager powerPointManager)
        {
            InitializeComponent();
            _powerPointManager = powerPointManager;
            RefreshData();
        }

        public void RefreshData()
        {
            try
            {
                var app = _powerPointManager.GetApplication();
                var slides = GetSelectedSlidesList(app);
                
                var styleErrors = StyleScanner.ScanSlides(app, slides);
                var consistencyErrors = ConsistencyScanner.ScanSlides(app, slides);
                
                DataContext = new UnifiedCheckViewModel(styleErrors, consistencyErrors, slides);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error analyzing slides: " + ex.Message);
            }
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            RefreshData();
        }

        private List<Slide> GetSelectedSlidesList(Microsoft.Office.Interop.PowerPoint.Application app)
        {
            var slides = new List<Slide>();
            try
            {
                if (app.ActiveWindow == null) return slides;
                var selection = app.ActiveWindow.Selection;

                if (selection.Type == PpSelectionType.ppSelectionSlides)
                {
                    foreach (Slide s in selection.SlideRange) slides.Add(s);
                }
                else if (selection.Type == PpSelectionType.ppSelectionShapes || selection.Type == PpSelectionType.ppSelectionText)
                {
                    try
                    {
                        var slideRange = selection.SlideRange; 
                        if (slideRange != null)
                        {
                            foreach (Slide s in slideRange) slides.Add(s);
                        }
                    } catch { }
                }

                if (slides.Count == 0)
                {
                    try { slides.Add((Slide)app.ActiveWindow.View.Slide); } catch { }
                }
            }
            catch { }
            return slides.Distinct().ToList();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Ignore_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is UnifiedErrorItem item)
            {
                try
                {
                    var app = _powerPointManager.GetApplication();
                    var pres = app.ActivePresentation;
                    Slide slide = null;
                    
                    try { slide = pres.Slides[item.SlideNumber]; } catch { }
                    if (slide == null) return;

                    if (item.ShapeName == "Slide")
                    {
                        slide.Tags.Add(item.ErrorCategory == "Consistency" ? "oPE_IgnoreConsistencyCheck" : "oPE_IgnoreStyleCheck", "True");
                    }
                    else
                    {
                        Shape shape = null;
                        try { shape = slide.Shapes[item.ShapeName]; } catch { }
                        if (shape != null)
                        {
                            shape.Tags.Add(item.ErrorCategory == "Consistency" ? "oPE_IgnoreConsistencyCheck" : "oPE_IgnoreStyleCheck", "True");
                        }
                    }
                    
                    // Optimistic UI update or full refresh
                    RefreshData();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not ignore item: " + ex.Message);
                }
            }
        }

        
        private void FixAll_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is UnifiedCheckViewModel vm)
            {
                var fixableItems = vm.BySlide.SelectMany(s => s.Errors)
                    .Where(err => err.CanAutoCorrect)
                    .OrderByDescending(err => err.TextStart)
                    .ToList();
                
                if (fixableItems.Count == 0) return;

                foreach (var item in fixableItems)
                {
                    ApplyAutoFix(item, false);
                }
                RefreshData();
            }
        }

        private void AutoFix_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is UnifiedErrorItem item)
            {
                ApplyAutoFix(item, true);
            }
        }

        private void ApplyAutoFix(UnifiedErrorItem item, bool refresh = true)
        {
            try
            {
                var app = _powerPointManager.GetApplication();
                var pres = app.ActivePresentation;
                Slide slide = null;
                try { slide = pres.Slides[item.SlideNumber]; } catch { }
                if (slide == null) return;

                Shape shape = null;
                try { shape = slide.Shapes[item.ShapeName]; } catch { }
                if (shape == null) return;

                if (item.ErrorCategory == "Style" && item.IsColor && !string.IsNullOrEmpty(item.SuggestedCorrection))
                {
                    var rgb = oPenEfficiency.Utils.ColorMath.HexToRgb(item.SuggestedCorrection);
                    int bgr = rgb.R | (rgb.G << 8) | (rgb.B << 16);

                    if (item.ErrorType == "Fill Color")
                    {
                        shape.Fill.ForeColor.RGB = bgr;
                        shape.Fill.ForeColor.ObjectThemeColor = Office.MsoThemeColorIndex.msoNotThemeColor;
                    }
                    else if (item.ErrorType == "Line Color")
                    {
                        shape.Line.ForeColor.RGB = bgr;
                        shape.Line.ForeColor.ObjectThemeColor = Office.MsoThemeColorIndex.msoNotThemeColor;
                    }
                    else if (item.ErrorType == "Text Color" && shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        if (item.TextLength > 0 && item.TextStart > 0)
                        {
                            shape.TextFrame.TextRange.Characters(item.TextStart, item.TextLength).Font.Color.RGB = bgr;
                        }
                        else
                        {
                            shape.TextFrame.TextRange.Font.Color.RGB = bgr;
                        }
                    }
                }
                else if (item.ErrorCategory == "Consistency" && item.ErrorType == "Wording" && !string.IsNullOrEmpty(item.SuggestedCorrection))
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        if (item.TextLength > 0 && item.TextStart > 0)
                        {
                            var tr = shape.TextFrame.TextRange.Characters(item.TextStart, item.TextLength);
                            tr.Text = item.SuggestedCorrection;
                        }
                        else
                        {
                            shape.TextFrame.TextRange.Text = item.SuggestedCorrection;
                        }
                    }
                }

                if (refresh) RefreshData();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not auto-fix item: " + ex.Message);
            }
        }

        private void SlideName_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBlock tb && tb.DataContext is UnifiedSlideSummaryItem slideItem)
            {
                try
                {
                    var app = _powerPointManager.GetApplication();
                    var pres = app.ActivePresentation;
                    if (pres == null) return;
                    
                    var slide = pres.Slides[slideItem.SlideNumber];
                    if (slide != null)
                    {
                        slide.Select();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error selecting slide: {ex.Message}");
                }
            }
        }
    }

    public class UnifiedErrorItem
    {
        public string ErrorCategory { get; set; } // "Style" or "Consistency"
        public string ErrorType { get; set; }
        public string Value { get; set; }
        public int SlideNumber { get; set; }
        public string ShapeName { get; set; }
        public string SuggestedCorrection { get; set; }
        public int TextStart { get; set; } = -1;
        public int TextLength { get; set; } = 0;

        public bool CanAutoCorrect => !string.IsNullOrEmpty(SuggestedCorrection);
        public bool IsColor => ErrorCategory == "Style" && ErrorType.Contains("Color");
    }

    public class UnifiedSlideSummaryItem
    {
        public string SlideName { get; set; }
        public int SlideNumber { get; set; }
        public bool IsPerfect { get; set; }
        public List<UnifiedErrorItem> Errors { get; set; }
    }

    public class SummaryGroupItem
    {
        public string GroupName { get; set; }
        public List<TypeSummaryItem> Items { get; set; }
    }

    public class TypeSummaryItem
    {
        public string Value { get; set; }
        public int Count { get; set; }
        public string CountText => $"{Count} instance{(Count > 1 ? "s" : "")}";
    }

    public class UnifiedCheckViewModel
    {
        public List<UnifiedSlideSummaryItem> BySlide { get; set; }
        public List<SummaryGroupItem> SummaryGroups { get; set; }

        public UnifiedCheckViewModel(List<StyleError> styleErrors, List<ConsistencyError> consistencyErrors, List<Slide> slides)
        {
            var unifiedErrors = new List<UnifiedErrorItem>();

            foreach (var se in styleErrors)
            {
                unifiedErrors.Add(new UnifiedErrorItem
                {
                    ErrorCategory = "Style",
                    ErrorType = se.ErrorType,
                    Value = se.Value,
                    SlideNumber = se.SlideNumber,
                    ShapeName = se.ShapeName,
                    SuggestedCorrection = se.SuggestedCorrectionValue,
                    TextStart = se.TextStart,
                    TextLength = se.TextLength
                });
            }

            foreach (var ce in consistencyErrors)
            {
                unifiedErrors.Add(new UnifiedErrorItem
                {
                    ErrorCategory = "Consistency",
                    ErrorType = ce.ErrorType,
                    Value = ce.Description,
                    SlideNumber = ce.SlideNumber,
                    ShapeName = ce.ShapeName,
                    SuggestedCorrection = ce.SuggestedCorrection,
                    TextStart = ce.TextStart,
                    TextLength = ce.TextLength
                });
            }

            // Group by slide
            var errorGroups = unifiedErrors.GroupBy(e => e.SlideNumber).ToDictionary(g => g.Key, g => g.ToList());
            var bySlideList = new List<UnifiedSlideSummaryItem>();

            foreach (var s in slides)
            {
                int num = s.SlideNumber;
                if (errorGroups.TryGetValue(num, out var slideErrors))
                {
                    bySlideList.Add(new UnifiedSlideSummaryItem
                    {
                        SlideNumber = num,
                        SlideName = "Slide " + num,
                        IsPerfect = false,
                        Errors = slideErrors.OrderBy(e => e.ErrorCategory).ThenBy(e => e.ErrorType).ToList()
                    });
                }
                else
                {
                    bySlideList.Add(new UnifiedSlideSummaryItem
                    {
                        SlideNumber = num,
                        SlideName = "Slide " + num,
                        IsPerfect = true,
                        Errors = new List<UnifiedErrorItem>()
                    });
                }
            }

            BySlide = bySlideList.OrderBy(x => x.SlideNumber).ToList();

            // Create Summary Groups
            SummaryGroups = unifiedErrors
                .GroupBy(e => e.ErrorType)
                .Select(g => new SummaryGroupItem
                {
                    GroupName = g.Key,
                    Items = g.GroupBy(x => x.Value)
                             .Select(vg => new TypeSummaryItem { Value = vg.Key, Count = vg.Count() })
                             .OrderByDescending(x => x.Count)
                             .ToList()
                })
                .OrderBy(g => g.GroupName)
                .ToList();
        }
    }
}