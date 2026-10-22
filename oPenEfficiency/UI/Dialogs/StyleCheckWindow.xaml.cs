using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using oPenEfficiency.Services;
using oPenEfficiency.Features;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    public partial class StyleCheckWindow : Window
    {
        private PowerPointManager _powerPointManager;
        private StyleCheckScope _scope;
        private bool _scopeSelectorReady;

        public StyleCheckWindow(
            PowerPointManager powerPointManager,
            StyleCheckScope scope = StyleCheckScope.SelectedSlides)
        {
            InitializeComponent();
            _powerPointManager = powerPointManager;
            SetScope(scope, false);
            _scopeSelectorReady = true;
            RefreshData();
        }

        public void SetScope(StyleCheckScope scope, bool refresh = true)
        {
            _scope = scope;
            if (ScopeSelector != null && ScopeSelector.SelectedIndex != (int)scope)
            {
                bool wasReady = _scopeSelectorReady;
                _scopeSelectorReady = false;
                ScopeSelector.SelectedIndex = (int)scope;
                _scopeSelectorReady = wasReady;
            }
            if (ScopeSummaryText != null)
            {
                ScopeSummaryText.Text = GetScopeDisplayName(scope);
            }
            if (refresh) RefreshData();
        }

        public void RefreshData()
        {
            try
            {
                var app = _powerPointManager.GetApplication();
                var slides = GetSlidesForScope(app, _scope);
                
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

        private void ScopeSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_scopeSelectorReady || ScopeSelector.SelectedIndex < 0) return;
            SetScope((StyleCheckScope)ScopeSelector.SelectedIndex);
        }

        private List<Slide> GetSlidesForScope(
            Microsoft.Office.Interop.PowerPoint.Application app,
            StyleCheckScope scope)
        {
            var slides = new List<Slide>();
            try
            {
                if (app.ActiveWindow == null) return slides;

                if (scope == StyleCheckScope.EntirePresentation)
                {
                    foreach (Slide slide in app.ActivePresentation.Slides)
                    {
                        slides.Add(slide);
                    }
                    return slides;
                }

                if (scope == StyleCheckScope.CurrentSlide)
                {
                    try { slides.Add((Slide)app.ActiveWindow.View.Slide); }
                    catch (Exception ex) { ReportOptionalStyleCheckFailure(ex, "StyleCheckWindow.GetSlides.CurrentSlide"); }
                    return slides;
                }

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
                    }
                    catch (Exception ex) { ReportOptionalStyleCheckFailure(ex, "StyleCheckWindow.GetSlides.SelectedSlides"); }
                }

                if (slides.Count == 0)
                {
                    try { slides.Add((Slide)app.ActiveWindow.View.Slide); }
                    catch (Exception ex) { ReportOptionalStyleCheckFailure(ex, "StyleCheckWindow.GetSlides.SelectionFallback"); }
                }
            }
            catch (Exception ex) { ReportOptionalStyleCheckFailure(ex, "StyleCheckWindow.GetSlides"); }
            return slides.Distinct().ToList();
        }

        private static string GetScopeDisplayName(StyleCheckScope scope)
        {
            switch (scope)
            {
                case StyleCheckScope.CurrentSlide:
                    return "Current slide";
                case StyleCheckScope.EntirePresentation:
                    return "Entire presentation";
                default:
                    return "Selected slides";
            }
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
                    
                    try { slide = pres.Slides[item.SlideNumber]; }
                    catch (Exception ex) { ReportOptionalStyleCheckFailure(ex, "StyleCheckWindow.ResultDoubleClick.FindSlide"); }
                    if (slide == null) return;

                    if (item.ShapeName == "Slide")
                    {
                        StyleCheckRuleIgnore.Ignore(slide, item.ErrorCategory, item.RuleId);
                    }
                    else
                    {
                        Shape shape = FindShapeByName(slide, item.ShapeName);
                        if (shape != null)
                        {
                            StyleCheckRuleIgnore.Ignore(shape, item.ErrorCategory, item.RuleId);
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

        
        private void FixSafeIssues_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is UnifiedCheckViewModel vm)
            {
                var fixableItems = vm.BySlide.SelectMany(s => s.Errors)
                    .Where(err => err.IsSafeAutoCorrect)
                    .OrderByDescending(err => err.TextStart)
                    .ToList();
                
                if (fixableItems.Count == 0)
                {
                    MessageBox.Show(
                        "No safe automatic corrections are available in the current scope.",
                        "Style Check",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                string preview = string.Join(
                    Environment.NewLine,
                    fixableItems
                        .GroupBy(item => item.ErrorType)
                        .OrderBy(group => group.Key)
                        .Select(group => $"• {group.Key}: {group.Count()}"));

                var result = MessageBox.Show(
                    $"Style Check can apply {fixableItems.Count} low-risk corrections:\n\n{preview}\n\n" +
                    "Only suggested brand-color replacements, double-space fixes, and repeated-word fixes are included. Continue?",
                    "Preview Safe Corrections",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);
                if (result != MessageBoxResult.Yes) return;

                try { _powerPointManager.GetApplication().StartNewUndoEntry(); }
                catch (Exception ex) { ReportOptionalStyleCheckFailure(ex, "StyleCheckWindow.FixAll.StartUndoEntry"); }
                foreach (var item in fixableItems)
                {
                    ApplyAutoFix(item, false);
                }
                RefreshData();
            }
        }

        private void Jump_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is UnifiedErrorItem item)
            {
                NavigateToItem(item);
            }
        }

        private void NavigateToItem(UnifiedErrorItem item)
        {
            try
            {
                var app = _powerPointManager.GetApplication();
                var slide = app.ActivePresentation.Slides[item.SlideNumber];
                app.ActiveWindow.View.GotoSlide(slide.SlideIndex);

                if (item.ShapeName == "Slide") return;

                Shape shape = FindShapeByName(slide, item.ShapeName);
                if (shape == null) return;

                shape.Select();
                if (item.TextStart > 0 &&
                    item.TextLength > 0 &&
                    shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    try
                    {
                        shape.TextFrame.TextRange
                            .Characters(item.TextStart, item.TextLength)
                            .Select();
                    }
                    catch (Exception ex) { ReportOptionalStyleCheckFailure(ex, "StyleCheckWindow.ApplyAutoFix.SelectText"); }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not navigate to the finding: " + ex.Message);
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
                try { slide = pres.Slides[item.SlideNumber]; }
                catch (Exception ex) { ReportOptionalStyleCheckFailure(ex, "StyleCheckWindow.ApplyAutoFix.FindSlide"); }
                if (slide == null) return;

                Shape shape = FindShapeByName(slide, item.ShapeName);
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

        private static Shape FindShapeByName(Slide slide, string shapeName)
        {
            if (slide == null || string.IsNullOrEmpty(shapeName)) return null;

            foreach (Shape shape in slide.Shapes)
            {
                Shape match = FindShapeByName(shape, shapeName);
                if (match != null) return match;
            }

            return null;
        }

        private static Shape FindShapeByName(Shape shape, string shapeName)
        {
            try
            {
                if (string.Equals(shape.Name, shapeName, StringComparison.Ordinal))
                {
                    return shape;
                }

                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    foreach (Shape child in shape.GroupItems)
                    {
                        Shape match = FindShapeByName(child, shapeName);
                        if (match != null) return match;
                    }
                }
            }
            catch (Exception ex) { ReportOptionalStyleCheckFailure(ex, "StyleCheckWindow.FindShapeByName"); }

            return null;
        }

        private static void ReportOptionalStyleCheckFailure(Exception exception, string context)
        {
            if (!ExceptionPolicy.IsExpectedComFailure(exception))
                ExceptionLogger.Log(exception, context);
            // A stale selection or shape proxy skips only the affected style-check item.
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
        public string RuleId { get; set; }
        public string ErrorCategory { get; set; } // "Style" or "Consistency"
        public string ErrorType { get; set; }
        public string Value { get; set; }
        public int SlideNumber { get; set; }
        public string ShapeName { get; set; }
        public string SuggestedCorrection { get; set; }
        public int TextStart { get; set; } = -1;
        public int TextLength { get; set; } = 0;

        public bool CanAutoCorrect => !string.IsNullOrEmpty(SuggestedCorrection);
        public bool IsSafeAutoCorrect =>
            CanAutoCorrect &&
            ((ErrorCategory == "Style" && IsColor) ||
             (ErrorCategory == "Consistency" &&
              (RuleId == "DoubleSpaces" || RuleId == "RepeatedWord")));
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
                    RuleId = se.RuleId ?? se.ErrorType,
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
                    RuleId = ce.RuleId ?? ce.ErrorType,
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
