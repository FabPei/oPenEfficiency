using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Features.Utilities;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class SeriesGeneratorWindow : Window
    {
        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
                this.DragMove();
        }

        private readonly PowerPointManager _manager;
        private Shape _templateShape;
        private Shape _boundaryShape;
        private SeriesConfig _existingConfig;

        public SeriesGeneratorWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            
            PopulateLanguages();
            
            LoadContext();
            UpdatePreview(null, null);
        }

        private void PopulateLanguages()
        {
            var cultures = System.Globalization.CultureInfo.GetCultures(System.Globalization.CultureTypes.SpecificCultures)
                .OrderBy(c => c.DisplayName)
                .ToList();
            
            CmbLanguage.ItemsSource = cultures;

            string defaultLang = "en-US";
            try
            {
                var app = _manager.GetApplication();
                if (app?.ActivePresentation != null)
                {
                    int lcid = (int)app.ActivePresentation.DefaultLanguageID;
                    
                    if (app.ActiveWindow?.Selection?.Type == PpSelectionType.ppSelectionShapes)
                    {
                        var sr = app.ActiveWindow.Selection.ShapeRange;
                        if (sr.Count > 0 && sr[1].HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            lcid = (int)sr[1].TextFrame.TextRange.LanguageID;
                        }
                    }
                    
                    if (lcid != 0) 
                    {
                        defaultLang = new System.Globalization.CultureInfo(lcid).Name;
                    }
                }
            }
            catch { }

            CmbLanguage.SelectedValue = defaultLang;
        }

        private void LoadContext()
        {
            var app = _manager.GetApplication();
            if (app?.ActiveWindow?.Selection?.Type == PpSelectionType.ppSelectionShapes)
            {
                var sr = app.ActiveWindow.Selection.ShapeRange;
                if (sr.Count > 0)
                {
                    _templateShape = sr[1];
                    if (sr.Count >= 2)
                    {
                        _boundaryShape = sr[2];
                    }

                    // Check for existing tags
                    try
                    {
                        string configJson = "";
                        foreach (string tagName in _templateShape.Tags)
                        {
                            if (tagName == SeriesGeneratorFeature.TagSeriesConfig)
                            {
                                configJson = _templateShape.Tags[tagName];
                                break;
                            }
                        }

                        if (!string.IsNullOrEmpty(configJson))
                        {
                            _existingConfig = JsonService.Deserialize<SeriesConfig>(configJson);
                            if (_existingConfig != null)
                            {
                                PopulateFromConfig(_existingConfig);
                                BtnGenerate.Content = "Update Series";
                                this.Title = "Update Series";
                            }
                        }
                    }
                    catch { }
                }
            }
        }

        private void PopulateFromConfig(SeriesConfig config)
        {
            // Disable events while populating
            CmbSeriesType.SelectionChanged -= CmbSeriesType_SelectionChanged;
            TxtStartValue.TextChanged -= UpdatePreview;
            TxtEndValue.TextChanged -= UpdatePreview;
            TxtCount.TextChanged -= UpdatePreview;
            TxtStep.TextChanged -= UpdatePreview;
            TxtFormat.TextChanged -= UpdatePreview;

            SelectComboBoxItemByTag(CmbSeriesType, config.SeriesType);
            SelectComboBoxItemByTag(CmbLanguage, config.Language);
            TxtStartValue.Text = config.StartValue;
            TxtEndValue.Text = config.EndValue;
            TxtCount.Text = config.Count.ToString();
            TxtStep.Text = config.Step.ToString();
            TxtFormat.Text = config.FormatString;

            SelectComboBoxItemByTag(CmbDirection, config.Direction);
            TxtMatrixColumns.Text = config.MatrixColumns.ToString();
            TxtSpacingX.Text = config.SpacingX.ToString();
            TxtSpacingY.Text = config.SpacingY.ToString();
            
            ChkAutoFit.IsChecked = config.AutoFitToBounds;
            SelectComboBoxItemByTag(CmbBoundaryType, config.BoundaryType);

            // Re-enable
            CmbSeriesType.SelectionChanged += CmbSeriesType_SelectionChanged;
            TxtStartValue.TextChanged += UpdatePreview;
            TxtEndValue.TextChanged += UpdatePreview;
            TxtCount.TextChanged += UpdatePreview;
            TxtStep.TextChanged += UpdatePreview;
            TxtFormat.TextChanged += UpdatePreview;
            
            PopulateFormatPresets(config.SeriesType);
            SelectComboBoxItemByTag(CmbFormatPresets, config.FormatString);
        }

        private void SelectComboBoxItemByTag(ComboBox cmb, string tag)
        {
            if (cmb == null) return;
            foreach (var item in cmb.Items)
            {
                if (item is ComboBoxItem cbi && cbi.Tag?.ToString() == tag)
                {
                    cmb.SelectedItem = cbi;
                    return;
                }
            }
        }

        private SeriesConfig GetCurrentConfig()
        {
            var config = new SeriesConfig();
            
            if (_existingConfig != null) config.SeriesId = _existingConfig.SeriesId;

            config.SeriesType = (CmbSeriesType.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "Numbers";
            config.Direction = (CmbDirection.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "Horizontal";
            config.StartValue = TxtStartValue?.Text;
            config.EndValue = TxtEndValue?.Text;
            config.FormatString = TxtFormat?.Text;

            if (TxtCount != null) { int.TryParse(TxtCount.Text, out int c); config.Count = c > 0 ? c : 5; }
            if (TxtStep != null) { int.TryParse(TxtStep.Text, out int s); config.Step = s > 0 ? s : 1; }

            if (config.Direction == "Matrix" && TxtMatrixColumns != null)
            {
                int.TryParse(TxtMatrixColumns.Text, out int mc);
                config.MatrixColumns = mc > 0 ? mc : 3;
            }

            if (TxtSpacingX != null) { double.TryParse(TxtSpacingX.Text, out double sx); config.SpacingX = sx; }
            if (TxtSpacingY != null) { double.TryParse(TxtSpacingY.Text, out double sy); config.SpacingY = sy; }

            config.AutoFitToBounds = ChkAutoFit?.IsChecked ?? true;
            config.BoundaryType = (CmbBoundaryType?.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "Slide";

            return config;
        }

        private void CmbDirection_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PanelMatrixCols == null) return;
            var tag = (CmbDirection.SelectedItem as ComboBoxItem)?.Tag?.ToString();
            PanelMatrixCols.Visibility = tag == "Matrix" ? Visibility.Visible : Visibility.Collapsed;
        }

        private void CmbLanguage_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdatePreview(null, null);
        }
        
        private void CmbBoundaryType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TxtBoundaryWarning == null) return;
            var tag = (CmbBoundaryType.SelectedItem as ComboBoxItem)?.Tag?.ToString();
            
            if (tag == "SelectedShape")
            {
                if (_boundaryShape == null)
                {
                    TxtBoundaryWarning.Visibility = Visibility.Visible;
                    if (BtnGenerate != null) BtnGenerate.IsEnabled = false;
                }
                else
                {
                    TxtBoundaryWarning.Visibility = Visibility.Collapsed;
                    UpdatePreview(null, null); // re-enable generate button
                }
            }
            else
            {
                TxtBoundaryWarning.Visibility = Visibility.Collapsed;
                UpdatePreview(null, null); // re-enable generate button
            }
        }
        
        private void CmbSeriesType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tag = (CmbSeriesType.SelectedItem as ComboBoxItem)?.Tag?.ToString();
            if (tag == null) return;
            
            if (PanelLanguage != null)
            {
                PanelLanguage.Visibility = (tag == "Days" || tag == "Months" || tag == "Years" || tag == "Quarters") ? Visibility.Visible : Visibility.Collapsed;
            }

            // Set defaults
            if (TxtStartValue != null)
            {
                TxtStartValue.TextChanged -= UpdatePreview;
                switch (tag)
                {
                    case "Numbers": TxtStartValue.Text = "1"; break;
                    case "Letters": TxtStartValue.Text = "A"; break;
                    case "Days": case "Months": case "Years": TxtStartValue.Text = DateTime.Now.ToString("yyyy-MM-dd"); break;
                    case "Quarters": TxtStartValue.Text = $"{DateTime.Now.Year}-Q1"; break;
                }
                TxtStartValue.TextChanged += UpdatePreview;
            }

            PopulateFormatPresets(tag);
            UpdatePreview(null, null);
        }

        private void PopulateFormatPresets(string seriesType)
        {
            if (CmbFormatPresets == null) return;
            
            CmbFormatPresets.SelectionChanged -= CmbFormatPresets_SelectionChanged;
            CmbFormatPresets.Items.Clear();

            var addPreset = new Action<string, string>((label, format) => 
            {
                CmbFormatPresets.Items.Add(new ComboBoxItem { Content = label, Tag = format });
            });

            addPreset("Default", "");

            switch (seriesType)
            {
                case "Numbers":
                    addPreset("01, 02, 03...", "00");
                    addPreset("001, 002, 003...", "000");
                    addPreset("Step 1, Step 2...", "Step {0}");
                    addPreset("Item 1, Item 2...", "Item {0}");
                    break;
                case "Letters":
                    addPreset("A), B), C)...", "{0})");
                    addPreset("Option A, Option B...", "Option {0}");
                    break;
                case "Days":
                    addPreset("01, 02, 03 (Day only)", "dd");
                    addPreset("Mon, Tue, Wed (Short name)", "ddd");
                    addPreset("Monday, Tuesday (Full name)", "dddd");
                    addPreset("01.01, 02.01", "dd.MM");
                    addPreset("Jan 01, Jan 02", "MMM dd");
                    break;
                case "Months":
                    addPreset("01, 02, 03 (Month num)", "MM");
                    addPreset("Jan, Feb, Mar (Short)", "MMM");
                    addPreset("January, February (Full)", "MMMM");
                    addPreset("Jan 24, Feb 24", "MMM yy");
                    addPreset("January 2024", "MMMM yyyy");
                    break;
                case "Years":
                    addPreset("24, 25, 26 (Short)", "yy");
                    addPreset("2024, 2025 (Full)", "yyyy");
                    addPreset("FY24, FY25", "FYyy");
                    break;
                case "Quarters":
                    addPreset("Q1, Q2, Q3", "Q{q}");
                    addPreset("Q1 24, Q2 24", "Q{q} {yy}");
                    addPreset("Q1 2024, Q2 2024", "Q{q} {yyyy}");
                    addPreset("Quarter 1", "Quarter {q}");
                    break;
            }

            CmbFormatPresets.SelectedIndex = 0;
            CmbFormatPresets.SelectionChanged += CmbFormatPresets_SelectionChanged;
        }

        private void CmbFormatPresets_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tag = (CmbFormatPresets.SelectedItem as ComboBoxItem)?.Tag?.ToString();
            if (tag != null && TxtFormat != null)
            {
                TxtFormat.Text = tag;
            }
        }

        private void UpdatePreview(object sender, RoutedEventArgs e)
        {
            if (TxtPreview == null) return;

            var config = GetCurrentConfig();
            var data = SeriesGeneratorFeature.GenerateSeriesData(config);

            if (data == null || data.Count == 0)
            {
                TxtPreview.Text = "[Invalid settings or empty result]";
                TxtPreview.Foreground = System.Windows.Media.Brushes.OrangeRed;
                if (BtnGenerate != null) BtnGenerate.IsEnabled = false;
            }
            else
            {
                int maxPreview = 10;
                var previewItems = data.Take(maxPreview).Select(x => $"[{x}]").ToList();
                if (data.Count > maxPreview) previewItems.Add("...");

                TxtPreview.Text = string.Join(" → ", previewItems);
                TxtPreview.Foreground = (System.Windows.Media.Brush)FindResource("TextPrimaryBrush");
                if (BtnGenerate != null) BtnGenerate.IsEnabled = true;
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            var config = GetCurrentConfig();
            SeriesGeneratorFeature.GenerateShapes(_manager, config, _templateShape, _boundaryShape);
            this.Close();
        }
    }
}
