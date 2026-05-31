using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using oPenEfficiency;
using oPenEfficiency.Models;
using System.Collections.Generic;

namespace oPenEfficiency.UI
{
    public class ShortcutItemViewModel : INotifyPropertyChanged
    {
        public string FeatureId { get; set; }
        public string FeatureName { get; set; }
        
        private string _shortcutText;
        public string ShortcutText 
        { 
            get => _shortcutText; 
            set { _shortcutText = value; OnPropertyChanged(); } 
        }

        public ICommand EditCommand { get; set; }
        public ICommand ClearCommand { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public class SettingsViewModel : INotifyPropertyChanged
    {
        private SidebarConfig _config;
        private string _searchText;
        private SidebarFeature _selectedPoolFeature;
        private string _shortcutSearchText;

        public SidebarConfig Config
        {
            get => _config;
            set { _config = value; OnPropertyChanged(); OnPropertyChanged(nameof(BackgroundColor)); OnPropertyChanged(nameof(SectionFontColor)); OnPropertyChanged(nameof(SectionFontSize)); OnPropertyChanged(nameof(ButtonBackgroundColor)); OnPropertyChanged(nameof(StickyNoteStyle)); OnPropertyChanged(nameof(StickyNoteDefaultText)); OnPropertyChanged(nameof(StickyNoteFont)); OnPropertyChanged(nameof(StickyNoteBackgroundColor)); OnPropertyChanged(nameof(StickyNoteFontColor)); OnPropertyChanged(nameof(SpacingIncrement)); OnPropertyChanged(nameof(IconSize)); OnPropertyChanged(nameof(AppTheme)); OnPropertyChanged(nameof(DeepLApiKey)); OnPropertyChanged(nameof(TranslationTargetLanguage)); }
        }

        public string DeepLApiKey
        {
            get => _config?.DeepLApiKey ?? "";
            set
            {
                if (_config != null && _config.DeepLApiKey != value)
                {
                    _config.DeepLApiKey = value;
                    OnPropertyChanged();
                }
            }
        }

        public string TranslationTargetLanguage
        {
            get => _config?.TranslationTargetLanguage ?? "EN-US";
            set
            {
                if (_config != null && _config.TranslationTargetLanguage != value)
                {
                    _config.TranslationTargetLanguage = value;
                    OnPropertyChanged();
                }
            }
        }

        public string SelectedTranslationTool
        {
            get => _config?.SelectedTranslationTool ?? "DeepL";
            set
            {
                if (_config != null && _config.SelectedTranslationTool != value)
                {
                    _config.SelectedTranslationTool = value;
                    OnPropertyChanged();
                }
            }
        }

        public string OpenAIApiKey
        {
            get => _config?.OpenAIApiKey ?? "";
            set
            {
                if (_config != null && _config.OpenAIApiKey != value)
                {
                    _config.OpenAIApiKey = value;
                    OnPropertyChanged();
                }
            }
        }

        public string OpenAIModel
        {
            get => _config?.OpenAIModel ?? "gpt-4o";
            set
            {
                if (_config != null && _config.OpenAIModel != value)
                {
                    _config.OpenAIModel = value;
                    OnPropertyChanged();
                }
            }
        }

        public string ClaudeApiKey
        {
            get => _config?.ClaudeApiKey ?? "";
            set
            {
                if (_config != null && _config.ClaudeApiKey != value)
                {
                    _config.ClaudeApiKey = value;
                    OnPropertyChanged();
                }
            }
        }

        public string ClaudeModel
        {
            get => _config?.ClaudeModel ?? "claude-3-5-sonnet-20240620";
            set
            {
                if (_config != null && _config.ClaudeModel != value)
                {
                    _config.ClaudeModel = value;
                    OnPropertyChanged();
                }
            }
        }

        public string LibreTranslateBaseUrl
        {
            get => _config?.LibreTranslateBaseUrl ?? "https://libretranslate.com/";
            set
            {
                if (_config != null && _config.LibreTranslateBaseUrl != value)
                {
                    _config.LibreTranslateBaseUrl = value;
                    OnPropertyChanged();
                }
            }
        }

        public string LibreTranslateApiKey
        {
            get => _config?.LibreTranslateApiKey ?? "";
            set
            {
                if (_config != null && _config.LibreTranslateApiKey != value)
                {
                    _config.LibreTranslateApiKey = value;
                    OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<string> AvailableTranslationTools { get; } = new ObservableCollection<string> { "DeepL", "OpenAI", "Claude", "LibreTranslate" };
        public ObservableCollection<string> AvailableOpenAIModels { get; } = new ObservableCollection<string> { "gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "gpt-3.5-turbo" };
        public ObservableCollection<string> AvailableClaudeModels { get; } = new ObservableCollection<string> { "claude-3-5-sonnet-20240620", "claude-3-opus-20240229", "claude-3-sonnet-20240229", "claude-3-haiku-20240307" };

        public ObservableCollection<string> AvailableTargetLanguages { get; } = new ObservableCollection<string>
        {
            "BG", "CS", "DA", "DE", "EL", "EN-GB", "EN-US", "ES", "ET", "FI", "FR", "HU", "ID", "IT", "JA", "KO", "LT", "LV", "NB", "NL", "PL", "PT-BR", "PT-PT", "RO", "RU", "SK", "SL", "SV", "TR", "UK", "ZH"
        };

        public string AppTheme
        {
            get => _config?.AppTheme ?? "Dark";
            set
            {
                if (_config != null && _config.AppTheme != value)
                {
                    _config.AppTheme = value;
                    oPenEfficiency.Services.ThemeManager.ApplyTheme(value);
                    OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<string> AvailableThemes { get; } = new ObservableCollection<string>();

        private void InitializeAvailableThemes()
        {
            AvailableThemes.Clear();
            AvailableThemes.Add("Dark");
            AvailableThemes.Add("Light");
            
            // Add company theme if it's already in the config and not a default one
            if (_config != null && !string.IsNullOrEmpty(_config.AppTheme))
            {
                if (_config.AppTheme != "Dark" && _config.AppTheme != "Light" && !AvailableThemes.Contains(_config.AppTheme))
                {
                    AvailableThemes.Add(_config.AppTheme);
                }
            }
        }

        public string AppFontFamily
        {
            get => _config?.AppFontFamily ?? "Segoe UI";
            set
            {
                if (_config != null && _config.AppFontFamily != value)
                {
                    _config.AppFontFamily = value;
                    OnPropertyChanged();
                }
            }
        }

        public string BackgroundColor
        {
            get => _config?.BackgroundColor ?? "Transparent";
            set
            {
                if (_config != null && _config.BackgroundColor != value)
                {
                    _config.BackgroundColor = value;
                    OnPropertyChanged();
                }
            }
        }

        public string SectionFontColor
        {
            get => _config?.SectionFontColor ?? "#888888";
            set
            {
                if (_config != null && _config.SectionFontColor != value)
                {
                    _config.SectionFontColor = value;
                    OnPropertyChanged();
                }
            }
        }

        public double SectionFontSize
        {
            get => _config?.SectionFontSize > 0 ? _config.SectionFontSize : 9;
            set
            {
                if (_config != null && Math.Abs(_config.SectionFontSize - value) > 0.01)
                {
                    _config.SectionFontSize = value;
                    OnPropertyChanged();
                }
            }
        }

        public string ButtonBackgroundColor
        {
            get => _config?.ButtonBackgroundColor ?? "Transparent";
            set
            {
                if (_config != null && _config.ButtonBackgroundColor != value)
                {
                    _config.ButtonBackgroundColor = value;
                    OnPropertyChanged();
                }
            }
        }

        public string StickyNoteStyle
        {
            get => _config?.StickyNoteStyle ?? "Theme";
            set
            {
                if (_config != null && _config.StickyNoteStyle != value)
                {
                    _config.StickyNoteStyle = value;
                    OnPropertyChanged();
                }
            }
        }

        public string StickyNoteDefaultText
        {
            get => _config?.StickyNoteDefaultText ?? "";
            set
            {
                if (_config != null && _config.StickyNoteDefaultText != value)
                {
                    _config.StickyNoteDefaultText = value;
                    OnPropertyChanged();
                }
            }
        }

        public string MasterObjectMode
        {
            get => _config?.MasterObjectMode ?? "First Selected";
            set
            {
                if (_config != null && _config.MasterObjectMode != value)
                {
                    _config.MasterObjectMode = value;
                    OnPropertyChanged();
                }
            }
        }

        public string StyleCheckCorrectColors
        {
            get => _config?.StyleCheckCorrectColors ?? "";
            set
            {
                if (_config != null && _config.StyleCheckCorrectColors != value)
                {
                    _config.StyleCheckCorrectColors = value;
                    OnPropertyChanged();
                }
            }
        }

        public string StyleCheckCorrectFonts
        {
            get => _config?.StyleCheckCorrectFonts ?? "";
            set
            {
                if (_config != null && _config.StyleCheckCorrectFonts != value)
                {
                    _config.StyleCheckCorrectFonts = value;
                    OnPropertyChanged();
                }
            }
        }

        public string StickyNoteFont
        {
            get => string.IsNullOrEmpty(_config?.StickyNoteFont) ? "Theme Font" : _config.StickyNoteFont;
            set
            {
                if (_config != null && _config.StickyNoteFont != value)
                {
                    _config.StickyNoteFont = value == "Theme Font" ? "" : value;
                    OnPropertyChanged();
                }
            }
        }

        public string StickyNoteBackgroundColor
        {
            get => _config?.StickyNoteBackgroundColor ?? "#94F4FF";
            set
            {
                if (_config != null && _config.StickyNoteBackgroundColor != value)
                {
                    _config.StickyNoteBackgroundColor = value;
                    OnPropertyChanged();
                }
            }
        }

        public string StickyNoteFontColor
        {
            get => _config?.StickyNoteFontColor ?? "#323232";
            set
            {
                if (_config != null && _config.StickyNoteFontColor != value)
                {
                    _config.StickyNoteFontColor = value;
                    OnPropertyChanged();
                }
            }
        }

        public string IllustrativeStickerColor
        {
            get => _config?.IllustrativeStickerColor ?? "#FF0000";
            set
            {
                if (_config != null && _config.IllustrativeStickerColor != value)
                {
                    _config.IllustrativeStickerColor = value;
                    OnPropertyChanged();
                }
            }
        }

        public string IllustrativeStickerTexts
        {
            get => _config?.IllustrativeStickerTexts ?? "Work in progress, Updated, For discussion, Backup, Confidential, Strictly confidential";
            set
            {
                if (_config != null && _config.IllustrativeStickerTexts != value)
                {
                    _config.IllustrativeStickerTexts = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool EnableSwapPositionsContextMenu
        {
            get => _config?.EnableSwapPositionsContextMenu ?? true;
            set
            {
                if (_config != null && _config.EnableSwapPositionsContextMenu != value)
                {
                    _config.EnableSwapPositionsContextMenu = value;
                    OnPropertyChanged();
                }
            }
        }

        public double IllustrativeStickerLineSize
        {
            get => _config?.IllustrativeStickerLineSize ?? 3.0;
            set
            {
                if (_config != null && _config.IllustrativeStickerLineSize != value)
                {
                    _config.IllustrativeStickerLineSize = value;
                    OnPropertyChanged();
                }
            }
        }

        public double IllustrativeStickerFontSize
        {
            get => _config?.IllustrativeStickerFontSize ?? 14;
            set
            {
                if (_config != null && _config.IllustrativeStickerFontSize != value)
                {
                    _config.IllustrativeStickerFontSize = value;
                    OnPropertyChanged();
                }
            }
        }

        public string IllustrativeStickerFont
        {
            get => string.IsNullOrEmpty(_config?.IllustrativeStickerFont) ? "Theme Font" : _config.IllustrativeStickerFont;
            set
            {
                if (_config != null && _config.IllustrativeStickerFont != value)
                {
                    _config.IllustrativeStickerFont = value == "Theme Font" ? "" : value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IllustrativeStickerFontBold
        {
            get => _config?.IllustrativeStickerFontBold ?? true;
            set
            {
                if (_config != null && _config.IllustrativeStickerFontBold != value)
                {
                    _config.IllustrativeStickerFontBold = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IllustrativeStickerFontItalic
        {
            get => _config?.IllustrativeStickerFontItalic ?? false;
            set
            {
                if (_config != null && _config.IllustrativeStickerFontItalic != value)
                {
                    _config.IllustrativeStickerFontItalic = value;
                    OnPropertyChanged();
                }
            }
        }

        public double SpacingIncrement
        {
            get => _config?.SpacingIncrement ?? 10.0;
            set
            {
                if (_config != null && Math.Abs(_config.SpacingIncrement - value) > 0.01)
                {
                    _config.SpacingIncrement = value;
                    OnPropertyChanged();
                }
            }
        }

        public double IconSize
        {
            get => _config?.IconSize > 0 ? _config.IconSize : 12;
            set
            {
                if (_config != null && Math.Abs(_config.IconSize - value) > 0.01)
                {
                    _config.IconSize = value;
                    OnPropertyChanged();
                }
            }
        }


        public ObservableCollection<string> AvailableFonts { get; } = new ObservableCollection<string> 
        { 
            "Theme Font", "Arial", "Calibri", "Comic Sans MS", "Courier New", "Georgia", "Impact", "Segoe UI", "Tahoma", "Times New Roman", "Trebuchet MS", "Verdana" 
        };

        public ObservableCollection<string> AvailableMasterObjectModes { get; } = new ObservableCollection<string> 
        { 
            "First Selected", 
            "Last Selected" 
        };

        public string SearchText
        {
            get => _searchText;
            set 
            { 
                _searchText = value; 
                OnPropertyChanged(); 
                OnPropertyChanged(nameof(FilteredPool)); 
            }
        }

        public SidebarFeature SelectedPoolFeature
        {
            get => _selectedPoolFeature;
            set { _selectedPoolFeature = value; OnPropertyChanged(); }
        }

        public ObservableCollection<SidebarFeature> FilteredPool
        {
            get
            {
                var pool = FeatureLibrary.AllFeatures;
                if (string.IsNullOrWhiteSpace(SearchText)) return new ObservableCollection<SidebarFeature>(pool.OrderBy(f => f.Name));
                return new ObservableCollection<SidebarFeature>(pool.Where(f => f.Name.IndexOf(SearchText, StringComparison.OrdinalIgnoreCase) >= 0).OrderBy(f => f.Name));
            }
        }

        // Shortcuts Collection
        private List<ShortcutItemViewModel> _fullShortcutsList;
        public ObservableCollection<ShortcutItemViewModel> FilteredShortcuts { get; private set; } = new ObservableCollection<ShortcutItemViewModel>();

        public string ShortcutSearchText
        {
            get => _shortcutSearchText;
            set
            {
                _shortcutSearchText = value;
                OnPropertyChanged();
                RefreshFilteredShortcuts();
            }
        }

        private void LoadShortcutsList()
        {
            _fullShortcutsList = new List<ShortcutItemViewModel>();
            var customShortcuts = ShortcutManager.GetShortcuts();

            foreach (var feature in FeatureLibrary.AllFeatures)
            {
                string text = customShortcuts.ContainsKey(feature.Id) ? customShortcuts[feature.Id] : "Not Assigned";
                
                var item = new ShortcutItemViewModel
                {
                    FeatureId = feature.Id,
                    FeatureName = feature.Name,
                    ShortcutText = text,
                    EditCommand = new RelayCommand(_ => {
                        var w = new ShortcutConfigWindow(feature.Id);
                        w.Owner = System.Windows.Application.Current.Windows.OfType<oPenEfficiency.UI.SettingsWindow>().FirstOrDefault();
                        w.ShowDialog();
                        LoadShortcutsList(); // Refresh after edit
                    }),
                    ClearCommand = new RelayCommand(_ => {
                        ShortcutManager.ClearShortcut(feature.Id);
                        LoadShortcutsList(); // Refresh list
                    })
                };
                _fullShortcutsList.Add(item);
            }
            
            // Sort alphabetical by name
            _fullShortcutsList = _fullShortcutsList.OrderBy(x => x.FeatureName).ToList();
            RefreshFilteredShortcuts();
        }

        private void RefreshFilteredShortcuts()
        {
            FilteredShortcuts.Clear();
            if (_fullShortcutsList == null) return;

            foreach (var item in _fullShortcutsList)
            {
                if (string.IsNullOrWhiteSpace(ShortcutSearchText) || 
                    item.FeatureName.IndexOf(ShortcutSearchText, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    FilteredShortcuts.Add(item);
                }
            }
        }

        public ObservableCollection<SidebarSection> Sections => new ObservableCollection<SidebarSection>(Config?.Sections ?? new List<SidebarSection>());

        public ICommand AddSectionCommand { get; }
        public ICommand RemoveSectionCommand { get; }
        public ICommand AddFeatureToSectionCommand { get; }
        public ICommand ResetStickyStylesCommand { get; }
        public ICommand ResetSidebarStylesCommand { get; }
        public ICommand SaveCommand { get; }
        public ICommand CloseCommand { get; }
        public ICommand ResetToDefaultsCommand { get; }
        
        public ICommand ExportCommand { get; }
        public ICommand ImportCommand { get; }

        private bool _exportLayout = true;
        public bool ExportLayout { get => _exportLayout; set { _exportLayout = value; OnPropertyChanged(); } }
        private bool _exportColors = true;
        public bool ExportColors { get => _exportColors; set { _exportColors = value; OnPropertyChanged(); } }
        private bool _exportSticky = true;
        public bool ExportSticky { get => _exportSticky; set { _exportSticky = value; OnPropertyChanged(); } }
        private bool _excludeApiKeys = true;
        public bool ExcludeApiKeys { get => _excludeApiKeys; set { _excludeApiKeys = value; OnPropertyChanged(); } }

        public Action CloseAction { get; set; }
        public Action RebuildAction { get; set; }

        private string _selectedPreset = "Standard Layout (Default)";
        public string SelectedPreset
        {
            get => _selectedPreset;
            set { _selectedPreset = value; OnPropertyChanged(); }
        }

        public ObservableCollection<string> AvailablePresets { get; } = new ObservableCollection<string>
        {
            "Standard Layout (Default)",
            "All Features (Categorized)",
            "Minimalist (Essentials Only)"
        };

        public ICommand ApplyPresetCommand { get; private set; }
        public ICommand ExportLayoutOnlyCommand { get; private set; }
        public ICommand ImportLayoutOnlyCommand { get; private set; }
        
        public ICommand ReportBugCommand { get; }
        public ICommand RequestFeatureCommand { get; }

        public SettingsViewModel()
        {
            LoadConfig();
            InitializeAvailableThemes();
            LoadShortcutsList();

            ApplyPresetCommand = new RelayCommand(_ => {
                if (Config == null) return;
                SidebarConfig presetConfig = null;
                if (SelectedPreset == "All Features (Categorized)") presetConfig = FeatureLibrary.GetAllFeaturesLayout();
                else if (SelectedPreset == "Minimalist (Essentials Only)") presetConfig = FeatureLibrary.GetMinimalistLayout();
                else presetConfig = FeatureLibrary.GetStandardLayout();

                Config.Sections = new System.Collections.Generic.List<SidebarSection>(presetConfig.Sections);
                RefreshSections();
            });

            ExportLayoutOnlyCommand = new RelayCommand(_ => {
                var sfd = new Microsoft.Win32.SaveFileDialog {
                    Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                    Title = "Export Layout Only",
                    FileName = "open_efficiency_layout.json"
                };
                if (sfd.ShowDialog() == true) {
                    var layoutOnly = new SidebarConfig { Sections = Config.Sections };
                    string json = JsonService.Serialize(layoutOnly);
                    System.IO.File.WriteAllText(sfd.FileName, json);
                    System.Windows.MessageBox.Show("Layout exported successfully.", "Export Complete", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                }
            });

            ImportLayoutOnlyCommand = new RelayCommand(_ => {
                var ofd = new Microsoft.Win32.OpenFileDialog {
                    Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                    Title = "Import Layout Only"
                };
                if (ofd.ShowDialog() == true) {
                    try {
                        string json = System.IO.File.ReadAllText(ofd.FileName);
                        var imported = JsonService.Deserialize<SidebarConfig>(json);
                        if (imported != null && imported.Sections != null) {
                            int missingCount = 0;
                            foreach (var s in imported.Sections) {
                                if (s.Features != null) {
                                    int removed = s.Features.RemoveAll(f => FeatureLibrary.AllFeatures.Find(x => x.Id == f.Id) == null);
                                    missingCount += removed;
                                }
                            }
                            Config.Sections = new System.Collections.Generic.List<SidebarSection>(imported.Sections);
                            RefreshSections();
                            
                            if (missingCount > 0) {
                                System.Windows.MessageBox.Show($"Layout imported successfully, but {missingCount} feature(s) from the import file were not recognized by this version of oPenEfficiency and have been skipped.", "Import Warning", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                            }
                        }
                    } catch (Exception ex) {
                        System.Windows.MessageBox.Show("Error importing layout: " + ex.Message);
                    }
                }
            });

            ResetStickyStylesCommand = new RelayCommand(_ => {
                StickyNoteBackgroundColor = "#FFF494";
                StickyNoteFontColor = "#323232";
            });

            ResetSidebarStylesCommand = new RelayCommand(_ => {
                BackgroundColor = "Transparent";
                SectionFontColor = "#9CA3AF";
                ButtonBackgroundColor = "Transparent";
                SectionFontSize = 9;
                IconSize = 16;
                AppTheme = "Dark";
                InitializeAvailableThemes(); // Ensure list is correct after reset
            });

            AddSectionCommand = new RelayCommand(_ => {
                Config.Sections.Add(new SidebarSection { Name = "New Section" });
                RefreshSections();
            });

            RemoveSectionCommand = new RelayCommand(param => {
                if (param is SidebarSection section) {
                    Config.Sections.Remove(section);
                    RefreshSections();
                }
            });

            AddFeatureToSectionCommand = new RelayCommand(param => {
                if (param is SidebarSection section && SelectedPoolFeature != null) {
                    if (!section.Features.Any(f => f.Id == SelectedPoolFeature.Id)) {
                        section.Features.Add(new SidebarFeature { Id = SelectedPoolFeature.Id, Name = SelectedPoolFeature.Name });
                        RefreshSections();
                    }
                }
            }, _ => SelectedPoolFeature != null);

            SaveCommand = new RelayCommand(_ => {
                JsonService.Save(Config, "sidebar_config.json");
                RebuildAction?.Invoke();
                CloseAction?.Invoke();
            });

            CloseCommand = new RelayCommand(_ => CloseAction?.Invoke());

            ResetToDefaultsCommand = new RelayCommand(_ => {
                var result = System.Windows.MessageBox.Show(
                    "Are you sure you want to restore all settings to their default values?\n\nThis will reset your layout, appearance, and shortcuts.",
                    "Confirm Reset",
                    System.Windows.MessageBoxButton.YesNo,
                    System.Windows.MessageBoxImage.Warning);

                if (result == System.Windows.MessageBoxResult.Yes)
                {
                    Config = FeatureLibrary.GetEmergencyDefault();
                    RefreshSections();
                }
            });

            ExportCommand = new RelayCommand(_ => {
                var sfd = new Microsoft.Win32.SaveFileDialog {
                    Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                    Title = "Export Configuration",
                    FileName = "open_efficiency_style.json"
                };

                if (sfd.ShowDialog() == true) {
                    try {
                        var exportData = new SidebarConfig();
                        
                        // Default the generic properties we're not exporting
                        exportData.Sections = null;
                        exportData.BackgroundColor = null;
                        exportData.SectionFontColor = null;
                        exportData.ButtonBackgroundColor = null;
                        exportData.SectionFontSize = 0;
                        exportData.IconSize = 0;
                        exportData.StickyNoteStyle = null;
                        exportData.StickyNoteDefaultText = null;
                        exportData.StickyNoteFont = null;

                        if (ExportLayout) {
                            var json = JsonService.Serialize(Config.Sections);
                            exportData.Sections = JsonService.Deserialize<List<SidebarSection>>(json);
                        }
                        if (ExportColors) {
                            exportData.BackgroundColor = Config.BackgroundColor;
                            exportData.SectionFontColor = Config.SectionFontColor;
                            exportData.ButtonBackgroundColor = Config.ButtonBackgroundColor;
                            exportData.SectionFontSize = Config.SectionFontSize;
                            exportData.IconSize = Config.IconSize;
                            exportData.SpacingIncrement = Config.SpacingIncrement;
                            exportData.AppTheme = Config.AppTheme;
                        }
                        if (ExportSticky) {
                            exportData.StickyNoteStyle = Config.StickyNoteStyle;
                            exportData.StickyNoteDefaultText = Config.StickyNoteDefaultText;
                            exportData.StickyNoteFont = Config.StickyNoteFont;
                            exportData.StickyNoteBackgroundColor = Config.StickyNoteBackgroundColor;
                            exportData.StickyNoteFontColor = Config.StickyNoteFontColor;
                            exportData.IllustrativeStickerColor = Config.IllustrativeStickerColor;
                            exportData.IllustrativeStickerLineSize = Config.IllustrativeStickerLineSize;
                            exportData.IllustrativeStickerFontSize = Config.IllustrativeStickerFontSize;
                            exportData.IllustrativeStickerFont = Config.IllustrativeStickerFont;
                            exportData.IllustrativeStickerFontBold = Config.IllustrativeStickerFontBold;
                            exportData.IllustrativeStickerFontItalic = Config.IllustrativeStickerFontItalic;
                        }

                        // Always include translation settings, but exclude keys if requested
                        exportData.TranslationTargetLanguage = Config.TranslationTargetLanguage;
                        exportData.SelectedTranslationTool = Config.SelectedTranslationTool;
                        exportData.OpenAIModel = Config.OpenAIModel;
                        exportData.ClaudeModel = Config.ClaudeModel;

                        exportData.LibreTranslateBaseUrl = Config.LibreTranslateBaseUrl;

                        if (ExcludeApiKeys)
                        {
                            exportData.DeepLApiKey = "";
                            exportData.OpenAIApiKey = "";
                            exportData.ClaudeApiKey = "";
                            exportData.LibreTranslateApiKey = "";
                        }
                        else
                        {
                            exportData.DeepLApiKey = Config.DeepLApiKey;
                            exportData.OpenAIApiKey = Config.OpenAIApiKey;
                            exportData.ClaudeApiKey = Config.ClaudeApiKey;
                            exportData.LibreTranslateApiKey = Config.LibreTranslateApiKey;
                        }

                        string resultJson = JsonService.Serialize(exportData);
                        System.IO.File.WriteAllText(sfd.FileName, resultJson);
                        System.Windows.MessageBox.Show("Configuration exported successfully.", "Export Complete", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                    } catch (Exception ex) {
                        System.Windows.MessageBox.Show($"Error exporting configuration: {ex.Message}");
                    }
                }
            });

            ImportCommand = new RelayCommand(_ => {
                var ofd = new Microsoft.Win32.OpenFileDialog {
                    Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                    Title = "Import Configuration"
                };

                if (ofd.ShowDialog() == true) {
                    try {
                        string json = System.IO.File.ReadAllText(ofd.FileName);
                        var imported = JsonService.Deserialize<SidebarConfig>(json);
                        if (imported != null) {
                            if (imported.Sections != null && imported.Sections.Count > 0) {
                                Config.Sections = new List<SidebarSection>(imported.Sections);
                                RefreshSections();
                            }
                            if (!string.IsNullOrEmpty(imported.BackgroundColor)) BackgroundColor = imported.BackgroundColor;
                            if (!string.IsNullOrEmpty(imported.SectionFontColor)) SectionFontColor = imported.SectionFontColor;
                            if (!string.IsNullOrEmpty(imported.ButtonBackgroundColor)) ButtonBackgroundColor = imported.ButtonBackgroundColor;
                            if (imported.SectionFontSize > 0) SectionFontSize = imported.SectionFontSize;
                            if (imported.IconSize > 0) Config.IconSize = imported.IconSize;
                            if (imported.SpacingIncrement > 0) SpacingIncrement = imported.SpacingIncrement;
                            if (imported.AppTheme != null) AppTheme = imported.AppTheme;
                            if (imported.AppFontFamily != null) AppFontFamily = imported.AppFontFamily;
                            InitializeAvailableThemes(); // Ensure theme list reflects imported config
                            if (imported.StickyNoteStyle != null) StickyNoteStyle = imported.StickyNoteStyle;
                            if (imported.StickyNoteDefaultText != null) StickyNoteDefaultText = imported.StickyNoteDefaultText;
                            if (imported.StickyNoteFont != null) StickyNoteFont = string.IsNullOrEmpty(imported.StickyNoteFont) ? "Theme Font" : imported.StickyNoteFont;
                            if (!string.IsNullOrEmpty(imported.StickyNoteBackgroundColor)) StickyNoteBackgroundColor = imported.StickyNoteBackgroundColor;
                            if (!string.IsNullOrEmpty(imported.StickyNoteFontColor)) StickyNoteFontColor = imported.StickyNoteFontColor;
                            if (imported.IllustrativeStickerTexts != null) IllustrativeStickerTexts = imported.IllustrativeStickerTexts;
                            if (imported.IllustrativeStickerColor != null) IllustrativeStickerColor = imported.IllustrativeStickerColor;
                            if (imported.IllustrativeStickerLineSize > 0) IllustrativeStickerLineSize = imported.IllustrativeStickerLineSize;
                            if (imported.IllustrativeStickerFontSize > 0) IllustrativeStickerFontSize = imported.IllustrativeStickerFontSize;
                            if (imported.IllustrativeStickerFont != null) IllustrativeStickerFont = string.IsNullOrEmpty(imported.IllustrativeStickerFont) ? "Theme Font" : imported.IllustrativeStickerFont;
                            IllustrativeStickerFontBold = imported.IllustrativeStickerFontBold;
                            IllustrativeStickerFontItalic = imported.IllustrativeStickerFontItalic;

                            System.Windows.MessageBox.Show("Configuration imported successfully! Please click 'Save & Apply' to confirm changes.", "Import Complete", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                            OnPropertyChanged(string.Empty);
                        }
                    } catch (Exception ex) {
                        System.Windows.MessageBox.Show($"Error importing configuration: {ex.Message}");
                    }
                }
            });
        }

        public void RemoveFeatureFromSection(SidebarSection section, SidebarFeature feature)
        {
            if (section != null && feature != null)
            {
                var target = section.Features.FirstOrDefault(f => f.Id == feature.Id);
                if (target != null)
                {
                    section.Features.Remove(target);
                    RefreshSections();
                }
            }
        }

        private void RefreshSections()
        {
            OnPropertyChanged(nameof(Sections));
        }

        public FeatureDisplayInfo GetInfo(string id) => FeatureLibrary.GetFeatureInfo(id);

        private void LoadConfig()
        {
            Config = JsonService.Load<SidebarConfig>("sidebar_config.json");
            if (Config == null)
            {
                Config = FeatureLibrary.GetEmergencyDefault();
            }
            else
            {
                // If background is the old hardcoded default, change to Transparent to allow theme tracking
                if (Config.BackgroundColor == "#252525")
                {
                    Config.BackgroundColor = "Transparent";
                }

                // Ensure legacy features are mapped correctly in the Settings window too
                bool changed = false;
                foreach (var section in Config.Sections)
                {
                    if (section.Features == null) continue;
                    for (int i = 0; i < section.Features.Count; i++)
                    {
                        var f = section.Features[i];
                        if (f != null)
                        {
                            if (f.Id == "BtnAlignToTable")
                            {
                                f.Id = "BtnAlignToTableCell";
                                f.Name = "Align to Table Cell";
                                changed = true;
                            }
                            if (f.Id == "BtnReplaceFont" && f.Name != "Bulk Format Replacer")
                            {
                                f.Name = "Bulk Format Replacer";
                                changed = true;
                            }
                        }
                    }
                }
                if (changed) JsonService.Save(Config, "sidebar_config.json");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
