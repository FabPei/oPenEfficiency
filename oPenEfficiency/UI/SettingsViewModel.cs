using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using oPenEfficiency;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using System.Collections.Generic;

namespace oPenEfficiency.UI
{
    public class ShortcutItemViewModel : INotifyPropertyChanged
    {
        public string FeatureId { get; set; }
        public string FeatureName { get; set; }
        public string AttentionText { get; set; }
        public System.Windows.Visibility AttentionVisibility { get; set; }
        public bool HasAssignment { get; set; }
        
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
            set
            {
                _config = value;
                ReloadStyleCheckTokens();
                OnPropertyChanged();
                OnPropertyChanged(nameof(BackgroundColor));
                OnPropertyChanged(nameof(UseThemeBackground));
                OnPropertyChanged(nameof(UseCustomBackground));
                OnPropertyChanged(nameof(SectionFontColor));
                OnPropertyChanged(nameof(UseThemeSectionFontColor));
                OnPropertyChanged(nameof(UseCustomSectionFontColor));
                OnPropertyChanged(nameof(SectionFontSize));
                OnPropertyChanged(nameof(ButtonBackgroundColor));
                OnPropertyChanged(nameof(StickyNoteStyle));
                OnPropertyChanged(nameof(StickyNoteDefaultText));
                OnPropertyChanged(nameof(StickyNoteFont));
                OnPropertyChanged(nameof(StickyNoteBackgroundColor));
                OnPropertyChanged(nameof(StickyNoteFontColor));
                OnPropertyChanged(nameof(SpacingIncrement));
                OnPropertyChanged(nameof(IconSize));
                OnPropertyChanged(nameof(IconColorMode));
                OnPropertyChanged(nameof(AppTheme));
                OnPropertyChanged(nameof(AppFontFamily));
                OnPropertyChanged(nameof(DeepLApiKey));
                OnPropertyChanged(nameof(TranslationTargetLanguage));
                NotifyStickerPreviewChanged();
            }
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
                    NotifyStickerPreviewChanged();
                }
            }
        }

        public string IconColorMode
        {
            get => IconColorModes.Normalize(_config?.IconColorMode);
            set
            {
                string normalized = IconColorModes.Normalize(value);
                if (_config != null && !string.Equals(_config.IconColorMode, normalized, StringComparison.Ordinal))
                {
                    _config.IconColorMode = normalized;
                    OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<string> AvailableIconColorModes { get; } =
            new ObservableCollection<string>
            {
                IconColorModes.PowerPoint,
                IconColorModes.FeatureColors
            };

        public string BackgroundColor
        {
            get => _config?.BackgroundColor ?? "Transparent";
            set
            {
                if (_config != null && _config.BackgroundColor != value)
                {
                    _config.BackgroundColor = value;
                    _config.BackgroundMode = SidebarBackgroundModes.Custom;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(UseThemeBackground));
                    OnPropertyChanged(nameof(UseCustomBackground));
                }
            }
        }

        public bool UseThemeBackground
        {
            get => SidebarBackgroundModes.UsesTheme(_config);
            set
            {
                if (_config == null || value == SidebarBackgroundModes.UsesTheme(_config)) return;
                _config.BackgroundMode = value
                    ? SidebarBackgroundModes.Theme
                    : SidebarBackgroundModes.Custom;
                if (value) _config.BackgroundColor = "Transparent";
                OnPropertyChanged();
                OnPropertyChanged(nameof(BackgroundColor));
                OnPropertyChanged(nameof(UseCustomBackground));
            }
        }

        public bool UseCustomBackground
        {
            get => !UseThemeBackground;
            set
            {
                if (value) UseThemeBackground = false;
                else if (!value && !UseThemeBackground) UseThemeBackground = true;
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
                    _config.SectionFontColorMode = SidebarSectionFontColorModes.Custom;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(UseThemeSectionFontColor));
                    OnPropertyChanged(nameof(UseCustomSectionFontColor));
                }
            }
        }

        public bool UseThemeSectionFontColor
        {
            get => SidebarSectionFontColorModes.UsesTheme(_config);
            set
            {
                if (_config == null || value == SidebarSectionFontColorModes.UsesTheme(_config)) return;
                _config.SectionFontColorMode = value
                    ? SidebarSectionFontColorModes.Theme
                    : SidebarSectionFontColorModes.Custom;
                OnPropertyChanged();
                OnPropertyChanged(nameof(UseCustomSectionFontColor));
            }
        }

        public bool UseCustomSectionFontColor
        {
            get => !UseThemeSectionFontColor;
            set
            {
                if (value) UseThemeSectionFontColor = false;
                else if (!value && !UseThemeSectionFontColor) UseThemeSectionFontColor = true;
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
                    StickerPreviewModel.SetStickyNoteColors(
                        _config,
                        value,
                        GetRawStickyNoteBackgroundColor(),
                        GetRawStickyNoteFontColor());
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(StickyNoteBackgroundColor));
                    OnPropertyChanged(nameof(StickyNoteFontColor));
                    NotifyStickerPreviewChanged();
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
                    NotifyStickerPreviewChanged();
                }
            }
        }

        public string MasterObjectMode
        {
            get => _config?.MasterObjectMode ?? "First selected";
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
                    if (!_syncingStyleCheckTokens) ReloadStyleCheckTokens();
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
                    if (!_syncingStyleCheckTokens) ReloadStyleCheckTokens();
                }
            }
        }

        private bool _syncingStyleCheckTokens;
        private string _newStyleCheckColor;
        private string _newStyleCheckFont;
        private string _styleCheckColorError;
        private string _styleCheckFontError;

        public ObservableCollection<string> StyleCheckColorTokens { get; } = new ObservableCollection<string>();
        public ObservableCollection<string> StyleCheckFontTokens { get; } = new ObservableCollection<string>();

        public string NewStyleCheckColor
        {
            get => _newStyleCheckColor;
            set { _newStyleCheckColor = value; StyleCheckColorError = null; OnPropertyChanged(); }
        }

        public string NewStyleCheckFont
        {
            get => _newStyleCheckFont;
            set { _newStyleCheckFont = value; StyleCheckFontError = null; OnPropertyChanged(); }
        }

        public string StyleCheckColorError
        {
            get => _styleCheckColorError;
            private set { _styleCheckColorError = value; OnPropertyChanged(); }
        }

        public string StyleCheckFontError
        {
            get => _styleCheckFontError;
            private set { _styleCheckFontError = value; OnPropertyChanged(); }
        }

        private void ReloadStyleCheckTokens()
        {
            if (_syncingStyleCheckTokens) return;
            ReplaceTokens(StyleCheckColorTokens, _config?.StyleCheckCorrectColors);
            ReplaceTokens(StyleCheckFontTokens, _config?.StyleCheckCorrectFonts);
        }

        private static void ReplaceTokens(ObservableCollection<string> target, string serialized)
        {
            target.Clear();
            foreach (string token in DelimitedSettingTokens.Parse(serialized))
            {
                target.Add(token);
            }
        }

        private void CommitStyleCheckTokens()
        {
            _syncingStyleCheckTokens = true;
            try
            {
                StyleCheckCorrectColors = DelimitedSettingTokens.Serialize(StyleCheckColorTokens);
                StyleCheckCorrectFonts = DelimitedSettingTokens.Serialize(StyleCheckFontTokens);
            }
            finally
            {
                _syncingStyleCheckTokens = false;
            }
        }

        private void AddStyleCheckColor()
        {
            if (!DelimitedSettingTokens.TryNormalizeHexColor(NewStyleCheckColor, out string value))
            {
                StyleCheckColorError = "Enter a color as #RRGGBB or #AARRGGBB.";
                return;
            }
            if (StyleCheckColorTokens.Any(item => string.Equals(item, value, StringComparison.OrdinalIgnoreCase)))
            {
                StyleCheckColorError = "This color is already approved.";
                return;
            }
            StyleCheckColorTokens.Add(value);
            NewStyleCheckColor = string.Empty;
            CommitStyleCheckTokens();
        }

        private void AddStyleCheckFont()
        {
            if (!DelimitedSettingTokens.TryNormalizeFontName(NewStyleCheckFont, out string value))
            {
                StyleCheckFontError = "Enter one font name without commas.";
                return;
            }
            if (StyleCheckFontTokens.Any(item => string.Equals(item, value, StringComparison.OrdinalIgnoreCase)))
            {
                StyleCheckFontError = "This font is already approved.";
                return;
            }
            StyleCheckFontTokens.Add(value);
            NewStyleCheckFont = string.Empty;
            CommitStyleCheckTokens();
        }

        private void RemoveStyleCheckColor(object parameter)
        {
            if (parameter is string value && StyleCheckColorTokens.Remove(value)) CommitStyleCheckTokens();
        }

        private void RemoveStyleCheckFont(object parameter)
        {
            if (parameter is string value && StyleCheckFontTokens.Remove(value)) CommitStyleCheckTokens();
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
                    NotifyStickerPreviewChanged();
                }
            }
        }

        public string StickyNoteBackgroundColor
        {
            get => GetRawStickyNoteBackgroundColor();
            set
            {
                if (_config != null)
                {
                    if (string.Equals(GetRawStickyNoteBackgroundColor(), value, StringComparison.Ordinal)) return;
                    StickerPreviewModel.SetStickyNoteColors(
                        _config,
                        StickyNoteStyle,
                        value,
                        GetRawStickyNoteFontColor());
                    OnPropertyChanged();
                    NotifyStickerPreviewChanged();
                }
            }
        }

        public string StickyNoteFontColor
        {
            get => GetRawStickyNoteFontColor();
            set
            {
                if (_config != null)
                {
                    if (string.Equals(GetRawStickyNoteFontColor(), value, StringComparison.Ordinal)) return;
                    StickerPreviewModel.SetStickyNoteColors(
                        _config,
                        StickyNoteStyle,
                        GetRawStickyNoteBackgroundColor(),
                        value);
                    OnPropertyChanged();
                    NotifyStickerPreviewChanged();
                }
            }
        }

        private string GetRawStickyNoteBackgroundColor()
        {
            string value;
            switch (StickyNoteStyle)
            {
                case "Consulting": value = _config?.StickyNoteConsultingBackgroundColor; break;
                case "GreenB": value = _config?.StickyNoteGreenBBackgroundColor; break;
                default: value = _config?.StickyNoteThemeBackgroundColor; break;
            }
            return value ?? StickerPreviewModel.CreateStickyNote(_config, AppFontFamily).BackgroundColor;
        }

        private string GetRawStickyNoteFontColor()
        {
            string value;
            switch (StickyNoteStyle)
            {
                case "Consulting": value = _config?.StickyNoteConsultingFontColor; break;
                case "GreenB": value = _config?.StickyNoteGreenBFontColor; break;
                default: value = _config?.StickyNoteThemeFontColor; break;
            }
            return value ?? StickerPreviewModel.CreateStickyNote(_config, AppFontFamily).FontColor;
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
                    NotifyStickerPreviewChanged();
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
                    NotifyStickerPreviewChanged();
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
                    NotifyStickerPreviewChanged();
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
                    NotifyStickerPreviewChanged();
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
                    NotifyStickerPreviewChanged();
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
                    NotifyStickerPreviewChanged();
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
                    NotifyStickerPreviewChanged();
                }
            }
        }

        public string IllustrativeStickerPosition
        {
            get => _config?.IllustrativeStickerPosition ?? "TopRight";
            set
            {
                if (_config != null && _config.IllustrativeStickerPosition != value)
                {
                    _config.IllustrativeStickerPosition = value;
                    OnPropertyChanged();
                    NotifyStickerPreviewChanged();
                }
            }
        }

        public bool IllustrativeStickerUseThemeColor
        {
            get => _config?.IllustrativeStickerUseThemeColor ?? false;
            set
            {
                if (_config != null && _config.IllustrativeStickerUseThemeColor != value)
                {
                    _config.IllustrativeStickerUseThemeColor = value;
                    OnPropertyChanged();
                    NotifyStickerPreviewChanged();
                }
            }
        }

        public bool IllustrativeStickerIncludeDate
        {
            get => _config?.IllustrativeStickerIncludeDate ?? false;
            set
            {
                if (_config != null && _config.IllustrativeStickerIncludeDate != value)
                {
                    _config.IllustrativeStickerIncludeDate = value;
                    OnPropertyChanged();
                    NotifyStickerPreviewChanged();
                }
            }
        }

        public bool IllustrativeStickerIncludeAuthor
        {
            get => _config?.IllustrativeStickerIncludeAuthor ?? false;
            set
            {
                if (_config != null && _config.IllustrativeStickerIncludeAuthor != value)
                {
                    _config.IllustrativeStickerIncludeAuthor = value;
                    OnPropertyChanged();
                    NotifyStickerPreviewChanged();
                }
            }
        }

        public string StickyNotePreviewStyle => StickerPreviewModel.CreateStickyNote(_config, AppFontFamily).Style;
        public string StickyNotePreviewText => StickerPreviewModel.CreateStickyNote(_config, AppFontFamily).Text;
        public string StickyNotePreviewFontFamily => StickerPreviewModel.CreateStickyNote(_config, AppFontFamily).FontFamily;
        public string StickyNotePreviewBackgroundColor => StickerPreviewModel.CreateStickyNote(_config, AppFontFamily).BackgroundColor;
        public string StickyNotePreviewFontColor => StickerPreviewModel.CreateStickyNote(_config, AppFontFamily).FontColor;

        public string StatusStickerPreviewText => StickerPreviewModel.CreateStatusSticker(_config, AppFontFamily).Text;
        public string StatusStickerPreviewFontFamily => StickerPreviewModel.CreateStatusSticker(_config, AppFontFamily).FontFamily;
        public string StatusStickerPreviewColor => StickerPreviewModel.CreateStatusSticker(_config, AppFontFamily).Color;
        public double StatusStickerPreviewFontSize => Math.Max(8d, Math.Min(32d, StickerPreviewModel.CreateStatusSticker(_config, AppFontFamily).FontSize));
        public double StatusStickerPreviewLineSize => Math.Max(1d, Math.Min(8d, StickerPreviewModel.CreateStatusSticker(_config, AppFontFamily).LineSize));
        public bool StatusStickerPreviewBold => StickerPreviewModel.CreateStatusSticker(_config, AppFontFamily).Bold;
        public bool StatusStickerPreviewItalic => StickerPreviewModel.CreateStatusSticker(_config, AppFontFamily).Italic;
        public bool StatusStickerPreviewUsesThemeColor => StickerPreviewModel.CreateStatusSticker(_config, AppFontFamily).UseThemeColor;
        public string StatusStickerPreviewPositionLabel => StickerPreviewModel.CreateStatusSticker(_config, AppFontFamily).PositionLabel;

        private void NotifyStickerPreviewChanged()
        {
            OnPropertyChanged(nameof(StickyNotePreviewStyle));
            OnPropertyChanged(nameof(StickyNotePreviewText));
            OnPropertyChanged(nameof(StickyNotePreviewFontFamily));
            OnPropertyChanged(nameof(StickyNotePreviewBackgroundColor));
            OnPropertyChanged(nameof(StickyNotePreviewFontColor));
            OnPropertyChanged(nameof(StatusStickerPreviewText));
            OnPropertyChanged(nameof(StatusStickerPreviewFontFamily));
            OnPropertyChanged(nameof(StatusStickerPreviewColor));
            OnPropertyChanged(nameof(StatusStickerPreviewFontSize));
            OnPropertyChanged(nameof(StatusStickerPreviewLineSize));
            OnPropertyChanged(nameof(StatusStickerPreviewBold));
            OnPropertyChanged(nameof(StatusStickerPreviewItalic));
            OnPropertyChanged(nameof(StatusStickerPreviewUsesThemeColor));
            OnPropertyChanged(nameof(StatusStickerPreviewPositionLabel));
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
            "First selected",
            "Last selected"
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
                var matches = new HashSet<string>(
                    FeatureLibrary.Catalog
                        .Where(record => record.UserVisible &&
                            !record.IsObsolete &&
                            FeatureCatalogSearch.Matches(record, SearchText))
                        .Select(record => record.Id),
                    StringComparer.Ordinal);
                return new ObservableCollection<SidebarFeature>(
                    pool.Where(feature => matches.Contains(feature.Id)).OrderBy(feature => feature.Name));
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
            var assignments = ShortcutManager.GetAssignments()
                .ToDictionary(assignment => assignment.FeatureId, StringComparer.Ordinal);

            var shortcutEligibleIds = new HashSet<string>(
                FeatureLibrary.Catalog
                    .Where(record => record.SupportsUserShortcut)
                    .Select(record => record.Id),
                StringComparer.Ordinal);
            foreach (var feature in FeatureLibrary.AllFeatures.Where(feature =>
                shortcutEligibleIds.Contains(feature.Id) ||
                assignments.ContainsKey(feature.Id)))
            {
                assignments.TryGetValue(feature.Id, out var assignment);
                string text = assignment?.ShortcutText ?? "Not assigned";
                
                var item = new ShortcutItemViewModel
                {
                    FeatureId = feature.Id,
                    FeatureName = feature.Name,
                    ShortcutText = text,
                    HasAssignment = assignment != null,
                    AttentionText = assignment?.NeedsAttention == true
                        ? "Requires attention: " + ResolveShortcutFeatureNames(assignment.AttentionReason)
                        : null,
                    AttentionVisibility = assignment?.NeedsAttention == true
                        ? System.Windows.Visibility.Visible
                        : System.Windows.Visibility.Collapsed,
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

        private static string ResolveShortcutFeatureNames(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;
            foreach (var feature in FeatureLibrary.AllFeatures)
                text = text.Replace(feature.Id, feature.Name);
            return text;
        }

        private void RefreshFilteredShortcuts()
        {
            FilteredShortcuts.Clear();
            if (_fullShortcutsList == null) return;

            foreach (var item in _fullShortcutsList)
            {
                if (string.IsNullOrWhiteSpace(ShortcutSearchText) || 
                    FeatureCatalogSearch.Matches(
                        FeatureLibrary.Catalog.FirstOrDefault(record =>
                            string.Equals(record.Id, item.FeatureId, StringComparison.Ordinal)),
                        ShortcutSearchText))
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
        public ICommand AddStyleCheckColorCommand { get; }
        public ICommand RemoveStyleCheckColorCommand { get; }
        public ICommand AddStyleCheckFontCommand { get; }
        public ICommand RemoveStyleCheckFontCommand { get; }
        
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

        private string _selectedPreset = SidebarPresetFactory.ConsultingStandardName;
        public string SelectedPreset
        {
            get => _selectedPreset;
            set
            {
                _selectedPreset = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(SelectedPresetDescription));
            }
        }

        public string SelectedPresetDescription =>
            SidebarPresetFactory.GetDescription(SelectedPreset);

        public ObservableCollection<string> AvailablePresets { get; } =
            new ObservableCollection<string>(SidebarPresetFactory.PresetNames);

        public ICommand ApplyPresetCommand { get; private set; }
        public ICommand ExportLayoutOnlyCommand { get; private set; }
        public ICommand ImportLayoutOnlyCommand { get; private set; }

        public SettingsViewModel()
        {
            LoadConfig();
            InitializeAvailableThemes();
            LoadShortcutsList();

            AddStyleCheckColorCommand = new RelayCommand(_ => AddStyleCheckColor());
            RemoveStyleCheckColorCommand = new RelayCommand(RemoveStyleCheckColor);
            AddStyleCheckFontCommand = new RelayCommand(_ => AddStyleCheckFont());
            RemoveStyleCheckFontCommand = new RelayCommand(RemoveStyleCheckFont);

            ApplyPresetCommand = new RelayCommand(_ => {
                if (Config == null) return;
                var presetConfig = SidebarPresetFactory.Create(
                    SelectedPreset,
                    FeatureLibrary.Catalog);

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
                    try {
                        var layoutOnly = new SidebarConfig { Sections = Config.Sections };
                        string json = JsonService.Serialize(layoutOnly);
                        System.IO.File.WriteAllText(sfd.FileName, json);
                        System.Windows.MessageBox.Show("Layout exported successfully.", "Export Complete", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                    } catch (Exception ex) {
                        System.Windows.MessageBox.Show("The layout could not be exported: " + ex.Message, "Export Failed", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                    }
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
                            Config.Sections = SanitizeSections(imported.Sections, ref missingCount);
                            RefreshSections();
                            
                            if (missingCount > 0) {
                                System.Windows.MessageBox.Show($"Layout imported successfully, but {missingCount} feature(s) from the import file were not recognized by this version of openEfficiency and have been skipped.", "Import Warning", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                            }
                        }
                        else {
                            System.Windows.MessageBox.Show("The selected file is not a valid openEfficiency layout.", "Import Failed", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                        }
                    } catch (Exception ex) {
                        System.Windows.MessageBox.Show("Error importing layout: " + ex.Message);
                    }
                }
            });

            ResetStickyStylesCommand = new RelayCommand(_ => {
                string background = StickerPreviewModel.GetDefaultStickyColor(StickyNoteStyle, true);
                string foreground = StickerPreviewModel.GetDefaultStickyColor(StickyNoteStyle, false);
                StickerPreviewModel.SetStickyNoteColors(Config, StickyNoteStyle, background, foreground);
                OnPropertyChanged(nameof(StickyNoteBackgroundColor));
                OnPropertyChanged(nameof(StickyNoteFontColor));
                NotifyStickerPreviewChanged();
            });

            ResetSidebarStylesCommand = new RelayCommand(_ => {
                BackgroundColor = "Transparent";
                UseThemeBackground = true;
                SectionFontColor = "#9CA3AF";
                UseThemeSectionFontColor = true;
                ButtonBackgroundColor = "Transparent";
                SectionFontSize = 9;
                IconSize = 16;
                IconColorMode = IconColorModes.PowerPoint;
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
                Config.IconColorMode = IconColorModes.Normalize(Config.IconColorMode);
                SidebarBackgroundModes.Normalize(Config);
                SidebarSectionFontColorModes.Normalize(Config);
                if (!JsonService.TrySave(Config, "sidebar_config.json")) {
                    System.Windows.MessageBox.Show("Settings could not be saved. Your existing configuration was left unchanged.", "Save Failed", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                    return;
                }
                ThemeManager.ApplyFontFamily(Config?.AppFontFamily);
                RebuildAction?.Invoke();
                CloseAction?.Invoke();
            });

            CloseCommand = new RelayCommand(_ => CloseAction?.Invoke());

            ResetToDefaultsCommand = new RelayCommand(_ => {
                var result = System.Windows.MessageBox.Show(
                    "Are you sure you want to restore the sidebar layout and appearance settings to their defaults? Keyboard shortcuts are managed separately and will not be changed.",
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
                        if (!ExcludeApiKeys)
                        {
                            var confirmation = System.Windows.MessageBox.Show(
                                "This portable export will contain your API keys in plain text. Store and share the file securely. Continue?",
                                "Export API Keys",
                                System.Windows.MessageBoxButton.YesNo,
                                System.Windows.MessageBoxImage.Warning);
                            if (confirmation != System.Windows.MessageBoxResult.Yes) return;
                        }

                        var exportData = new SidebarConfig();
                        
                        // Default the generic properties we're not exporting
                        exportData.Sections = null;
                        exportData.BackgroundColor = null;
                        exportData.BackgroundMode = null;
                        exportData.SectionFontColor = null;
                        exportData.SectionFontColorMode = null;
                        exportData.ButtonBackgroundColor = null;
                        exportData.SectionFontSize = 0;
                        exportData.IconSize = 0;
                        exportData.IconColorMode = null;
                        exportData.StickyNoteStyle = null;
                        exportData.StickyNoteDefaultText = null;
                        exportData.StickyNoteFont = null;
                        exportData.StickyNoteBackgroundColor = null;
                        exportData.StickyNoteFontColor = null;
                        exportData.StickyNoteThemeBackgroundColor = null;
                        exportData.StickyNoteThemeFontColor = null;
                        exportData.StickyNoteConsultingBackgroundColor = null;
                        exportData.StickyNoteConsultingFontColor = null;
                        exportData.StickyNoteGreenBBackgroundColor = null;
                        exportData.StickyNoteGreenBFontColor = null;

                        if (ExportLayout) {
                            var json = JsonService.Serialize(Config.Sections);
                            exportData.Sections = JsonService.Deserialize<List<SidebarSection>>(json);
                        }
                        if (ExportColors) {
                            exportData.BackgroundColor = Config.BackgroundColor;
                            exportData.BackgroundMode = Config.BackgroundMode;
                            exportData.SectionFontColor = Config.SectionFontColor;
                            exportData.SectionFontColorMode = Config.SectionFontColorMode;
                            exportData.ButtonBackgroundColor = Config.ButtonBackgroundColor;
                            exportData.SectionFontSize = Config.SectionFontSize;
                            exportData.IconSize = Config.IconSize;
                            exportData.IconColorMode = IconColorModes.Normalize(Config.IconColorMode);
                            exportData.SpacingIncrement = Config.SpacingIncrement;
                            exportData.AppTheme = Config.AppTheme;
                            exportData.AppFontFamily = Config.AppFontFamily;
                        }
                        if (ExportSticky) {
                            exportData.StickyNoteStyle = Config.StickyNoteStyle;
                            exportData.StickyNoteDefaultText = Config.StickyNoteDefaultText;
                            exportData.StickyNoteFont = Config.StickyNoteFont;
                            exportData.StickyNoteBackgroundColor = Config.StickyNoteBackgroundColor;
                            exportData.StickyNoteFontColor = Config.StickyNoteFontColor;
                            exportData.StickyNoteThemeBackgroundColor = Config.StickyNoteThemeBackgroundColor;
                            exportData.StickyNoteThemeFontColor = Config.StickyNoteThemeFontColor;
                            exportData.StickyNoteConsultingBackgroundColor = Config.StickyNoteConsultingBackgroundColor;
                            exportData.StickyNoteConsultingFontColor = Config.StickyNoteConsultingFontColor;
                            exportData.StickyNoteGreenBBackgroundColor = Config.StickyNoteGreenBBackgroundColor;
                            exportData.StickyNoteGreenBFontColor = Config.StickyNoteGreenBFontColor;
                            exportData.IllustrativeStickerTexts = Config.IllustrativeStickerTexts;
                            exportData.IllustrativeStickerColor = Config.IllustrativeStickerColor;
                            exportData.IllustrativeStickerLineSize = Config.IllustrativeStickerLineSize;
                            exportData.IllustrativeStickerFontSize = Config.IllustrativeStickerFontSize;
                            exportData.IllustrativeStickerFont = Config.IllustrativeStickerFont;
                            exportData.IllustrativeStickerFontBold = Config.IllustrativeStickerFontBold;
                            exportData.IllustrativeStickerFontItalic = Config.IllustrativeStickerFontItalic;
                            exportData.IllustrativeStickerPosition = Config.IllustrativeStickerPosition;
                            exportData.IllustrativeStickerUseThemeColor = Config.IllustrativeStickerUseThemeColor;
                            exportData.IllustrativeStickerIncludeDate = Config.IllustrativeStickerIncludeDate;
                            exportData.IllustrativeStickerIncludeAuthor = Config.IllustrativeStickerIncludeAuthor;
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
                            int missingFeatureCount = 0;
                            if (imported.Sections != null && imported.Sections.Count > 0) {
                                Config.Sections = SanitizeSections(imported.Sections, ref missingFeatureCount);
                                RefreshSections();
                            }
                            if (!string.IsNullOrEmpty(imported.BackgroundColor))
                            {
                                Config.BackgroundColor = imported.BackgroundColor;
                                Config.BackgroundMode = imported.BackgroundMode;
                                SidebarBackgroundModes.Normalize(Config);
                                OnPropertyChanged(nameof(BackgroundColor));
                                OnPropertyChanged(nameof(UseThemeBackground));
                                OnPropertyChanged(nameof(UseCustomBackground));
                            }
                            if (!string.IsNullOrEmpty(imported.SectionFontColor))
                            {
                                Config.SectionFontColor = imported.SectionFontColor;
                                Config.SectionFontColorMode = imported.SectionFontColorMode;
                                SidebarSectionFontColorModes.Normalize(Config);
                                OnPropertyChanged(nameof(SectionFontColor));
                                OnPropertyChanged(nameof(UseThemeSectionFontColor));
                                OnPropertyChanged(nameof(UseCustomSectionFontColor));
                            }
                            if (!string.IsNullOrEmpty(imported.ButtonBackgroundColor)) ButtonBackgroundColor = imported.ButtonBackgroundColor;
                            if (imported.SectionFontSize > 0) SectionFontSize = imported.SectionFontSize;
                            if (imported.IconSize > 0) Config.IconSize = imported.IconSize;
                            if (imported.IconColorMode != null) IconColorMode = imported.IconColorMode;
                            if (imported.SpacingIncrement > 0) SpacingIncrement = imported.SpacingIncrement;
                            if (imported.AppTheme != null) AppTheme = imported.AppTheme;
                            if (imported.AppFontFamily != null) AppFontFamily = imported.AppFontFamily;
                            InitializeAvailableThemes(); // Ensure theme list reflects imported config
                            if (imported.StickyNoteStyle != null) StickyNoteStyle = imported.StickyNoteStyle;
                            if (imported.StickyNoteDefaultText != null) StickyNoteDefaultText = imported.StickyNoteDefaultText;
                            if (imported.StickyNoteFont != null) StickyNoteFont = string.IsNullOrEmpty(imported.StickyNoteFont) ? "Theme Font" : imported.StickyNoteFont;
                            bool hasStyleColors = imported.StickyNoteThemeBackgroundColor != null ||
                                imported.StickyNoteThemeFontColor != null ||
                                imported.StickyNoteConsultingBackgroundColor != null ||
                                imported.StickyNoteConsultingFontColor != null ||
                                imported.StickyNoteGreenBBackgroundColor != null ||
                                imported.StickyNoteGreenBFontColor != null;
                            if (hasStyleColors)
                            {
                                Config.StickyNoteThemeBackgroundColor = imported.StickyNoteThemeBackgroundColor;
                                Config.StickyNoteThemeFontColor = imported.StickyNoteThemeFontColor;
                                Config.StickyNoteConsultingBackgroundColor = imported.StickyNoteConsultingBackgroundColor;
                                Config.StickyNoteConsultingFontColor = imported.StickyNoteConsultingFontColor;
                                Config.StickyNoteGreenBBackgroundColor = imported.StickyNoteGreenBBackgroundColor;
                                Config.StickyNoteGreenBFontColor = imported.StickyNoteGreenBFontColor;
                                Config.StickyNoteBackgroundColor = imported.StickyNoteBackgroundColor;
                                Config.StickyNoteFontColor = imported.StickyNoteFontColor;
                                OnPropertyChanged(nameof(StickyNoteBackgroundColor));
                                OnPropertyChanged(nameof(StickyNoteFontColor));
                                NotifyStickerPreviewChanged();
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(imported.StickyNoteBackgroundColor)) StickyNoteBackgroundColor = imported.StickyNoteBackgroundColor;
                                if (!string.IsNullOrEmpty(imported.StickyNoteFontColor)) StickyNoteFontColor = imported.StickyNoteFontColor;
                            }
                            if (imported.IllustrativeStickerTexts != null) IllustrativeStickerTexts = imported.IllustrativeStickerTexts;
                            if (imported.IllustrativeStickerColor != null) IllustrativeStickerColor = imported.IllustrativeStickerColor;
                            if (imported.IllustrativeStickerLineSize > 0) IllustrativeStickerLineSize = imported.IllustrativeStickerLineSize;
                            if (imported.IllustrativeStickerFontSize > 0) IllustrativeStickerFontSize = imported.IllustrativeStickerFontSize;
                            if (imported.IllustrativeStickerFont != null) IllustrativeStickerFont = string.IsNullOrEmpty(imported.IllustrativeStickerFont) ? "Theme Font" : imported.IllustrativeStickerFont;
                            IllustrativeStickerFontBold = imported.IllustrativeStickerFontBold;
                            IllustrativeStickerFontItalic = imported.IllustrativeStickerFontItalic;
                            if (!string.IsNullOrWhiteSpace(imported.IllustrativeStickerPosition)) IllustrativeStickerPosition = imported.IllustrativeStickerPosition;
                            IllustrativeStickerUseThemeColor = imported.IllustrativeStickerUseThemeColor;
                            IllustrativeStickerIncludeDate = imported.IllustrativeStickerIncludeDate;
                            IllustrativeStickerIncludeAuthor = imported.IllustrativeStickerIncludeAuthor;

                            if (!string.IsNullOrEmpty(imported.TranslationTargetLanguage)) TranslationTargetLanguage = imported.TranslationTargetLanguage;
                            if (!string.IsNullOrEmpty(imported.SelectedTranslationTool)) SelectedTranslationTool = imported.SelectedTranslationTool;
                            if (!string.IsNullOrEmpty(imported.OpenAIModel)) OpenAIModel = imported.OpenAIModel;
                            if (!string.IsNullOrEmpty(imported.ClaudeModel)) ClaudeModel = imported.ClaudeModel;
                            if (!string.IsNullOrEmpty(imported.LibreTranslateBaseUrl)) LibreTranslateBaseUrl = imported.LibreTranslateBaseUrl;
                            if (!string.IsNullOrEmpty(imported.DeepLApiKey)) DeepLApiKey = imported.DeepLApiKey;
                            if (!string.IsNullOrEmpty(imported.OpenAIApiKey)) OpenAIApiKey = imported.OpenAIApiKey;
                            if (!string.IsNullOrEmpty(imported.ClaudeApiKey)) ClaudeApiKey = imported.ClaudeApiKey;
                            if (!string.IsNullOrEmpty(imported.LibreTranslateApiKey)) LibreTranslateApiKey = imported.LibreTranslateApiKey;

                            string importMessage = "Configuration imported. Click 'Save & Apply' to persist the changes.";
                            if (missingFeatureCount > 0) importMessage += $"\n\n{missingFeatureCount} unknown feature(s) were skipped.";
                            System.Windows.MessageBox.Show(importMessage, "Import Complete", System.Windows.MessageBoxButton.OK,
                                missingFeatureCount > 0 ? System.Windows.MessageBoxImage.Warning : System.Windows.MessageBoxImage.Information);
                            OnPropertyChanged(string.Empty);
                        }
                        else {
                            System.Windows.MessageBox.Show("The selected file is not a valid openEfficiency configuration.", "Import Failed", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                        }
                    } catch (Exception ex) {
                        System.Windows.MessageBox.Show($"Error importing configuration: {ex.Message}");
                    }
                }
            });
        }

        private static List<SidebarSection> SanitizeSections(IEnumerable<SidebarSection> sections, ref int missingCount)
        {
            var result = new List<SidebarSection>();
            if (sections == null) return result;
            foreach (var section in sections.Where(section => section != null))
            {
                section.Features = section.Features ?? new List<SidebarFeature>();
                var valid = new List<SidebarFeature>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var feature in section.Features.Where(feature => feature != null && !string.IsNullOrWhiteSpace(feature.Id)))
                {
                    var known = FeatureLibrary.AllFeatures.FirstOrDefault(item => string.Equals(item.Id, feature.Id, StringComparison.OrdinalIgnoreCase));
                    if (known == null) { missingCount++; continue; }
                    if (!seen.Add(known.Id)) continue;
                    valid.Add(new SidebarFeature { Id = known.Id, Name = known.Name });
                }
                section.Features = valid;
                result.Add(section);
            }
            return result;
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
            var savedConfig = JsonService.Load<SidebarConfig>("sidebar_config.json");
            Config = SidebarPresetFactory.ResolveInitialLayout(
                savedConfig,
                FeatureLibrary.Catalog);
            if (savedConfig == null)
            {
                return;
            }

            bool changed = StickerPreviewModel.NormalizeStickyNoteColors(Config);
            changed |= SidebarBackgroundModes.Normalize(Config);
            changed |= SidebarSectionFontColorModes.Normalize(Config);
            changed |= SidebarLayoutMigration.Migrate(
                Config,
                FeatureLibrary.AllFeatures).Changed;
            string canonicalIconColorMode = IconColorModes.Normalize(Config.IconColorMode);
            if (!string.Equals(Config.IconColorMode, canonicalIconColorMode, StringComparison.Ordinal))
            {
                Config.IconColorMode = canonicalIconColorMode;
                changed = true;
            }
            string canonicalAnchorMode = AnchorModes.IsLastSelected(Config.MasterObjectMode)
                ? "Last selected"
                : "First selected";
            if (!string.Equals(Config.MasterObjectMode, canonicalAnchorMode, StringComparison.Ordinal))
            {
                Config.MasterObjectMode = canonicalAnchorMode;
                changed = true;
            }
            if (changed) JsonService.Save(Config, "sidebar_config.json");
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
