[CmdletBinding()]
param(
    [string]$RepositoryRoot,

    [string]$BaselinePath,

    [switch]$Audit,

    [switch]$SkipSelfTest
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) {
    $RepositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
}
else {
    $RepositoryRoot = [IO.Path]::GetFullPath($RepositoryRoot)
}
if ([string]::IsNullOrWhiteSpace($BaselinePath)) {
    $BaselinePath = Join-Path $PSScriptRoot "visual-consistency-baseline.json"
}

function ConvertTo-NormalSignature([string]$Value) {
    if ($null -eq $Value) { return "" }
    return [regex]::Replace($Value.Trim(), '\s+', ' ')
}

function Get-Fingerprint([string]$Rule, [string]$Path, [string]$Signature) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $identity = "$Rule`n$Path`n$Signature"
        $bytes = [Text.Encoding]::UTF8.GetBytes($identity)
        return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-LineNumber([string]$Source, [int]$Index) {
    if ($Index -le 0) { return 1 }
    return [regex]::Matches($Source.Substring(0, $Index), "`n").Count + 1
}

function Test-AdjacentMarker([string[]]$Lines, [int]$LineNumber, [string]$Marker) {
    $first = [Math]::Max(0, $LineNumber - 4)
    $last = [Math]::Min($Lines.Count - 1, $LineNumber - 1)
    for ($index = $first; $index -le $last; $index++) {
        if ($Lines[$index].Contains($Marker)) { return $true }
    }
    return $false
}

function Add-Finding {
    param(
        [Collections.Generic.List[object]]$Findings,
        [string]$Rule,
        [string]$Path,
        [int]$Line,
        [string]$Signature,
        [string]$Message
    )

    $normalizedSignature = ConvertTo-NormalSignature $Signature
    $Findings.Add([pscustomobject]@{
        Rule = $Rule
        Path = $Path
        Line = $Line
        Signature = $normalizedSignature
        Fingerprint = Get-Fingerprint $Rule $Path $normalizedSignature
        Message = $Message
    }) | Out-Null
}

function Get-RelativePath([string]$Root, [string]$FullName) {
    return $FullName.Substring($Root.Length + 1).Replace('\', '/')
}

function Get-AttributeValue([string]$Tag, [string]$Name) {
    $match = [regex]::Match(
        $Tag,
        ('\b' + [regex]::Escape($Name) + '\s*=\s*"(?<value>[^"]*)"'),
        [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) { return $null }
    return $match.Groups['value'].Value
}

$script:SettingsAccessibleControlContracts = @(
    @{ Needle = 'ItemsSource="{Binding AvailableTranslationTools}"'; Name = 'Translation Provider' },
    @{ Needle = 'Text="{Binding DeepLApiKey,'; Name = 'DeepL API Key' },
    @{ Needle = 'Text="{Binding OpenAIApiKey,'; Name = 'OpenAI API Key' },
    @{ Needle = 'ItemsSource="{Binding AvailableOpenAIModels}"'; Name = 'OpenAI Model' },
    @{ Needle = 'Text="{Binding ClaudeApiKey,'; Name = 'Claude API Key' },
    @{ Needle = 'ItemsSource="{Binding AvailableClaudeModels}"'; Name = 'Claude Model' },
    @{ Needle = 'Text="{Binding LibreTranslateBaseUrl,'; Name = 'LibreTranslate Base URL' },
    @{ Needle = 'Text="{Binding LibreTranslateApiKey,'; Name = 'LibreTranslate API Key' },
    @{ Needle = 'ItemsSource="{Binding AvailableTargetLanguages}"'; Name = 'Translation Target Language' },
    @{ Needle = 'Text="{Binding StickyNoteBackgroundColor,'; Name = 'Sticky Note Background Color' },
    @{ Needle = 'Click="PickStickyBgColor_Click"'; Name = 'Choose sticky note background color' },
    @{ Needle = 'Text="{Binding StickyNoteFontColor,'; Name = 'Sticky Note Font Color' },
    @{ Needle = 'Click="PickStickyFontColor_Click"'; Name = 'Choose sticky note font color' },
    @{ Needle = 'Text="{Binding IllustrativeStickerTexts,'; Name = 'Status Sticker Available Texts' },
    @{ Needle = 'Text="{Binding IllustrativeStickerColor,'; Name = 'Status Sticker Color' },
    @{ Needle = 'Value="{Binding IllustrativeStickerLineSize,'; Name = 'Status Sticker Line Size' },
    @{ Needle = 'Text="{Binding IllustrativeStickerLineSize,'; Name = 'Status Sticker Line Size in points' },
    @{ Needle = 'ItemsSource="{Binding AvailableFonts}" SelectedItem="{Binding IllustrativeStickerFont}"'; Name = 'Status Sticker Font' },
    @{ Needle = 'Value="{Binding IllustrativeStickerFontSize,'; Name = 'Status Sticker Font Size' },
    @{ Needle = 'Text="{Binding IllustrativeStickerFontSize,'; Name = 'Status Sticker Font Size in points' },
    @{ Needle = 'SelectedValue="{Binding IllustrativeStickerPosition}"'; Name = 'Status Sticker Default Position' }
)

function Add-SettingsXamlContractFindings {
    param(
        [Collections.Generic.List[object]]$Findings,
        [string]$Source,
        [string]$Path
    )

    $swatchStyle = [regex]::Match($Source, '(?s)<Style\b(?<attributes>[^>]*\bx:Key="SwatchBtn"[^>]*)>(?<body>.*?)</Style>')
    $styleValid = $swatchStyle.Success -and
        (Get-AttributeValue $swatchStyle.Groups['attributes'].Value 'BasedOn') -eq '{StaticResource CompactIconButton}' -and
        $swatchStyle.Groups['body'].Value -match '<Setter\b[^>]*Property="Width"[^>]*Value="\{StaticResource IconButtonSizeCompact\}"[^>]*/>' -and
        $swatchStyle.Groups['body'].Value -match '<Setter\b[^>]*Property="Height"[^>]*Value="\{StaticResource IconButtonSizeCompact\}"[^>]*/>' -and
        $swatchStyle.Groups['body'].Value -notmatch '<ControlTemplate\b'
    if (-not $styleValid) {
        Add-Finding $Findings 'GoldenReferenceContract' $Path 1 'SwatchBtnStyle' `
            'Settings SwatchBtn must inherit CompactIconButton, use the compact 28-DIP size token, and not replace its template.'
    }

    $expectedSwatches = [ordered]@{
        '#FFB900' = 'Apply gold color'
        '#FF8C00' = 'Apply orange color'
        '#0078D4' = 'Apply blue color'
        '#00B294' = 'Apply teal color'
    }
    $swatches = @([regex]::Matches($Source, '<Button\b(?<attributes>[^>]*Style="\{StaticResource SwatchBtn\}"[^>]*)/?>'))
    $seenNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $seenToolTips = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    if ($swatches.Count -ne $expectedSwatches.Count) {
        Add-Finding $Findings 'GoldenReferenceContract' $Path 1 "StaticSwatchCount:$($swatches.Count)" `
            'Settings must expose exactly four static accent swatches.'
    }
    foreach ($swatch in $swatches) {
        $attributes = $swatch.Groups['attributes'].Value
        $tag = Get-AttributeValue $attributes 'Tag'
        $name = Get-AttributeValue $attributes 'AutomationProperties.Name'
        $toolTip = Get-AttributeValue $attributes 'ToolTip'
        $expectedName = if ($expectedSwatches.Contains($tag)) { $expectedSwatches[$tag] } else { $null }
        if ([string]::IsNullOrWhiteSpace($expectedName) -or $name -ne $expectedName -or $toolTip -ne $expectedName -or
            -not $seenNames.Add($name) -or -not $seenToolTips.Add($toolTip)) {
            Add-Finding $Findings 'GoldenReferenceContract' $Path (Get-LineNumber $Source $swatch.Index) "StaticSwatch:$tag" `
                'Every Settings accent swatch must keep its unique color tag, tooltip, and matching accessible name.'
        }
    }

    $controlTags = @([regex]::Matches($Source, '(?s)<(?:ComboBox|PasswordBox|TextBox|Slider|Button)\b(?<attributes>[^>]*)>'))
    foreach ($contract in $script:SettingsAccessibleControlContracts) {
        $matchingTag = @($controlTags | Where-Object { $_.Groups['attributes'].Value.Contains($contract.Needle) -and
                (Get-AttributeValue $_.Groups['attributes'].Value 'AutomationProperties.Name') -eq $contract.Name })
        if ($matchingTag.Count -ne 1) {
            Add-Finding $Findings 'GoldenReferenceContract' $Path 1 "AccessibleControl:$($contract.Needle)" `
                "Settings control '$($contract.Needle)' must have accessible name '$($contract.Name)'."
        }
    }

    foreach ($kind in @('Color', 'Font')) {
        $command = "DataContext.RemoveStyleCheck${kind}Command"
        $expectedName = "{Binding StringFormat='Remove approved $($kind.ToLowerInvariant()) {0}'}"
        $button = @($controlTags | Where-Object { $_.Groups['attributes'].Value.Contains($command) })
        if ($button.Count -ne 1 -or
            (Get-AttributeValue $button[0].Groups['attributes'].Value 'CommandParameter') -ne '{Binding}' -or
            (Get-AttributeValue $button[0].Groups['attributes'].Value 'AutomationProperties.Name') -ne $expectedName) {
            Add-Finding $Findings 'GoldenReferenceContract' $Path 1 "TokenRemoveButton:$kind" `
                "Settings $($kind.ToLowerInvariant()) token removal must keep its command, bound token parameter, and token-specific accessible name on one button."
        }
    }
}

function Add-SettingsComboBoxContractFindings {
    param(
        [Collections.Generic.List[object]]$Findings,
        [string]$SettingsSource,
        [string]$ControlsSource
    )

    $settingsStyle = [regex]::Match(
        $SettingsSource,
        '(?s)<Style\b(?<attributes>[^>]*\bx:Key="SettingsComboBox"[^>]*)>(?<body>.*?)</Style>')
    if (-not $settingsStyle.Success -or
        (Get-AttributeValue $settingsStyle.Groups['attributes'].Value 'TargetType') -ne 'ComboBox' -or
        (Get-AttributeValue $settingsStyle.Groups['attributes'].Value 'BasedOn') -ne '{StaticResource {x:Type ComboBox}}' -or
        $settingsStyle.Groups['body'].Value -notmatch '<Setter\b[^>]*Property="Width"[^>]*Value="320"' -or
        $settingsStyle.Groups['body'].Value -notmatch '<Setter\b[^>]*Property="MinHeight"[^>]*Value="\{StaticResource ControlHeightStandard\}"' -or
        $settingsStyle.Groups['body'].Value -match '<ControlTemplate\b') {
        Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/Dialogs/SettingsWindow.xaml' 1 'SettingsComboBoxStyle' `
            'SettingsComboBox must inherit the shared ComboBox template and add layout-only width, height, and alignment.'
    }

    $settingsComboBoxes = @([regex]::Matches(
        $SettingsSource,
        '<ComboBox\b[^>]*Style="\{StaticResource SettingsComboBox\}"'))
    $bypassedComboBoxes = @([regex]::Matches(
        $SettingsSource,
        '<ComboBox\b[^>]*Style="\{StaticResource SettingsFieldControl\}"'))
    if ($settingsComboBoxes.Count -ne 6 -or $bypassedComboBoxes.Count -ne 0) {
        Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/Dialogs/SettingsWindow.xaml' 1 `
            "SettingsComboBoxUsage:$($settingsComboBoxes.Count):$($bypassedComboBoxes.Count)" `
            'All six full-width Settings ComboBoxes must use SettingsComboBox; none may bypass the shared template through SettingsFieldControl.'
    }

    foreach ($required in @(
        'SurfaceBrush',
        'TextPrimaryBrush',
        'BorderBrush',
        'SurfaceHoverBrush',
        'SurfaceSelectedBrush',
        'AccentBrush',
        'TextMutedBrush',
        'IsKeyboardFocusWithin',
        'IsKeyboardFocused',
        'IsEnabled'
    )) {
        if (-not $ControlsSource.Contains($required)) {
            Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/Styles/Controls.xaml' 1 "ComboBoxState:$required" `
                "Shared ComboBox chrome must retain its semantic normal, hover, selected, focus, and disabled state contract; missing '$required'."
        }
    }
    if ($ControlsSource -notmatch '(?s)<Style\s+TargetType="ComboBoxItem".*?<Setter\s+TargetName="Bd"\s+Property="Background"\s+Value="\{DynamicResource SurfaceSelectedBrush\}"' -or
        $ControlsSource -notmatch '(?s)x:Name="FocusBorder".*?<Trigger\s+Property="IsKeyboardFocused"\s+Value="True">.*?TargetName="FocusBorder"' -or
        $ControlsSource -notmatch '(?s)<Style\s+TargetType="ComboBox">.*?<Trigger\s+Property="IsEnabled"\s+Value="False">.*?TextMutedBrush') {
        Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/Styles/Controls.xaml' 1 'ComboBoxSemanticStates' `
            'Shared ComboBox and ComboBoxItem templates must use semantic selected, keyboard-focus, and disabled treatments.'
    }
}

function Add-SettingsTextBoxContractFindings {
    param(
        [Collections.Generic.List[object]]$Findings,
        [string]$SettingsSource,
        [string]$ControlsSource
    )

    $settingsStyle = [regex]::Match(
        $SettingsSource,
        '(?s)<Style\b(?<attributes>[^>]*\bx:Key="SettingsTextBox"[^>]*)>(?<body>.*?)</Style>')
    if (-not $settingsStyle.Success -or
        (Get-AttributeValue $settingsStyle.Groups['attributes'].Value 'TargetType') -ne 'TextBox' -or
        (Get-AttributeValue $settingsStyle.Groups['attributes'].Value 'BasedOn') -ne '{StaticResource {x:Type TextBox}}' -or
        $settingsStyle.Groups['body'].Value -notmatch '<Setter\b[^>]*Property="Width"[^>]*Value="320"' -or
        $settingsStyle.Groups['body'].Value -notmatch '<Setter\b[^>]*Property="MinHeight"[^>]*Value="\{StaticResource ControlHeightStandard\}"' -or
        $settingsStyle.Groups['body'].Value -match '<ControlTemplate\b') {
        Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/Dialogs/SettingsWindow.xaml' 1 'SettingsTextBoxStyle' `
            'SettingsTextBox must inherit the shared TextBox template and add layout-only width, height, and alignment.'
    }

    $settingsTextBoxes = @([regex]::Matches(
        $SettingsSource,
        '<TextBox\b[^>]*Style="\{StaticResource SettingsTextBox\}"'))
    $bypassedTextBoxes = @([regex]::Matches(
        $SettingsSource,
        '<TextBox\b[^>]*Style="\{StaticResource SettingsFieldControl\}"'))
    if ($settingsTextBoxes.Count -ne 1 -or $bypassedTextBoxes.Count -ne 0) {
        Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/Dialogs/SettingsWindow.xaml' 1 `
            "SettingsTextBoxUsage:$($settingsTextBoxes.Count):$($bypassedTextBoxes.Count)" `
            'The full-width Settings TextBox must use SettingsTextBox; no TextBox may bypass the shared template through SettingsFieldControl.'
    }

    $sharedStyle = [regex]::Match(
        $ControlsSource,
        '(?s)<Style\s+TargetType="TextBox">(?<body>.*?)</Style>')
    foreach ($required in @(
        'Background" Value="{DynamicResource SurfaceBrush}',
        'Foreground" Value="{DynamicResource TextPrimaryBrush}',
        'BorderBrush" Value="{DynamicResource BorderBrush}',
        'CaretBrush" Value="{DynamicResource AccentBrush}',
        'Property="IsFocused" Value="True"',
        'Property="IsMouseOver" Value="True"',
        'Background" Value="{DynamicResource SurfaceHoverBrush}'
    )) {
        if (-not $sharedStyle.Success -or -not $sharedStyle.Groups['body'].Value.Contains($required)) {
            Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/Styles/Controls.xaml' 1 "TextBoxState:$required" `
                "Shared TextBox chrome must retain its semantic theme and interaction-state contract; missing '$required'."
        }
    }
}

function Add-StickerPreviewContractFindings {
    param(
        [Collections.Generic.List[object]]$Findings,
        [string]$SettingsSource,
        [string]$ViewModelSource,
        [string]$ModelSource,
        [string]$StatusFeatureSource,
        [string]$StickyFeatureSource,
        [string]$TestsSource
    )

    foreach ($required in @(
        'AutomationProperties.Name="Sticky Note Preview"',
        'AutomationProperties.Name="Status Sticker Preview"',
        'Text="{Binding StickyNotePreviewText}"',
        'Binding StickyNotePreviewStyle',
        'Binding StickyNotePreviewBackgroundColor',
        'Binding StickyNotePreviewFontColor',
        'Text="Colors apply to the selected Sticky Note style and do not follow the application theme."',
        'Content="Reset Selected Style Colors"',
        'GreenB preview reproduces the opaque black slide-content shadow block.',
        'Width="310" Height="54" Background="#000000"',
        'Text="{Binding StatusStickerPreviewText}"',
        'Height="{Binding StatusStickerPreviewLineSize}"',
        'FontSize="{Binding StatusStickerPreviewFontSize}"',
        'Binding StatusStickerPreviewUsesThemeColor',
        'Binding StatusStickerPreviewBold',
        'Binding StatusStickerPreviewItalic',
        'SelectedItem="{Binding IllustrativeStickerFont}"',
        'Value="{Binding IllustrativeStickerFontSize, Mode=TwoWay}"'
    )) {
        if (-not $SettingsSource.Contains($required)) {
            Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/Dialogs/SettingsWindow.xaml' 1 "StickerPreviewUi:$required" `
                "Stickers & Notes must retain its side-effect-free live preview contract; missing '$required'."
        }
    }

    foreach ($required in @(
        'public string StickyNotePreviewText',
        'public string StickyNotePreviewStyle',
        'public string StatusStickerPreviewText',
        'public double StatusStickerPreviewLineSize',
        'public double StatusStickerPreviewFontSize',
        'public bool StatusStickerPreviewUsesThemeColor',
        'OnPropertyChanged(nameof(StickyNoteBackgroundColor));',
        'OnPropertyChanged(nameof(StickyNoteFontColor));',
        'private void NotifyStickerPreviewChanged()'
    )) {
        if (-not $ViewModelSource.Contains($required)) {
            Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/SettingsViewModel.cs' 1 "StickerPreviewViewModel:$required" `
                "Sticker preview bindings must update from unsaved settings; missing '$required'."
        }
    }
    if ($ViewModelSource -match 'Microsoft\.Office\.Interop|PowerPointManager') {
        Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/SettingsViewModel.cs' 1 'StickerPreviewHostIsolation' `
            'SettingsViewModel must remain free of PowerPoint COM dependencies for sticker previews.'
    }

    foreach ($required in @(
        'public static StickyNotePreviewState CreateStickyNote',
        'public static StatusStickerPreviewState CreateStatusSticker',
        'public static string GetFirstStatusLabel',
        'public static string BuildStatusDisplayText',
        'public static string NormalizeHexColor',
        'public static void SetStickyNoteColors',
        'public static bool NormalizeStickyNoteColors',
        'public static string GetDefaultStickyColor',
        'case "Consulting": return background ? "#FFFF99" : "#333333";',
        'case "GreenB": return background ? "#FFFF00" : "#575757";',
        'default: return background ? "#FFF494" : "#323232";',
        'DefaultStickyText = "Sticky note"',
        'DefaultStatusText = "Work in progress"'
    )) {
        if (-not $ModelSource.Contains($required)) {
            Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency.Core/Models/StickerPreviewModel.cs' 1 "StickerPreviewModel:$required" `
                "Host-neutral sticker preview fallbacks must remain centralized; missing '$required'."
        }
    }
    if ($ModelSource -match 'Microsoft\.Office\.Interop|PowerPointManager' -or
        -not $StatusFeatureSource.Contains('StickerPreviewModel.GetFirstStatusLabel') -or
        [regex]::Matches($StickyFeatureSource, 'StickerPreviewModel\.CreateStickyNote\(config, null, "(?:Theme|Consulting|GreenB)"\)').Count -ne 3) {
        Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency.Core/Models/StickerPreviewModel.cs' 1 'StickerPreviewSharedInterpretation' `
            'The preview model must stay host-neutral while production Sticky Note and Status Sticker insertion share its style fallbacks.'
    }

    foreach ($required in @(
        'EmptySettingsProduceSafeRepresentativePreviews',
        'StatusPreviewUsesFirstLabelAndConfiguredMetadata',
        'StickyPreviewKeepsThreeSupportedVisualStylesDistinct',
        'CustomColorsAreScopedToSelectedStickyStyle',
        'LegacyCustomColorIsRetainedForFormerlyActiveStyleOnly',
        'LegacyPaletteNormalizationPinsCustomColorsToFormerActiveStyle'
    )) {
        if (-not $TestsSource.Contains($required)) {
            Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency.Core.Tests/StickerPreviewModelTests.cs' 1 "StickerPreviewTest:$required" `
                "Sticker preview fallback and style coverage lost '$required'."
        }
    }
}

function Add-SidebarThemeContractFindings {
    param(
        [Collections.Generic.List[object]]$Findings,
        [string]$LightSource,
        [string]$DarkSource,
        [string]$ControlsSource,
        [string]$SidebarSource,
        [string]$ViewModelSource,
        [string]$SettingsSource,
        [string]$SettingsCodeSource,
        [string]$MigrationTestsSource
    )

    $themeContracts = @(
        @{ Source = $LightSource; Path = 'oPenEfficiency/UI/Styles/Themes/Light.xaml'; Key = 'SidebarSectionHeader'; Color = '#000000' },
        @{ Source = $DarkSource; Path = 'oPenEfficiency/UI/Styles/Themes/Dark.xaml'; Key = 'SidebarSectionHeader'; Color = '#FFFFFF' },
        @{ Source = $LightSource; Path = 'oPenEfficiency/UI/Styles/Themes/Light.xaml'; Key = 'SidebarDivider'; Color = '#B8BEC7' },
        @{ Source = $DarkSource; Path = 'oPenEfficiency/UI/Styles/Themes/Dark.xaml'; Key = 'SidebarDivider'; Color = '#656565' }
    )
    foreach ($contract in $themeContracts) {
        $colorPattern = '<Color\s+x:Key="' + [regex]::Escape($contract.Key + 'Color') + '">' + [regex]::Escape($contract.Color) + '</Color>'
        $brushPattern = '<SolidColorBrush\s+x:Key="' + [regex]::Escape($contract.Key + 'Brush') + '"\s+Color="' + [regex]::Escape($contract.Color) + '"\s*/>'
        if ($contract.Source -notmatch $colorPattern -or $contract.Source -notmatch $brushPattern) {
            Add-Finding $Findings 'GoldenReferenceContract' $contract.Path 1 "ThemeRole:$($contract.Key)" `
                "$($contract.Key) must retain its dedicated $($contract.Color) color/brush pair."
        }
    }
    if ($LightSource -notmatch '<Color\s+x:Key="BorderColor">#D1D5DB</Color>' -or
        $DarkSource -notmatch '<Color\s+x:Key="BorderColor">#2D2D35</Color>') {
        Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/Styles/Themes' 1 'GlobalBorderIsolation' `
            'Sidebar refinement must not change generic Light or Dark BorderColor resources.'
    }

    $sectionStyle = [regex]::Match($ControlsSource, '(?s)<Style\b[^>]*x:Key="SectionHeader"[^>]*>(?<body>.*?)</Style>')
    $sidebarSectionStyle = [regex]::Match($ControlsSource, '(?s)<Style\b(?<attributes>[^>]*x:Key="SidebarSectionHeader"[^>]*)>(?<body>.*?)</Style>')
    if (-not $sectionStyle.Success -or
        $sectionStyle.Groups['body'].Value -notmatch '<Setter\b[^>]*Property="Foreground"[^>]*Value="\{DynamicResource TextSecondaryBrush\}"' -or
        -not $sidebarSectionStyle.Success -or
        (Get-AttributeValue $sidebarSectionStyle.Groups['attributes'].Value 'BasedOn') -ne '{StaticResource SectionHeader}' -or
        $sidebarSectionStyle.Groups['body'].Value -notmatch '<Setter\b[^>]*Property="Foreground"[^>]*Value="\{DynamicResource SidebarSectionHeaderBrush\}"') {
        Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/Styles/Controls.xaml' 1 'SectionHeaderThemeRole' `
            'SidebarSectionHeader must derive from the generic SectionHeader and scope the live Sidebar brush to the task pane.'
    }

    if ([regex]::Matches($SidebarSource, 'SetResourceReference\([^;]*"SidebarDividerBrush"\)').Count -ne 3 -or
        [regex]::Matches($SidebarSource, 'BorderThickness = new Thickness\(0, 0, 0\.5, 0\)').Count -ne 2 -or
        [regex]::Matches($SidebarSource, 'Height = 0\.5').Count -ne 1 -or
        [regex]::Matches($SidebarSource, 'FindResource\("SidebarSectionHeader"\)').Count -ne 4 -or
        [regex]::Matches($SidebarSource, 'ApplySectionHeaderForeground\((?:header|settingsHeader), config\);').Count -ne 4 -or
        $SidebarSource -notmatch 'SidebarSectionFontColorModes\.UsesTheme\(config\)' -or
        $SidebarSource -notmatch 'SetResourceReference\(TextBlock\.ForegroundProperty, "SidebarSectionHeaderBrush"\)') {
        Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/Panels/MainSidebar.xaml.cs' 1 'SidebarThemeApplication' `
            'All three half-weight Sidebar separators and Auto section headers must consume their dedicated live theme resources.'
    }

    foreach ($required in @(
        'public bool UseThemeSectionFontColor',
        'public bool UseCustomSectionFontColor',
        '_config.SectionFontColorMode = SidebarSectionFontColorModes.Custom;',
        'SidebarSectionFontColorModes.Normalize(Config);',
        'UseThemeSectionFontColor = true;',
        'exportData.SectionFontColorMode = Config.SectionFontColorMode;',
        'Config.SectionFontColorMode = imported.SectionFontColorMode;'
    )) {
        if (-not $ViewModelSource.Contains($required)) {
            Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/SettingsViewModel.cs' 1 "SectionMode:$required" `
                "Settings section-header Auto/Custom contract lost '$required'."
        }
    }
    foreach ($required in @(
        'IsChecked="{Binding UseThemeSectionFontColor}"',
        'IsChecked="{Binding UseCustomSectionFontColor}"',
        'IsEnabled="{Binding UseCustomSectionFontColor}"',
        'AutomationProperties.Name="Automatically follow app theme for section header text"',
        'AutomationProperties.Name="Use custom section header text color"',
        'Value="{DynamicResource SidebarSectionHeaderBrush}"',
        'Fill="{DynamicResource SidebarDividerBrush}"'
    )) {
        if (-not $SettingsSource.Contains($required)) {
            Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/Dialogs/SettingsWindow.xaml' 1 "SectionModeUi:$required" `
                "Settings section-header Auto/Custom UI contract lost '$required'."
            }
        }
    if ($SettingsSource -notmatch '(?s)x:Name="SidebarPreviewBorder".*?Binding UseThemeBackground.*?DynamicResource SidebarBackgroundBrush.*?x:Name="PreviewLabel1".*?DynamicResource SidebarSectionHeaderBrush.*?Height="0\.75"\s+Fill="\{DynamicResource SidebarDividerBrush\}"') {
        Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/Dialogs/SettingsWindow.xaml' 1 'SidebarPreviewThemeRoles' `
            'The Settings live preview must resolve Auto background, header, and divider through Sidebar theme resources.'
    }
    foreach ($required in @(
        '_viewModel.PropertyChanged += ViewModel_PropertyChanged;',
        'ThemeManager.ThemeChanged += ThemeManager_ThemeChanged;',
        'ResolveContrastColor(',
        '"SidebarBackgroundBrush"',
        '"SidebarSectionHeaderBrush"'
    )) {
        if (-not $SettingsCodeSource.Contains($required)) {
            Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency/UI/Dialogs/SettingsWindow.xaml.cs' 1 "ContrastMode:$required" `
                "Settings contrast feedback must refresh from the effective Auto resources; missing '$required'."
        }
    }
    foreach ($required in @(
        'NewSidebarConfigTracksThemeSectionHeaderByDefault',
        'LegacyDefaultSectionHeaderIsMigratedToThemeMode',
        'LegacyCustomSectionHeaderIsPreservedAsCustom',
        'ThemeSectionHeaderModeRetainsSavedCustomColor'
    )) {
        if (-not $MigrationTestsSource.Contains($required)) {
            Add-Finding $Findings 'GoldenReferenceContract' 'oPenEfficiency.Core.Tests/SidebarMigrationBoundaryTests.cs' 1 "SectionModeTest:$required" `
                "Sidebar section-header migration coverage lost '$required'."
        }
    }
}

function Get-VisualFindings {
    param(
        [string]$Root,
        [bool]$RequireProductionContracts
    )

    $findings = [Collections.Generic.List[object]]::new()
    $uiRoot = Join-Path $Root "oPenEfficiency/UI"
    $projectRoot = Join-Path $Root "oPenEfficiency"
    $controlsPath = Join-Path $uiRoot "Styles/Controls.xaml"
    if (-not (Test-Path -LiteralPath $controlsPath -PathType Leaf)) {
        throw "Visual consistency scan requires '$controlsPath'."
    }

    $xamlFiles = @(Get-ChildItem -LiteralPath $uiRoot -Filter "*.xaml" -File -Recurse)
    $controlsSource = Get-Content -LiteralPath $controlsPath -Raw
    $canonicalStyleKeys = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($match in [regex]::Matches($controlsSource, '<Style\b(?<attributes>[^>]*)>', 'IgnoreCase')) {
        $styleKey = Get-AttributeValue $match.Groups['attributes'].Value 'x:Key'
        if (-not [string]::IsNullOrWhiteSpace($styleKey)) {
            $canonicalStyleKeys.Add($styleKey) | Out-Null
        }
    }

    $rootSurfaces = [Collections.Generic.List[object]]::new()
    foreach ($file in $xamlFiles) {
        $relativePath = Get-RelativePath $Root $file.FullName
        $source = Get-Content -LiteralPath $file.FullName -Raw
        $lines = @($source -split "`r?`n")
        $isThemeDictionary = $relativePath -match '^oPenEfficiency/UI/Styles/Themes/'

        if (-not $isThemeDictionary) {
            $colorPattern = '\b(?<property>Background|Foreground|BorderBrush|Fill|Stroke|CaretBrush|Color)\s*=\s*"(?<value>[^"{][^"]*)"'
            foreach ($match in [regex]::Matches($source, $colorPattern, 'IgnoreCase')) {
                $value = $match.Groups['value'].Value.Trim()
                if ($value -eq 'Transparent') { continue }
                $lineNumber = Get-LineNumber $source $match.Index
                if (Test-AdjacentMarker $lines $lineNumber 'VisualColorException:') { continue }
                Add-Finding $findings 'HardCodedChromeColor' $relativePath $lineNumber `
                    $match.Value "literal $($match.Groups['property'].Value) '$value' must use a semantic DynamicResource or a documented VisualColorException."
            }
        }

        foreach ($match in [regex]::Matches($source, '\bFontFamily\s*=\s*"(?<value>[^"{][^"]*)"', 'IgnoreCase')) {
            $lineNumber = Get-LineNumber $source $match.Index
            if (Test-AdjacentMarker $lines $lineNumber 'VisualFontException:') { continue }
            Add-Finding $findings 'UnauthorizedFontFamily' $relativePath $lineNumber `
                $match.Value "literal FontFamily '$($match.Groups['value'].Value)' bypasses AppFontFamily; document a content-font exception when intentional."
        }

        if ($relativePath -ne 'oPenEfficiency/UI/Styles/Controls.xaml') {
            foreach ($match in [regex]::Matches($source, '<DropShadowEffect\b[^>]*/?>', 'IgnoreCase')) {
                $lineNumber = Get-LineNumber $source $match.Index
                if (Test-AdjacentMarker $lines $lineNumber 'VisualShadowException:') { continue }
                Add-Finding $findings 'AdHocWindowShadow' $relativePath $lineNumber `
                    $match.Value 'local DropShadowEffect bypasses the shared elevation tokens.'
            }

            foreach ($match in [regex]::Matches($source, '<Style\b(?<attributes>[^>]*)>', 'IgnoreCase')) {
                $styleKey = Get-AttributeValue $match.Groups['attributes'].Value 'x:Key'
                if (-not [string]::IsNullOrWhiteSpace($styleKey) -and $canonicalStyleKeys.Contains($styleKey)) {
                    Add-Finding $findings 'DuplicateCanonicalStyle' $relativePath (Get-LineNumber $source $match.Index) `
                        $match.Value "local style '$styleKey' duplicates a canonical Controls.xaml key."
                }
            }

            foreach ($match in [regex]::Matches(
                $source,
                '(?is)<Style\b(?<attributes>[^>]*)TargetType\s*=\s*"Button"[^>]*>(?<body>.*?)</Style>')) {
                if ($match.Groups['body'].Value -notmatch 'Property\s*=\s*"Template"') { continue }
                $lineNumber = Get-LineNumber $source $match.Index
                if (Test-AdjacentMarker $lines $lineNumber 'VisualButtonTemplateException:') { continue }
                Add-Finding $findings 'LocalCanonicalButtonTemplate' $relativePath $lineNumber `
                    $match.Value 'local Button ControlTemplate bypasses canonical primary/secondary/icon/toolbar styles.'
            }
        }

        $rootMatch = [regex]::Match($source, '(?s)^\s*<(?<type>Window|UserControl)\b(?<attributes>.*?)>')
        if ($rootMatch.Success) {
            $rootSurfaces.Add([pscustomobject]@{
                File = $file
                Path = $relativePath
                Source = $source
                Lines = $lines
                Type = $rootMatch.Groups['type'].Value
                RootTag = $rootMatch.Value
            }) | Out-Null
        }

        if ($source -match '(?i)_hover\.svg') {
            foreach ($match in [regex]::Matches($source, '(?i)[^"''\s<>]*_hover\.svg')) {
                Add-Finding $findings 'AccidentalSecondIconStateSystem' $relativePath (Get-LineNumber $source $match.Index) `
                    $match.Value 'hover-specific SVG references create a second icon-state system; use control state styling or document an exception.'
            }
        }
        foreach ($match in [regex]::Matches(
            $source,
            '(?is)<Trigger\b[^>]*Property\s*=\s*"IsMouseOver"[^>]*>.*?<Setter\b[^>]*Property\s*=\s*"Source"[^>]*/>.*?</Trigger>')) {
            Add-Finding $findings 'AccidentalSecondIconStateSystem' $relativePath (Get-LineNumber $source $match.Index) `
                $match.Value 'mouse-over swaps an image Source instead of using the canonical control/icon state treatment.'
        }
    }

    foreach ($surface in $rootSurfaces) {
        $source = $surface.Source
        $styleValue = Get-AttributeValue $surface.RootTag 'Style'
        $isWindow = $surface.Type -eq 'Window'
        $usesSidebarStyle = $styleValue -match 'SidebarTaskPaneStyle'
        if (($isWindow -or $usesSidebarStyle) -and
            -not $source.Contains('Styles/Controls.xaml') -and
            -not $source.Contains('VisualResourceException:')) {
            Add-Finding $findings 'MissingSharedVisualResources' $surface.Path 1 `
                "$($surface.Type)|$styleValue" 'surface does not merge Controls.xaml and has no documented VisualResourceException.'
        }

        $archetypeContracts = [ordered]@{
            'ContextPopupWindowStyle' = 'ContextPopupShell'
            'FloatingPaletteWindowStyle' = 'FloatingPaletteShell'
            'CompactTaskWindowStyle' = 'CompactTaskShell'
            'CompactModalWindowStyle' = 'CompactModalShell'
            'LargeWorkspaceWindowStyle' = 'WorkspaceShell'
            'SidebarTaskPaneStyle' = 'SidebarShell'
        }
        $matchedStyle = $null
        foreach ($candidate in $archetypeContracts.Keys) {
            if ($styleValue -match [regex]::Escape($candidate)) {
                $matchedStyle = $candidate
                break
            }
        }

        if ($isWindow -and $null -eq $matchedStyle -and -not $source.Contains('ArchetypeException:')) {
            $className = Get-AttributeValue $surface.RootTag 'x:Class'
            Add-Finding $findings 'InvalidArchetypeWindowContract' $surface.Path 1 `
                "unclassified|$className" 'Window has no canonical archetype style and no documented ArchetypeException.'
        }
        elseif ($null -ne $matchedStyle) {
            $requiredShell = $archetypeContracts[$matchedStyle]
            if ($source -notmatch ('\bStyle\s*=\s*"\{(?:Static|Dynamic)Resource\s+' + [regex]::Escape($requiredShell) + '\}"')) {
                Add-Finding $findings 'InvalidArchetypeWindowContract' $surface.Path 1 `
                    "$matchedStyle|missing:$requiredShell" "archetype '$matchedStyle' must use shell '$requiredShell'."
            }

            if ($isWindow) {
                $allowsTransparency = Get-AttributeValue $surface.RootTag 'AllowsTransparency'
                $windowStyle = Get-AttributeValue $surface.RootTag 'WindowStyle'
                $showInTaskbar = Get-AttributeValue $surface.RootTag 'ShowInTaskbar'
                $topmost = Get-AttributeValue $surface.RootTag 'Topmost'
                $background = Get-AttributeValue $surface.RootTag 'Background'
                $resizeMode = Get-AttributeValue $surface.RootTag 'ResizeMode'
                $sizeToContent = Get-AttributeValue $surface.RootTag 'SizeToContent'
                $conflicts = [Collections.Generic.List[string]]::new()
                if ($allowsTransparency -and $allowsTransparency -ne 'True') { $conflicts.Add("AllowsTransparency=$allowsTransparency") }
                if ($windowStyle -and $windowStyle -ne 'None') { $conflicts.Add("WindowStyle=$windowStyle") }
                if ($showInTaskbar -and $showInTaskbar -ne 'False') { $conflicts.Add("ShowInTaskbar=$showInTaskbar") }
                if ($topmost -eq 'True') { $conflicts.Add('Topmost=True') }
                if ($background -and $background -ne 'Transparent') { $conflicts.Add("Background=$background") }
                if ($matchedStyle -eq 'LargeWorkspaceWindowStyle') {
                    if ($resizeMode -eq 'NoResize') { $conflicts.Add('ResizeMode=NoResize') }
                    if ($sizeToContent -eq 'WidthAndHeight') { $conflicts.Add('SizeToContent=WidthAndHeight') }
                }
                elseif ($resizeMode -and $resizeMode -ne 'NoResize') {
                    $conflicts.Add("ResizeMode=$resizeMode")
                }
                if ($conflicts.Count -gt 0) {
                    Add-Finding $findings 'InvalidArchetypeWindowContract' $surface.Path 1 `
                        "$matchedStyle|$($conflicts -join ',')" "archetype '$matchedStyle' has conflicting root properties: $($conflicts -join ', ')."
                }
            }
        }

        if ($isWindow) {
            if ($surface.Path -ne 'oPenEfficiency/UI/Dialogs/FormatNumbersWindow.xaml' -and
                $source.Contains('DialogSubtitleText')) {
                Add-Finding $findings 'HeaderCopyPolicy' $surface.Path 1 'DialogSubtitleText' `
                    'unapproved DialogSubtitleText usage; use body copy or ContextMetaText.'
            }
            $titleAdjacentCopy = [regex]::Matches(
                $source,
                '(?s)<StackPanel[^>]*>\s*<TextBlock[^>]*DialogTitleText[^>]*/>\s*<TextBlock(?![^>]*ContextMetaText)[^>]*>')
            if ($titleAdjacentCopy.Count -gt 0) {
                Add-Finding $findings 'HeaderCopyPolicy' $surface.Path 1 $titleAdjacentCopy[0].Value `
                    'unreviewed title-adjacent explanatory copy.'
            }
        }
    }

    if (Test-Path -LiteralPath $projectRoot -PathType Container) {
        $iconFiles = @(Get-ChildItem -LiteralPath $projectRoot -File -Recurse |
            Where-Object {
                $_.FullName -notmatch '[\\/](bin|obj|packages)[\\/]' -and
                $_.Name -match '(?i)_hover\.svg$'
            })
        foreach ($file in $iconFiles) {
            $relativePath = Get-RelativePath $Root $file.FullName
            Add-Finding $findings 'AccidentalSecondIconStateSystem' $relativePath 1 $file.Name `
                'hover-specific SVG asset creates a second icon-state system.'
        }
    }

    if ($RequireProductionContracts) {
        $goldenReferences = @(
            'oPenEfficiency/UI/Menus/TrafficLightMenu.xaml',
            'oPenEfficiency/UI/Toolbars/GanttToolbar.xaml',
            'oPenEfficiency/UI/Dialogs/ArrangeGridWindow.xaml',
            'oPenEfficiency/UI/Panels/MainSidebar.xaml',
            'oPenEfficiency/UI/Dialogs/FeatureExplorerWindow.xaml'
        )
        $semanticThemeRoles = @(
            'WindowBackground', 'Surface', 'SurfaceLight', 'SurfaceHover', 'SurfaceSelected',
            'WindowHeader', 'Accent', 'AccentHover', 'AccentText', 'TextPrimary', 'TextSecondary',
            'TextMuted', 'Border', 'BorderLight', 'Success', 'Warning', 'Error'
        )
        foreach ($relativePath in $goldenReferences) {
            $fullPath = Join-Path $Root $relativePath
            if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
                Add-Finding $findings 'GoldenReferenceContract' $relativePath 1 'missing' 'Golden Reference surface is missing.'
                continue
            }
            $source = Get-Content -LiteralPath $fullPath -Raw
            foreach ($role in $semanticThemeRoles) {
                if ($source -match ('\{StaticResource\s+' + [regex]::Escape($role) + '(?:Brush|Color)\}')) {
                    Add-Finding $findings 'GoldenReferenceContract' $relativePath 1 "StaticResource:$role" `
                        "theme role '$role' must use DynamicResource."
                }
            }
        }

        $controlsPath = 'oPenEfficiency/UI/Styles/Controls.xaml'
        $controlsSource = Get-Content -LiteralPath (Join-Path $Root $controlsPath) -Raw
        foreach ($required in @(
            '<Thickness x:Key="FloatingPaletteEdgePadding">6</Thickness>',
            '<Setter Property="Padding" Value="{StaticResource FloatingPaletteEdgePadding}"/>'
        )) {
            if (-not $controlsSource.Contains($required)) {
                Add-Finding $findings 'GoldenReferenceContract' $controlsPath 1 "missing:$required" `
                    "compact floating palettes lost calibrated shared geometry '$required'."
            }
        }

        $trafficPath = 'oPenEfficiency/UI/Menus/TrafficLightMenu.xaml'
        $trafficSource = Get-Content -LiteralPath (Join-Path $Root $trafficPath) -Raw
        foreach ($required in @(
            '<StackPanel Orientation="Horizontal">',
            'BasedOn="{StaticResource CompactIconButton}"',
            '<Setter Property="Margin" Value="1,0"/>',
            'BasedOn="{StaticResource CompactActionButton}"',
            '<Setter Property="Padding" Value="6,0"/>',
            '<Setter Property="Margin" Value="4,0"/>',
            'ToolTip="Set outline frame weight in points"',
            'AutomationProperties.Name="Outline frame weight in points"'
        )) {
            if (-not $trafficSource.Contains($required)) {
                Add-Finding $findings 'GoldenReferenceContract' $trafficPath 1 "missing:$required" `
                    "Traffic Light compact one-row contract lost '$required'."
            }
        }

        $ganttPath = 'oPenEfficiency/UI/Toolbars/GanttToolbar.xaml'
        $ganttSource = Get-Content -LiteralPath (Join-Path $Root $ganttPath) -Raw
        foreach ($required in @(
            'Style="{DynamicResource FloatingPaletteShell}"',
            'Style="{DynamicResource CompactComboBox}"',
            'Width="104"',
            'Width="{StaticResource IconButtonSizeCompact}"',
            'Height="{StaticResource IconButtonSizeCompact}"'
        )) {
            if (-not $ganttSource.Contains($required)) {
                Add-Finding $findings 'GoldenReferenceContract' $ganttPath 1 "missing:$required" `
                    "Gantt compact floating-palette contract lost '$required'."
            }
        }
        $ganttAutoColumns = [regex]::Matches($ganttSource, '<ColumnDefinition\s+Width="Auto"\s*/>').Count
        if ($ganttAutoColumns -ne 4) {
            Add-Finding $findings 'GoldenReferenceContract' $ganttPath 1 "AutoColumns:$ganttAutoColumns" `
                'Gantt palette must remain one content-sized row with four columns.'
        }
        $ganttDragHandlers = [regex]::Matches($ganttSource, 'MouseLeftButtonDown="DragSurface_MouseLeftButtonDown"').Count
        if ($ganttDragHandlers -ne 1 -or
            $ganttSource -notmatch '(?s)<Border[^>]*MouseLeftButtonDown="DragSurface_MouseLeftButtonDown"[^>]*>\s*<TextBlock') {
            Add-Finding $findings 'GoldenReferenceContract' $ganttPath 1 "DragHandlers:$ganttDragHandlers" `
                'Gantt drag handling must remain scoped to its compact label/header region.'
        }

        $settingsPath = 'oPenEfficiency/UI/Dialogs/SettingsWindow.xaml'
        $settingsSource = Get-Content -LiteralPath (Join-Path $Root $settingsPath) -Raw
        foreach ($required in @(
            'Style="{DynamicResource LargeWorkspaceWindowStyle}"',
            'Style="{StaticResource SettingsPage}"',
            'Style="{StaticResource SettingsSectionTitle}"',
            'Style="{StaticResource SettingsCard}"',
            'Style="{StaticResource SettingsFormRow}"',
            'Style="{StaticResource SettingsFieldLabel}"',
            'Style="{StaticResource SettingsTextBox}"',
            'Style="{StaticResource SettingsComboBox}"',
            'Property="Background" Value="{DynamicResource SurfaceSelectedBrush}"',
            'Style="{StaticResource ErrorStatusBanner}"',
            'Style="{StaticResource DestructiveButton}"',
            'ItemsSource="{Binding StyleCheckColorTokens}"',
            'ItemsSource="{Binding StyleCheckFontTokens}"',
            'Command="{Binding AddStyleCheckColorCommand}"',
            'Command="{Binding AddStyleCheckFontCommand}"',
            'Content="Auto (follow app theme)"',
            'IsChecked="{Binding UseCustomBackground}"',
            'AutomationProperties.Name="Anchor Object"',
            'AutomationProperties.Name="Spacing Increment"',
            'AutomationProperties.Name="Application Theme"',
            'AutomationProperties.Name="Application UI Font"',
            'AutomationProperties.Name="Sticky Note Default Text"',
            'AutomationProperties.Name="Sticky Note Font"',
            'AutomationProperties.Name="Export Options"',
            'AutomationProperties.Name="Export Options"'
        )) {
            if (-not $settingsSource.Contains($required)) {
                Add-Finding $findings 'GoldenReferenceContract' $settingsPath 1 "missing:$required" `
                    "Settings compact Large Workspace contract lost '$required'."
            }
        }
        if ([regex]::Matches($settingsSource, 'Style="\{StaticResource SettingsFormRow\}"').Count -lt 5) {
            Add-Finding $findings 'GoldenReferenceContract' $settingsPath 1 'SettingsFormRows' `
                'Settings ordinary pages must reuse the aligned label/control form-row grammar.'
        }
        Add-SettingsXamlContractFindings $findings $settingsSource $settingsPath
        Add-SettingsComboBoxContractFindings $findings $settingsSource $controlsSource
        Add-SettingsTextBoxContractFindings $findings $settingsSource $controlsSource
        $navColumns = [regex]::Match($settingsSource, '(?s)<Grid\.ColumnDefinitions>\s*<ColumnDefinition Width="(?<width>\d+)"/>\s*<ColumnDefinition Width="1"/>')
        if (-not $navColumns.Success -or [int]$navColumns.Groups['width'].Value -ge 220) {
            Add-Finding $findings 'GoldenReferenceContract' $settingsPath 1 'NavigationWidth' `
                'Settings navigation must remain content-fit and narrower than the legacy 220 DIP column.'
        }
        $preview = [regex]::Match($settingsSource, '<Border\b(?<attributes>[^>]*\bx:Name="SidebarPreviewBorder"[^>]*)>')
        $previewWidth = if ($preview.Success) { Get-AttributeValue $preview.Groups['attributes'].Value 'Width' } else { $null }
        $previewHeight = if ($preview.Success) { Get-AttributeValue $preview.Groups['attributes'].Value 'Height' } else { $null }
        if (-not $preview.Success -or [double]$previewWidth -lt 120 -or [double]$previewHeight -lt 180) {
            Add-Finding $findings 'GoldenReferenceContract' $settingsPath 1 'SidebarPreviewSize' `
                'Settings live preview must remain large enough to evaluate sidebar appearance.'
        }
        $footer = [regex]::Match($settingsSource, '<Border\b(?<attributes>[^>]*\bx:Name="SettingsFooter"[^>]*)>')
        if (-not $footer.Success -or
            (Get-AttributeValue $footer.Groups['attributes'].Value 'Background') -ne '{DynamicResource SurfaceLightBrush}' -or
            (Get-AttributeValue $footer.Groups['attributes'].Value 'Padding') -ne '{StaticResource Space10}') {
            Add-Finding $findings 'GoldenReferenceContract' $settingsPath 1 'SettingsFooter' `
                'Settings sticky footer must use the compact semantic workspace treatment.'
        }
        if ($settingsSource -match 'Margin="[^"]*\{(?:Static|Dynamic)Resource [^}]+\}[^"]*,') {
            Add-Finding $findings 'InvalidMarkupExtensionComposition' $settingsPath 1 $Matches[0] `
                'A markup extension cannot be concatenated with comma-separated Thickness values.'
        }
        foreach ($forbidden in @(
            'Width="1020" Height="740"',
            'Margin="30,30,30,30"',
            'Background="#2A1515"',
            'Text="Colour Options"',
            'Background="#1D2E3E"',
            'Background="#2E251D"',
            'Text="{Binding StyleCheckCorrectColors',
            'Text="{Binding StyleCheckCorrectFonts'
        )) {
            if ($settingsSource.Contains($forbidden)) {
                Add-Finding $findings 'GoldenReferenceContract' $settingsPath 1 "forbidden:$forbidden" `
                    "Settings compact Large Workspace contract regressed to '$forbidden'."
            }
        }

        $settingsViewModelPath = 'oPenEfficiency/UI/SettingsViewModel.cs'
        $settingsViewModelSource = Get-Content -LiteralPath (Join-Path $Root $settingsViewModelPath) -Raw
        $stickerPreviewModelSource = Get-Content -LiteralPath (Join-Path $Root 'oPenEfficiency.Core/Models/StickerPreviewModel.cs') -Raw
        $statusStickerFeatureSource = Get-Content -LiteralPath (Join-Path $Root 'oPenEfficiency/Features/Visuals/IllustrativeStickerFeature.cs') -Raw
        $stickyNoteFeatureSource = Get-Content -LiteralPath (Join-Path $Root 'oPenEfficiency/Features/Utilities/StickyNoteFeature.cs') -Raw
        $stickerPreviewTestsSource = Get-Content -LiteralPath (Join-Path $Root 'oPenEfficiency.Core.Tests/StickerPreviewModelTests.cs') -Raw
        Add-StickerPreviewContractFindings $findings $settingsSource $settingsViewModelSource `
            $stickerPreviewModelSource $statusStickerFeatureSource $stickyNoteFeatureSource $stickerPreviewTestsSource
        foreach ($required in @(
            'DelimitedSettingTokens.Parse(serialized)',
            'DelimitedSettingTokens.Serialize(StyleCheckColorTokens)',
            'DelimitedSettingTokens.TryNormalizeHexColor(',
            'DelimitedSettingTokens.TryNormalizeFontName(',
            'public bool UseCustomBackground',
            'if (value) UseThemeBackground = false;'
        )) {
            if (-not $settingsViewModelSource.Contains($required)) {
                Add-Finding $findings 'GoldenReferenceContract' $settingsViewModelPath 1 "missing:$required" `
                    "Settings token/color interaction contract lost '$required'."
            }
        }
        if ($settingsViewModelSource -notmatch '(?s)Config\.BackgroundMode = imported\.BackgroundMode;.*?OnPropertyChanged\(nameof\(UseThemeBackground\)\);\s*OnPropertyChanged\(nameof\(UseCustomBackground\)\);') {
            Add-Finding $findings 'GoldenReferenceContract' $settingsViewModelPath 1 'ImportedBackgroundModeNotifications' `
                'Settings import must refresh both Auto and Custom background-mode controls.'
        }

        $sidebarPath = 'oPenEfficiency/UI/Panels/MainSidebar.xaml.cs'
        $sidebarSource = Get-Content -LiteralPath (Join-Path $Root $sidebarPath) -Raw
        foreach ($forbidden in @('#9CA3AF', 'Color.FromRgb(85, 85, 85)', 'FontSize = 9')) {
            if ($sidebarSource.Contains($forbidden)) {
                Add-Finding $findings 'DynamicSidebarChrome' $sidebarPath 1 $forbidden `
                    "dynamic product chrome still contains '$forbidden'."
            }
        }
        foreach ($required in @('FindResource("SidebarSectionHeader")', '"SidebarDividerBrush"', 'FindResource("Space4")')) {
            if (-not $sidebarSource.Contains($required)) {
                Add-Finding $findings 'DynamicSidebarChrome' $sidebarPath 1 "missing:$required" `
                    "dynamic Sidebar lost shared resource use '$required'."
            }
        }

        $lightThemeSource = Get-Content -LiteralPath (Join-Path $Root 'oPenEfficiency/UI/Styles/Themes/Light.xaml') -Raw
        $darkThemeSource = Get-Content -LiteralPath (Join-Path $Root 'oPenEfficiency/UI/Styles/Themes/Dark.xaml') -Raw
        $settingsCodeSource = Get-Content -LiteralPath (Join-Path $Root 'oPenEfficiency/UI/Dialogs/SettingsWindow.xaml.cs') -Raw
        $migrationTestsSource = Get-Content -LiteralPath (Join-Path $Root 'oPenEfficiency.Core.Tests/SidebarMigrationBoundaryTests.cs') -Raw
        Add-SidebarThemeContractFindings $findings $lightThemeSource $darkThemeSource $controlsSource `
            $sidebarSource $settingsViewModelSource $settingsSource $settingsCodeSource $migrationTestsSource

        $factoryPath = 'oPenEfficiency/UI/Panels/SidebarControlFactory.cs'
        $factorySource = Get-Content -LiteralPath (Join-Path $Root $factoryPath) -Raw
        foreach ($required in @('Text = info.Name', 'info.Tooltip', 'Right-click for more options', 'info.Variants != null')) {
            if (-not $factorySource.Contains($required)) {
                Add-Finding $findings 'DynamicSidebarChrome' $factoryPath 1 "missing:$required" `
                    "hover-card contract lost '$required'."
            }
        }

        $iconColorModeContracts = [ordered]@{
            # PQ-016 split geometry creation from state application; both must stay on the
            # shared factory so icon colors keep following the one audited policy.
            'oPenEfficiency/UI/Controls/FeatureIcon.xaml.cs' = 'FeatureIconPathFactory.ApplyState('
            'oPenEfficiency/UI/Panels/SidebarControlFactory.cs' = 'ColorMode = IconColorModes.Normalize(iconColorMode)'
            'oPenEfficiency/UI/Dialogs/FeatureExplorerWindow.xaml' = '<controls:FeatureIcon'
            'oPenEfficiency/UI/Dialogs/SettingsWindow.xaml' = 'SelectedItem="{Binding IconColorMode}"'
        }
        foreach ($entry in $iconColorModeContracts.GetEnumerator()) {
            $source = Get-Content -LiteralPath (Join-Path $Root $entry.Key) -Raw
            if (-not $source.Contains($entry.Value)) {
                Add-Finding $findings 'FeatureIconColorModeContract' $entry.Key 1 "missing:$($entry.Value)" `
                    'command-icon surface must use the shared icon color mode resolver.'
            }
        }

        $iconAccentContracts = [ordered]@{
            'oPenEfficiency.Core/UI/FeatureIconColorPolicy.cs' = 'FeatureIconAccentRole accentRole'
            'oPenEfficiency/Services/FeatureIconColorResolver.cs' = 'element.AccentRole'
            'oPenEfficiency/UI/ParsedIconGeometry.cs' = 'HasExplicitAccentRole'
            'oPenEfficiency/UI/Controls/FeatureIcon.xaml' = 'x:Name="IconLayerHost"'
            'oPenEfficiency/UI/Panels/SidebarControlFactory.cs' = 'FeatureIconSizing.ResolveGlyphSize(iconSize)'
        }
        foreach ($entry in $iconAccentContracts.GetEnumerator()) {
            $source = Get-Content -LiteralPath (Join-Path $Root $entry.Key) -Raw
            if (-not $source.Contains($entry.Value)) {
                Add-Finding $findings 'FeatureIconColorModeContract' $entry.Key 1 "missing:$($entry.Value)" `
                    'command-icon sizing and selective theme accents must stay on the shared policy path.'
            }
        }

        $policyPath = 'oPenEfficiency.Core/UI/FeatureIconColorPolicy.cs'
        $policySource = Get-Content -LiteralPath (Join-Path $Root $policyPath) -Raw
        foreach ($required in @(
            'SuccessGreen,',
            'SuccessGreenResourceKey = "SuccessGreenBrush"',
            'case FeatureIconAccentRole.SuccessGreen:'
        )) {
            if (-not $policySource.Contains($required)) {
                Add-Finding $findings 'FeatureIconColorModeContract' $policyPath 1 "missing:$required" `
                    'SuccessGreen must remain a first-class semantic role with its own theme resource.'
            }
        }
        foreach ($obsolete in @('FeatureIconAccentPolicy', 'ResolveForFeature(', 'useThemeAccent')) {
            if ($policySource.Contains($obsolete)) {
                Add-Finding $findings 'FeatureIconColorModeContract' $policyPath 1 $obsolete `
                    'whole-icon feature-ID accent policies are superseded by audited SVG element roles.'
            }
        }

        foreach ($themePath in @(
            'oPenEfficiency/UI/Styles/Themes/Light.xaml',
            'oPenEfficiency/UI/Styles/Themes/Dark.xaml'
        )) {
            $themeSource = Get-Content -LiteralPath (Join-Path $Root $themePath) -Raw
            foreach ($resourceKey in @('SuccessGreenColor', 'SuccessGreenBrush')) {
                if ($themeSource -notmatch ('x:Key\s*=\s*["'']' + [regex]::Escape($resourceKey) + '["'']')) {
                    Add-Finding $findings 'FeatureIconColorModeContract' $themePath 1 "missing:$resourceKey" `
                        'Light and Dark themes must each expose the dedicated SuccessGreen semantic resource.'
                }
            }
        }

        $featureIconPath = 'oPenEfficiency/UI/Controls/FeatureIcon.xaml.cs'
        $featureIconSource = Get-Content -LiteralPath (Join-Path $Root $featureIconPath) -Raw
        if (-not $featureIconSource.Contains('PropertyMetadata(IconColorModes.PowerPoint')) {
            Add-Finding $findings 'FeatureIconColorModeContract' $featureIconPath 1 'ColorMode default' `
                'FeatureIcon must fail safe to the PowerPoint icon color mode.'
        }

        $commandIconRoot = Join-Path $Root 'oPenEfficiency/UI/Assets/Icons'
        foreach ($iconFile in Get-ChildItem -LiteralPath $commandIconRoot -File -Filter '*.svg') {
            $source = Get-Content -LiteralPath $iconFile.FullName -Raw
            foreach ($match in [regex]::Matches(
                $source,
                '<[^>]+\bdata-oe-role\s*=\s*["'']accent["''][^>]*>',
                'IgnoreCase')) {
                if ($match.Value -notmatch '\bdata-oe-accent-role\s*=') {
                    $relativePath = Get-RelativePath $Root $iconFile.FullName
                    Add-Finding $findings 'FeatureIconColorModeContract' $relativePath `
                        (Get-LineNumber $source $match.Index) $match.Value `
                        'accent SVG elements must carry an explicit audited semantic role.'
                }
            }
        }

        $featureExplorerCodePath = 'oPenEfficiency/UI/Dialogs/FeatureExplorerWindow.xaml.cs'
        $featureExplorerCode = Get-Content -LiteralPath (Join-Path $Root $featureExplorerCodePath) -Raw
        if ($featureExplorerCode.Contains('SvgParser.ParseLucideSvg')) {
            Add-Finding $findings 'FeatureIconColorModeContract' $featureExplorerCodePath 1 'SvgParser.ParseLucideSvg' `
                'Feature Explorer must not maintain a second command-icon rendering path.'
        }

        $featureExplorerXamlPath = 'oPenEfficiency/UI/Dialogs/FeatureExplorerWindow.xaml'
        $featureExplorerXaml = Get-Content -LiteralPath (Join-Path $Root $featureExplorerXamlPath) -Raw
        $featureExplorerOverflowContracts = [ordered]@{
            'CategoryScrollViewer' = 'category navigation needs a named horizontally scrollable viewport.'
            'HorizontalScrollBarVisibility="Hidden"' = 'category navigation must hide standalone horizontal scrollbar chrome.'
            'CategoryScrollLeftButton' = 'category navigation needs an integrated earlier-categories action.'
            'CategoryScrollRightButton' = 'category navigation needs an integrated later-categories action.'
            'ToolTip="Show earlier categories"' = 'the earlier-categories icon action needs an action-oriented tooltip.'
            'ToolTip="Show later categories"' = 'the later-categories icon action needs an action-oriented tooltip.'
            'AutomationProperties.Name="Show earlier categories"' = 'the earlier-categories icon action needs an accessible name.'
            'AutomationProperties.Name="Show later categories"' = 'the later-categories icon action needs an accessible name.'
            'IsHitTestVisible="False" Data="M14,5 L7,12 L14,19"' = 'the earlier chevron glyph must remain decorative.'
            'IsHitTestVisible="False" Data="M10,5 L17,12 L10,19"' = 'the later chevron glyph must remain decorative.'
            '<Setter Property="ToolTip" Value="{Binding Name}"/>' = 'feature cards must expose the canonical name as a tooltip.'
            '<Setter Property="AutomationProperties.Name" Value="{Binding Name}"/>' = 'feature cards must expose the canonical name to assistive technology.'
            'TextWrapping="NoWrap" TextTrimming="CharacterEllipsis"' = 'feature-card descriptions must remain one-line scan-first summaries.'
            'ScrollViewer.VerticalScrollBarVisibility="Auto"' = 'the card ListBox must own its vertical scrolling.'
        }
        foreach ($entry in $featureExplorerOverflowContracts.GetEnumerator()) {
            if (-not $featureExplorerXaml.Contains($entry.Key)) {
                Add-Finding $findings 'FeatureExplorerDensityContract' $featureExplorerXamlPath 1 "missing:$($entry.Key)" $entry.Value
            }
        }
        if ($featureExplorerXaml.Contains('<ScrollViewer VerticalScrollBarVisibility="Auto" Margin="10,12,0,12">')) {
            Add-Finding $findings 'FeatureExplorerDensityContract' $featureExplorerXamlPath 1 'nested-card-scroll-viewer' `
                'Feature Explorer card content must not wrap the ListBox in a second vertical ScrollViewer.'
        }
        $featureExplorerPlannerPath = 'oPenEfficiency.Core/UI/FeatureExplorerNavigationPlanner.cs'
        $featureExplorerPlanner = Join-Path $Root $featureExplorerPlannerPath
        if (-not (Test-Path -LiteralPath $featureExplorerPlanner -PathType Leaf) -or
            -not $featureExplorerCode.Contains('FeatureExplorerNavigationPlanner.Create(') -or
            -not $featureExplorerCode.Contains('FeatureExplorerNavigationPlanner.CalculateRevealOffset(')) {
            Add-Finding $findings 'FeatureExplorerDensityContract' $featureExplorerCodePath 1 'missing:FeatureExplorerNavigationPlanner' `
                'Feature Explorer must delegate category overflow and reveal behavior to the tested host-neutral planner.'
        }

        $staleHeaderChecks = [ordered]@{
            'oPenEfficiency/UI/Dialogs/AutomationWizardWindow.xaml' = 'DOCUMENT AUTOMATION'
            'oPenEfficiency/UI/Dialogs/WizardDesignerWindow.xaml' = 'DOCUMENT AUTOMATION DESIGNER'
            'oPenEfficiency/UI/Dialogs/SeriesGeneratorWindow.xaml' = 'SERIES GENERATOR'
        }
        foreach ($entry in $staleHeaderChecks.GetEnumerator()) {
            $source = Get-Content -LiteralPath (Join-Path $Root $entry.Key) -Raw
            if ($source.Contains($entry.Value)) {
                Add-Finding $findings 'HeaderCopyPolicy' $entry.Key 1 $entry.Value `
                    "stale visible header '$($entry.Value)'."
            }
        }
    }

    return @($findings)
}

function Invoke-DetectorSelfTests {
    $fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ("oPenEfficiency-visual-consistency-" + [Guid]::NewGuid().ToString('N'))
    try {
        $stylesRoot = Join-Path $fixtureRoot 'oPenEfficiency/UI/Styles'
        $dialogsRoot = Join-Path $fixtureRoot 'oPenEfficiency/UI/Dialogs'
        $assetsRoot = Join-Path $fixtureRoot 'oPenEfficiency/Assets/Icons'
        [IO.Directory]::CreateDirectory($stylesRoot) | Out-Null
        [IO.Directory]::CreateDirectory($dialogsRoot) | Out-Null
        [IO.Directory]::CreateDirectory($assetsRoot) | Out-Null
        @'
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <Style x:Key="PrimaryButton" TargetType="Button"/>
  <Style x:Key="SecondaryButton" TargetType="Button"/>
  <Style x:Key="CompactTaskWindowStyle" TargetType="Window"/>
  <Style x:Key="CompactTaskShell" TargetType="Border"/>
</ResourceDictionary>
'@ | Set-Content -LiteralPath (Join-Path $stylesRoot 'Controls.xaml') -Encoding UTF8
        @'
<Window x:Class="Fixture.Good" xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Style="{DynamicResource CompactTaskWindowStyle}">
  <Window.Resources><ResourceDictionary><ResourceDictionary.MergedDictionaries><ResourceDictionary Source="../Styles/Controls.xaml"/></ResourceDictionary.MergedDictionaries></ResourceDictionary></Window.Resources>
  <Border Style="{DynamicResource CompactTaskShell}" Background="{DynamicResource SurfaceBrush}"/>
</Window>
'@ | Set-Content -LiteralPath (Join-Path $dialogsRoot 'GoodWindow.xaml') -Encoding UTF8
        @'
<Window x:Class="Fixture.Bad" xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Background="#123456" FontFamily="Arial">
  <Window.Resources><Style x:Key="PrimaryButton" TargetType="Button"><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border/></ControlTemplate></Setter.Value></Setter></Style></Window.Resources>
  <Border><Border.Effect><DropShadowEffect BlurRadius="20"/></Border.Effect></Border>
</Window>
'@ | Set-Content -LiteralPath (Join-Path $dialogsRoot 'BadWindow.xaml') -Encoding UTF8
        @'
<UserControl xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
  <!-- VisualColorException: test content swatch -->
  <Border Background="#ABCDEF"/>
</UserControl>
'@ | Set-Content -LiteralPath (Join-Path $dialogsRoot 'ContentPreview.xaml') -Encoding UTF8
        '<svg xmlns="http://www.w3.org/2000/svg"/>' |
            Set-Content -LiteralPath (Join-Path $assetsRoot 'sample_hover.svg') -Encoding UTF8

        $findings = @(Get-VisualFindings $fixtureRoot $false)
        $expectedRules = @(
            'HardCodedChromeColor',
            'DuplicateCanonicalStyle',
            'UnauthorizedFontFamily',
            'MissingSharedVisualResources',
            'InvalidArchetypeWindowContract',
            'AdHocWindowShadow',
            'LocalCanonicalButtonTemplate',
            'AccidentalSecondIconStateSystem'
        )
        foreach ($rule in $expectedRules) {
            if (-not ($findings | Where-Object { $_.Rule -eq $rule })) {
                throw "Visual consistency detector self-test did not trigger '$rule'."
            }
        }
        if ($findings | Where-Object {
                $_.Rule -eq 'HardCodedChromeColor' -and $_.Path -eq 'oPenEfficiency/UI/Dialogs/ContentPreview.xaml'
            }) {
            throw 'VisualColorException marker did not suppress the adjacent content-color fixture.'
        }

        $settingsPath = Join-Path $RepositoryRoot 'oPenEfficiency/UI/Dialogs/SettingsWindow.xaml'
        $settingsFixture = Get-Content -LiteralPath $settingsPath -Raw
        $settingsFindings = [Collections.Generic.List[object]]::new()
        Add-SettingsXamlContractFindings $settingsFindings $settingsFixture 'SettingsWindow.xaml'
        if ($settingsFindings.Count -ne 0) {
            throw "Settings golden-reference fixture unexpectedly produced $($settingsFindings.Count) finding(s)."
        }

        $negativeFixtures = [ordered]@{
            'SwatchInheritance' = $settingsFixture.Replace('BasedOn="{StaticResource CompactIconButton}"', 'BasedOn="{StaticResource SecondaryButton}"')
            'SwatchSize' = $settingsFixture.Replace('<Setter Property="Width"  Value="{StaticResource IconButtonSizeCompact}"/>', '<Setter Property="Width" Value="24"/>')
        }
        foreach ($swatchName in @('Apply gold color', 'Apply orange color', 'Apply blue color', 'Apply teal color')) {
            $negativeFixtures["SwatchName:$swatchName"] = $settingsFixture.Replace(
                ('AutomationProperties.Name="{0}"' -f $swatchName),
                'AutomationProperties.Name="Duplicate swatch"')
            $negativeFixtures["SwatchTooltip:$swatchName"] = $settingsFixture.Replace(
                ('ToolTip="{0}"' -f $swatchName),
                'ToolTip="Duplicate swatch"')
        }
        foreach ($contract in $script:SettingsAccessibleControlContracts) {
            $accessibleName = 'AutomationProperties.Name="{0}"' -f $contract.Name
            $missingAccessibleName = 'AutomationProperties.Name="Missing {0}"' -f $contract.Name
            $negativeFixtures["Accessible:$($contract.Name)"] = $settingsFixture.Replace(
                $accessibleName,
                $missingAccessibleName)
        }
        foreach ($kind in @('Color', 'Font')) {
            $negativeFixtures["TokenRemoveParameter:$kind"] = $settingsFixture.Replace(
                "DataContext.RemoveStyleCheck${kind}Command, RelativeSource={RelativeSource AncestorType=Window}}`" CommandParameter=`"{Binding}",
                "DataContext.RemoveStyleCheck${kind}Command, RelativeSource={RelativeSource AncestorType=Window}}`" CommandParameter=`"missing-token")
            $tokenName = "Remove approved $($kind.ToLowerInvariant()) {0}"
            $negativeFixtures["TokenRemoveName:$kind"] = $settingsFixture.Replace($tokenName, "Remove $($kind.ToLowerInvariant())")
        }
        foreach ($entry in $negativeFixtures.GetEnumerator()) {
            $negativeFindings = [Collections.Generic.List[object]]::new()
            Add-SettingsXamlContractFindings $negativeFindings $entry.Value 'SettingsWindow.xaml'
            if (-not ($negativeFindings | Where-Object { $_.Rule -eq 'GoldenReferenceContract' })) {
                throw "Settings golden-reference detector self-test did not reject '$($entry.Key)'."
            }
        }

        $lightFixture = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'oPenEfficiency/UI/Styles/Themes/Light.xaml') -Raw
        $darkFixture = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'oPenEfficiency/UI/Styles/Themes/Dark.xaml') -Raw
        $controlsFixture = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'oPenEfficiency/UI/Styles/Controls.xaml') -Raw
        $settingsComboBoxFindings = [Collections.Generic.List[object]]::new()
        Add-SettingsComboBoxContractFindings $settingsComboBoxFindings $settingsFixture $controlsFixture
        if ($settingsComboBoxFindings.Count -ne 0) {
            throw "Settings ComboBox golden-reference fixture unexpectedly produced $($settingsComboBoxFindings.Count) finding(s)."
        }
        foreach ($fixture in @(
            @{ Name = 'StyleInheritance'; Settings = $settingsFixture.Replace('BasedOn="{StaticResource {x:Type ComboBox}}"', 'BasedOn="{StaticResource SettingsFieldControl}"'); Controls = $controlsFixture },
            @{ Name = 'BypassedStyle'; Settings = $settingsFixture.Replace('Style="{StaticResource SettingsComboBox}"', 'Style="{StaticResource SettingsFieldControl}"'); Controls = $controlsFixture },
            @{ Name = 'SelectedSurface'; Settings = $settingsFixture; Controls = $controlsFixture.Replace('TargetName="Bd" Property="Background" Value="{DynamicResource SurfaceSelectedBrush}"', 'TargetName="Bd" Property="Background" Value="{DynamicResource AccentBrush}"') },
            @{ Name = 'KeyboardFocus'; Settings = $settingsFixture; Controls = $controlsFixture.Replace('x:Name="FocusBorder"', 'x:Name="MissingFocusBorder"') },
            @{ Name = 'DisabledState'; Settings = $settingsFixture; Controls = $controlsFixture.Replace('<Trigger Property="IsEnabled" Value="False">', '<Trigger Property="MissingEnabled" Value="False">') }
        )) {
            $negativeFindings = [Collections.Generic.List[object]]::new()
            Add-SettingsComboBoxContractFindings $negativeFindings $fixture.Settings $fixture.Controls
            if (-not ($negativeFindings | Where-Object { $_.Rule -eq 'GoldenReferenceContract' })) {
                throw "Settings ComboBox detector self-test did not reject '$($fixture.Name)'."
            }
        }
        $settingsTextBoxFindings = [Collections.Generic.List[object]]::new()
        Add-SettingsTextBoxContractFindings $settingsTextBoxFindings $settingsFixture $controlsFixture
        if ($settingsTextBoxFindings.Count -ne 0) {
            throw "Settings TextBox golden-reference fixture unexpectedly produced $($settingsTextBoxFindings.Count) finding(s)."
        }
        foreach ($fixture in @(
            @{ Name = 'StyleInheritance'; Settings = $settingsFixture.Replace('BasedOn="{StaticResource {x:Type TextBox}}"', 'BasedOn="{StaticResource SettingsFieldControl}"'); Controls = $controlsFixture },
            @{ Name = 'BypassedStyle'; Settings = $settingsFixture.Replace('Style="{StaticResource SettingsTextBox}"', 'Style="{StaticResource SettingsFieldControl}"'); Controls = $controlsFixture },
            @{ Name = 'ThemeSurface'; Settings = $settingsFixture; Controls = $controlsFixture.Replace('Property="Background" Value="{DynamicResource SurfaceBrush}"', 'Property="Background" Value="White"') },
            @{ Name = 'KeyboardFocus'; Settings = $settingsFixture; Controls = $controlsFixture.Replace('Property="IsFocused" Value="True"', 'Property="MissingFocused" Value="True"') },
            @{ Name = 'HoverState'; Settings = $settingsFixture; Controls = $controlsFixture.Replace('Property="IsMouseOver" Value="True"', 'Property="MissingMouseOver" Value="True"') }
        )) {
            $negativeFindings = [Collections.Generic.List[object]]::new()
            Add-SettingsTextBoxContractFindings $negativeFindings $fixture.Settings $fixture.Controls
            if (-not ($negativeFindings | Where-Object { $_.Rule -eq 'GoldenReferenceContract' })) {
                throw "Settings TextBox detector self-test did not reject '$($fixture.Name)'."
            }
        }
        $sidebarFixture = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'oPenEfficiency/UI/Panels/MainSidebar.xaml.cs') -Raw
        $viewModelFixture = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'oPenEfficiency/UI/SettingsViewModel.cs') -Raw
        $settingsCodeFixture = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'oPenEfficiency/UI/Dialogs/SettingsWindow.xaml.cs') -Raw
        $migrationTestsFixture = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'oPenEfficiency.Core.Tests/SidebarMigrationBoundaryTests.cs') -Raw
        $stickerPreviewModelFixture = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'oPenEfficiency.Core/Models/StickerPreviewModel.cs') -Raw
        $statusStickerFeatureFixture = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'oPenEfficiency/Features/Visuals/IllustrativeStickerFeature.cs') -Raw
        $stickyNoteFeatureFixture = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'oPenEfficiency/Features/Utilities/StickyNoteFeature.cs') -Raw
        $stickerPreviewTestsFixture = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'oPenEfficiency.Core.Tests/StickerPreviewModelTests.cs') -Raw
        $stickerPreviewFindings = [Collections.Generic.List[object]]::new()
        Add-StickerPreviewContractFindings $stickerPreviewFindings $settingsFixture $viewModelFixture `
            $stickerPreviewModelFixture $statusStickerFeatureFixture $stickyNoteFeatureFixture $stickerPreviewTestsFixture
        if ($stickerPreviewFindings.Count -ne 0) {
            throw "Sticker preview golden-reference fixture unexpectedly produced $($stickerPreviewFindings.Count) finding(s)."
        }
        foreach ($fixture in @(
            @{ Name = 'StickyPreviewUi'; Settings = $settingsFixture.Replace('AutomationProperties.Name="Sticky Note Preview"', 'AutomationProperties.Name="Missing Sticky Preview"'); ViewModel = $viewModelFixture; Model = $stickerPreviewModelFixture; StatusFeature = $statusStickerFeatureFixture; StickyFeature = $stickyNoteFeatureFixture; Tests = $stickerPreviewTestsFixture },
            @{ Name = 'GreenBShadow'; Settings = $settingsFixture.Replace('Width="310" Height="54" Background="#000000"', 'Width="310" Height="54" Background="Transparent"'); ViewModel = $viewModelFixture; Model = $stickerPreviewModelFixture; StatusFeature = $statusStickerFeatureFixture; StickyFeature = $stickyNoteFeatureFixture; Tests = $stickerPreviewTestsFixture },
            @{ Name = 'StatusPreviewUi'; Settings = $settingsFixture.Replace('Text="{Binding StatusStickerPreviewText}"', 'Text="Missing preview"'); ViewModel = $viewModelFixture; Model = $stickerPreviewModelFixture; StatusFeature = $statusStickerFeatureFixture; StickyFeature = $stickyNoteFeatureFixture; Tests = $stickerPreviewTestsFixture },
            @{ Name = 'LiveNotifications'; Settings = $settingsFixture; ViewModel = $viewModelFixture.Replace('private void NotifyStickerPreviewChanged()', 'private void MissingStickerPreviewNotifications()'); Model = $stickerPreviewModelFixture; StatusFeature = $statusStickerFeatureFixture; StickyFeature = $stickyNoteFeatureFixture; Tests = $stickerPreviewTestsFixture },
            @{ Name = 'HostNeutralModel'; Settings = $settingsFixture; ViewModel = $viewModelFixture; Model = $stickerPreviewModelFixture.Replace('using System;', 'using Microsoft.Office.Interop.PowerPoint;'); StatusFeature = $statusStickerFeatureFixture; StickyFeature = $stickyNoteFeatureFixture; Tests = $stickerPreviewTestsFixture },
            @{ Name = 'SharedLabelFallback'; Settings = $settingsFixture; ViewModel = $viewModelFixture; Model = $stickerPreviewModelFixture; StatusFeature = $statusStickerFeatureFixture.Replace('StickerPreviewModel.GetFirstStatusLabel', 'SplitLabels'); StickyFeature = $stickyNoteFeatureFixture; Tests = $stickerPreviewTestsFixture },
            @{ Name = 'SharedStickyPalette'; Settings = $settingsFixture; ViewModel = $viewModelFixture; Model = $stickerPreviewModelFixture; StatusFeature = $statusStickerFeatureFixture; StickyFeature = $stickyNoteFeatureFixture.Replace('StickerPreviewModel.CreateStickyNote(config, null, "GreenB")', 'new StickyNotePreviewState()'); Tests = $stickerPreviewTestsFixture },
            @{ Name = 'FallbackTests'; Settings = $settingsFixture; ViewModel = $viewModelFixture; Model = $stickerPreviewModelFixture; StatusFeature = $statusStickerFeatureFixture; StickyFeature = $stickyNoteFeatureFixture; Tests = $stickerPreviewTestsFixture.Replace('EmptySettingsProduceSafeRepresentativePreviews', 'MissingSafePreviewTest') }
        )) {
            $negativeFindings = [Collections.Generic.List[object]]::new()
            Add-StickerPreviewContractFindings $negativeFindings $fixture.Settings $fixture.ViewModel `
                $fixture.Model $fixture.StatusFeature $fixture.StickyFeature $fixture.Tests
            if (-not ($negativeFindings | Where-Object { $_.Rule -eq 'GoldenReferenceContract' })) {
                throw "Sticker preview detector self-test did not reject '$($fixture.Name)'."
            }
        }
        $sidebarThemeFindings = [Collections.Generic.List[object]]::new()
        Add-SidebarThemeContractFindings $sidebarThemeFindings $lightFixture $darkFixture $controlsFixture `
            $sidebarFixture $viewModelFixture $settingsFixture $settingsCodeFixture $migrationTestsFixture
        if ($sidebarThemeFindings.Count -ne 0) {
            throw "Sidebar theme golden-reference fixture unexpectedly produced $($sidebarThemeFindings.Count) finding(s)."
        }

        $sidebarThemeNegativeFixtures = @(
            @{ Name = 'LightHeader'; Light = $lightFixture.Replace('#000000', '#010101'); Dark = $darkFixture; Controls = $controlsFixture; Sidebar = $sidebarFixture; ViewModel = $viewModelFixture; Settings = $settingsFixture },
            @{ Name = 'DarkDivider'; Light = $lightFixture; Dark = $darkFixture.Replace('#656565', '#646464'); Controls = $controlsFixture; Sidebar = $sidebarFixture; ViewModel = $viewModelFixture; Settings = $settingsFixture },
            @{ Name = 'SectionStyle'; Light = $lightFixture; Dark = $darkFixture; Controls = $controlsFixture.Replace('{DynamicResource SidebarSectionHeaderBrush}', '{DynamicResource TextSecondaryBrush}'); Sidebar = $sidebarFixture; ViewModel = $viewModelFixture; Settings = $settingsFixture },
            @{ Name = 'SidebarDivider'; Light = $lightFixture; Dark = $darkFixture; Controls = $controlsFixture; Sidebar = $sidebarFixture.Replace('"SidebarDividerBrush"', '"BorderBrush"'); ViewModel = $viewModelFixture; Settings = $settingsFixture },
            @{ Name = 'SidebarDividerWeight'; Light = $lightFixture; Dark = $darkFixture; Controls = $controlsFixture; Sidebar = $sidebarFixture.Replace('new Thickness(0, 0, 0.5, 0)', 'new Thickness(0, 0, 1, 0)'); ViewModel = $viewModelFixture; Settings = $settingsFixture },
            @{ Name = 'SidebarHeaderPath'; Light = $lightFixture; Dark = $darkFixture; Controls = $controlsFixture; Sidebar = $sidebarFixture.Replace('FindResource("SidebarSectionHeader")', 'FindResource("SectionHeader")'); ViewModel = $viewModelFixture; Settings = $settingsFixture },
            @{ Name = 'ViewModelMode'; Light = $lightFixture; Dark = $darkFixture; Controls = $controlsFixture; Sidebar = $sidebarFixture; ViewModel = $viewModelFixture.Replace('public bool UseThemeSectionFontColor', 'public bool MissingThemeSectionFontColor'); Settings = $settingsFixture },
            @{ Name = 'ViewModelReset'; Light = $lightFixture; Dark = $darkFixture; Controls = $controlsFixture; Sidebar = $sidebarFixture; ViewModel = $viewModelFixture.Replace('UseThemeSectionFontColor = true;', 'UseThemeSectionFontColor = false;'); Settings = $settingsFixture },
            @{ Name = 'ViewModelExport'; Light = $lightFixture; Dark = $darkFixture; Controls = $controlsFixture; Sidebar = $sidebarFixture; ViewModel = $viewModelFixture.Replace('exportData.SectionFontColorMode = Config.SectionFontColorMode;', 'exportData.SectionFontColorMode = null;'); Settings = $settingsFixture },
            @{ Name = 'ViewModelImport'; Light = $lightFixture; Dark = $darkFixture; Controls = $controlsFixture; Sidebar = $sidebarFixture; ViewModel = $viewModelFixture.Replace('Config.SectionFontColorMode = imported.SectionFontColorMode;', 'Config.SectionFontColorMode = null;'); Settings = $settingsFixture },
            @{ Name = 'SettingsMode'; Light = $lightFixture; Dark = $darkFixture; Controls = $controlsFixture; Sidebar = $sidebarFixture; ViewModel = $viewModelFixture; Settings = $settingsFixture.Replace('IsChecked="{Binding UseThemeSectionFontColor}"', 'IsChecked="False"') },
            @{ Name = 'SettingsPreview'; Light = $lightFixture; Dark = $darkFixture; Controls = $controlsFixture; Sidebar = $sidebarFixture; ViewModel = $viewModelFixture; Settings = $settingsFixture.Replace('Fill="{DynamicResource SidebarDividerBrush}"', 'Fill="{DynamicResource BorderLightBrush}"') }
            @{ Name = 'SettingsPreviewDividerWeight'; Light = $lightFixture; Dark = $darkFixture; Controls = $controlsFixture; Sidebar = $sidebarFixture; ViewModel = $viewModelFixture; Settings = $settingsFixture.Replace('Height="0.75" Fill="{DynamicResource SidebarDividerBrush}"', 'Height="1.5" Fill="{DynamicResource SidebarDividerBrush}"') }
        )
        foreach ($fixture in $sidebarThemeNegativeFixtures) {
            $negativeFindings = [Collections.Generic.List[object]]::new()
            Add-SidebarThemeContractFindings $negativeFindings $fixture.Light $fixture.Dark $fixture.Controls `
                $fixture.Sidebar $fixture.ViewModel $fixture.Settings $settingsCodeFixture $migrationTestsFixture
            if (-not ($negativeFindings | Where-Object { $_.Rule -eq 'GoldenReferenceContract' })) {
                throw "Sidebar theme golden-reference detector self-test did not reject '$($fixture.Name)'."
            }
        }
        foreach ($fixture in @(
            @{ Name = 'SettingsContrast'; Code = $settingsCodeFixture.Replace('ResolveContrastColor(', 'ResolveStoredColor('); Tests = $migrationTestsFixture },
            @{ Name = 'MigrationCoverage'; Code = $settingsCodeFixture; Tests = $migrationTestsFixture.Replace('LegacyCustomSectionHeaderIsPreservedAsCustom', 'MissingLegacyCustomSectionHeaderTest') }
        )) {
            $negativeFindings = [Collections.Generic.List[object]]::new()
            Add-SidebarThemeContractFindings $negativeFindings $lightFixture $darkFixture $controlsFixture `
                $sidebarFixture $viewModelFixture $settingsFixture $fixture.Code $fixture.Tests
            if (-not ($negativeFindings | Where-Object { $_.Rule -eq 'GoldenReferenceContract' })) {
                throw "Sidebar theme golden-reference detector self-test did not reject '$($fixture.Name)'."
            }
        }
    }
    finally {
        if (Test-Path -LiteralPath $fixtureRoot) {
            Remove-Item -LiteralPath $fixtureRoot -Recurse -Force
        }
    }
}

if (-not $SkipSelfTest) {
    Invoke-DetectorSelfTests
}

$findings = @(Get-VisualFindings $RepositoryRoot $true)
$windowCount = @($findings | Out-Null; Get-ChildItem -LiteralPath (Join-Path $RepositoryRoot 'oPenEfficiency/UI') -Filter '*.xaml' -File -Recurse |
    Where-Object { (Get-Content -LiteralPath $_.FullName -Raw) -match '(?m)^\s*<Window(?:\s|>)' }).Count

if ($Audit) {
    $findings | Select-Object Rule, Path, Line, Signature, Fingerprint, Message | ConvertTo-Json -Depth 4
    return
}

if (-not (Test-Path -LiteralPath $BaselinePath -PathType Leaf)) {
    throw "Visual consistency baseline was not found at '$BaselinePath'."
}
$baseline = Get-Content -LiteralPath $BaselinePath -Raw | ConvertFrom-Json
if ($baseline.SchemaVersion -ne 1) {
    throw "Unsupported visual consistency baseline schema '$($baseline.SchemaVersion)'."
}

$allowedCounts = @{}
foreach ($entry in @($baseline.LegitimateExceptions) + @($baseline.LegacyViolations)) {
    if ([string]::IsNullOrWhiteSpace($entry.Fingerprint)) {
        throw "Visual consistency baseline contains an entry without a fingerprint."
    }
    $count = if ($null -eq $entry.Count) { 1 } else { [int]$entry.Count }
    if ($allowedCounts.ContainsKey($entry.Fingerprint)) {
        $allowedCounts[$entry.Fingerprint] += $count
    }
    else {
        $allowedCounts[$entry.Fingerprint] = $count
    }
}
foreach ($entry in @($baseline.LegitimateExceptions)) {
    if ([string]::IsNullOrWhiteSpace($entry.Reason)) {
        throw "Legitimate visual exception '$($entry.Rule)' in '$($entry.Path)' needs a reason."
    }
}

$errors = [Collections.Generic.List[string]]::new()
foreach ($group in @($findings | Group-Object Fingerprint)) {
    $allowed = if ($allowedCounts.ContainsKey($group.Name)) { $allowedCounts[$group.Name] } else { 0 }
    if ($group.Count -le $allowed) { continue }
    foreach ($finding in @($group.Group | Select-Object -First ($group.Count - $allowed))) {
        $errors.Add("$($finding.Path):$($finding.Line) [$($finding.Rule)] $($finding.Message)") | Out-Null
    }
}

if ($errors.Count -gt 0) {
    throw "Visual-consistency validation failed:`n$($errors -join "`n")"
}

Write-Output (
    'VISUAL_CONSISTENCY_PASS windows={0} findings={1} legacy={2} exceptions={3} detector_self_tests={4}' -f
    $windowCount,
    $findings.Count,
    @($baseline.LegacyViolations).Count,
    @($baseline.LegitimateExceptions).Count,
    (-not $SkipSelfTest).ToString().ToLowerInvariant())
