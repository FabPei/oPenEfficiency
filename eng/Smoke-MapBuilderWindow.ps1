[CmdletBinding()]
param([string]$ResultsPath)

$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') { throw 'Run this WPF smoke on an STA thread.' }
$root = Split-Path -Parent $PSScriptRoot
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -Path "$env:USERPROFILE/.nuget/packages/microsoftofficecore/15.0.0/lib/net35/Office.dll"
Add-Type -Path "$env:USERPROFILE/.nuget/packages/microsoft.office.interop.powerpoint/15.0.4420.1018/lib/net20/Microsoft.Office.Interop.PowerPoint.dll"
Add-Type -Path "$root/oPenEfficiency.Core/bin/Release/oPenEfficiency.Core.dll"
Add-Type -Path "$root/oPenEfficiency/bin/Release/oPenEfficiency.dll"
$observed = [ordered]@{}
$failures = [Collections.Generic.List[string]]::new()
function Assert-Map([bool]$Condition, [string]$Message) { if (-not $Condition) { $failures.Add($Message) } }
function Pump-Map($window) { $window.Dispatcher.Invoke([Action]{}, [Windows.Threading.DispatcherPriority]::ApplicationIdle) }
$window = $null
try {
    if ($null -eq [Windows.Application]::Current) {
        $application = [Windows.Application]::new()
        $application.ShutdownMode = 'OnExplicitShutdown'
    }
    $theme = [Windows.ResourceDictionary]::new()
    $theme.Source = [Uri]::new('pack://application:,,,/oPenEfficiency;component/UI/Styles/Themes/Dark.xaml')
    [Windows.Application]::Current.Resources.MergedDictionaries.Add($theme)
    $observed.Theme = 'Dark'
    $window = [oPenEfficiency.UI.Dialogs.MapWizardWindow]::new($null, $null, $null)
    $window.ShowInTaskbar = $false
    $window.WindowStartupLocation = 'Manual'
    $window.Left = -1000
    $window.Top = -1000
    $window.Show()
    Pump-Map $window
    $scope = $window.FindName('CmbMapScope')
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $scope.SelectedIndex = 2
    $window.UpdateLayout()
    $watch.Stop()
    $observed.ColdCountrySelectionMilliseconds = [math]::Round($watch.Elapsed.TotalMilliseconds, 1)
    Assert-Map ($watch.Elapsed.TotalSeconds -lt 2) 'First country selection blocks the dispatcher for at least two seconds.'

    # Exercise both parent lists as well as the availability counts, without inserting maps.
    $observed.GermanyAdm2Parents = [oPenEfficiency.Services.MapAssetService]::GetLowerLevelParentRegionCatalog('DEU', 'ADM2').Count
    $observed.GermanyAdm3Parents = [oPenEfficiency.Services.MapAssetService]::GetLowerLevelParentRegionCatalog('DEU', 'ADM3').Count
    $cacheField = [oPenEfficiency.Services.MapAssetService].GetField('AssetCache', [Reflection.BindingFlags]'Static,NonPublic')
    $geometryKeys = @($cacheField.GetValue($null).Keys | Where-Object { $_ -match '(administrative_regions_adm[23]|states_provinces_admin1).*json.gz' })
    $observed.AdministrativeGeometryCacheKeys = $geometryKeys
    Assert-Map ($geometryKeys.Count -eq 0) 'Selection/parent queries loaded administrative polygon geometry.'

    $body = @($window.Content.Child.Children | Where-Object { $_ -is [Windows.Controls.ScrollViewer] })[0]
    $observed.Layouts = @()
    foreach ($size in @(@(640, 480), @(760, 720), @(960, 760))) {
        $window.Width = $size[0]
        $window.Height = $size[1]
        $body.ScrollToTop()
        $window.UpdateLayout()
        Pump-Map $window
        $bounds = @()
        foreach ($name in @('CmbMapScope', 'CmbCountry', 'DgDataPreview', 'TxtMapStatus', 'BtnInsertMap', 'TxtLowColor', 'BtnApply')) {
            $item = $window.FindName($name)
            $point = $item.TranslatePoint([Windows.Point]::new(0, 0), $window)
            $bounds += [ordered]@{ Name = $name; Left = $point.X; Right = $point.X + $item.ActualWidth }
            Assert-Map ($point.X -ge 0 -and ($point.X + $item.ActualWidth) -le ($window.ActualWidth + 1)) "$name extends beyond window width $($window.ActualWidth)."
        }
        Assert-Map ($body.ExtentWidth -le ($body.ViewportWidth + 1)) 'The outer workspace requires horizontal scrolling.'
        $body.ScrollToBottom()
        Pump-Map $window
        $observed.Layouts += [ordered]@{ Width = $window.ActualWidth; Height = $window.ActualHeight; ExtentWidth = $body.ExtentWidth; ViewportWidth = $body.ViewportWidth; Controls = $bounds }
    }
    $observed.DpiScale = [Windows.Media.VisualTreeHelper]::GetDpi($window).DpiScaleX
    # Query native non-client routing only; no pointer/key input is injected.
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class MapWindowCaptionProbe {
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
    public static int Hit(IntPtr window, int x, int y) {
        if (x < -32768 || x > 32767 || y < -32768 || y > 32767)
            throw new ArgumentOutOfRangeException("Point must fit WM_NCHITTEST screen coordinates.");
        return SendMessage(window,0x84,IntPtr.Zero,new IntPtr(unchecked((y & 0xffff) << 16 | (x & 0xffff)))).ToInt32();
    }
}
"@
    $handle = [Windows.Interop.WindowInteropHelper]::new($window).Handle
    $header = $window.Content.Child.Children[0]
    $close = $header.Children[1]
    $targets = [ordered]@{
        Title = $header.TranslatePoint([Windows.Point]::new(200, $header.ActualHeight / 2), $window)
        TopPadding = [Windows.Point]::new(200, 10)
        Close = $close.TranslatePoint([Windows.Point]::new($close.ActualWidth / 2, $close.ActualHeight / 2), $window)
        Body = $scope.TranslatePoint([Windows.Point]::new($scope.ActualWidth / 2, $scope.ActualHeight / 2), $window)
    }
    $body.ScrollToTop()
    Pump-Map $window
    # Recalculate the body point after resetting the scroll position.
    $targets.Body = $scope.TranslatePoint([Windows.Point]::new($scope.ActualWidth / 2, $scope.ActualHeight / 2), $window)
    $observed.CaptionHitTests = [ordered]@{}
    foreach ($target in $targets.GetEnumerator()) {
        $point = $window.PointToScreen($target.Value)
        $hit = [MapWindowCaptionProbe]::Hit($handle, [int]$point.X, [int]$point.Y)
        $observed.CaptionHitTests[$target.Key] = $hit
        $expected = if ($target.Key -in @('Title', 'TopPadding')) { 2 } else { 1 }
        Assert-Map ($hit -eq $expected) "Native caption routing for $($target.Key) is $hit, expected $expected."
    }
    $observed.HeaderDrag = 'OWNER_RETAKE_REQUIRED; native movement not verified'
} catch {
    $failures.Add($_.Exception.ToString())
} finally {
    if ($null -ne $window) { $window.Close() }
    if (-not [string]::IsNullOrWhiteSpace($ResultsPath)) {
        $path = [IO.Path]::GetFullPath($ResultsPath)
        [IO.Directory]::CreateDirectory((Split-Path -Parent $path)) | Out-Null
        [ordered]@{ SchemaVersion = 1; Passed = ($failures.Count -eq 0); Observed = $observed; Failures = @($failures.ToArray()) } |
            ConvertTo-Json -Depth 9 | Set-Content -LiteralPath $path -Encoding UTF8
    }
}
if ($failures.Count -gt 0) { throw ('Map Builder regression: ' + ($failures -join '; ')) }
Write-Output ("MAP_BUILDER_WINDOW_PASS cold_selection_ms={0} layouts=3 geometry_cache_entries=0 header_drag=owner_retake" -f $observed.ColdCountrySelectionMilliseconds)
$global:LASTEXITCODE = 0
