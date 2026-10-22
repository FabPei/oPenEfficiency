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
function Assert-Sync([bool]$Condition, [string]$Message) { if (-not $Condition) { $failures.Add($Message) } }
function Pump-Sync($window) { $window.Dispatcher.Invoke([Action]{}, [Windows.Threading.DispatcherPriority]::ApplicationIdle) }
function Find-Visual($root2, [string]$typeName) {
    $result = [Collections.Generic.List[object]]::new()
    $queue = [Collections.Generic.Queue[object]]::new()
    $queue.Enqueue($root2)
    while ($queue.Count -gt 0) {
        $node = $queue.Dequeue()
        if ($node.GetType().Name -eq $typeName) { $result.Add($node) }
        if ($node -is [Windows.Media.Visual]) {
            $count = [Windows.Media.VisualTreeHelper]::GetChildrenCount($node)
            for ($i = 0; $i -lt $count; $i++) { $queue.Enqueue([Windows.Media.VisualTreeHelper]::GetChild($node, $i)) }
        }
    }
    return $result
}
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
    $window = [oPenEfficiency.UI.SyncObjectsWindow]::new($null)
    $window.ShowInTaskbar = $false
    $window.WindowStartupLocation = 'Manual'
    $window.Left = -1000
    $window.Top = -1000
    $window.Show()
    Pump-Sync $window

    # REG-SYNC-WINDOW-UI-001: exactly one resize grip (the workspace template's corner grip).
    # The former explicit footer ResizeGrip duplicated the triangle next to the Close button.
    $grips = @(Find-Visual $window 'ResizeGrip')
    $observed.ResizeGripCount = $grips.Count
    Assert-Sync ($grips.Count -eq 1) "Expected exactly one resize grip, found $($grips.Count)."

    # Native non-client routing: the full measured header is a caption; controls stay client.
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class SyncWindowCaptionProbe {
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
    public static int Hit(IntPtr window, int x, int y) {
        if (x < -32768 || x > 32767 || y < -32768 || y > 32767)
            throw new ArgumentOutOfRangeException("Point must fit WM_NCHITTEST screen coordinates.");
        return SendMessage(window,0x84,IntPtr.Zero,new IntPtr(unchecked((y & 0xffff) << 16 | (x & 0xffff)))).ToInt32();
    }
}
"@
    $handle = [Windows.Interop.WindowInteropHelper]::new($window).Handle
    $shell = $window.Content
    $grid = $shell.Child
    $header = $grid.Children[0]
    $close = $header.Children[1]
    $body = $window.FindName('BtnDistribute')
    $targets = [ordered]@{
        Title = $header.TranslatePoint([Windows.Point]::new(200, $header.ActualHeight / 2), $window)
        TopPadding = [Windows.Point]::new(200, 10)
        Close = $close.TranslatePoint([Windows.Point]::new($close.ActualWidth / 2, $close.ActualHeight / 2), $window)
        Body = $body.TranslatePoint([Windows.Point]::new($body.ActualWidth / 2, $body.ActualHeight / 2), $window)
    }
    Pump-Sync $window
    $observed.CaptionHitTests = [ordered]@{}
    foreach ($target in $targets.GetEnumerator()) {
        $point = $window.PointToScreen($target.Value)
        $hit = [SyncWindowCaptionProbe]::Hit($handle, [int]$point.X, [int]$point.Y)
        $observed.CaptionHitTests[$target.Key] = $hit
        $expected = if ($target.Key -in @('Title', 'TopPadding')) { 2 } else { 1 }
        Assert-Sync ($hit -eq $expected) "Native caption routing for $($target.Key) is $hit, expected $expected."
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
if ($failures.Count -gt 0) { throw ('Sync Objects window regression: ' + ($failures -join '; ')) }
Write-Output ("SYNC_OBJECTS_WINDOW_PASS resize_grips={0} header_drag=owner_retake" -f $observed.ResizeGripCount)
$global:LASTEXITCODE = 0
