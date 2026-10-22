[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$violations = [System.Collections.Generic.List[string]]::new()

function Get-SourceFiles([string[]]$path) {
    Get-ChildItem -LiteralPath $path -Filter "*.cs" -File -Recurse |
        Where-Object { $_.FullName -notmatch '\\(bin|obj)\\' }
}

# Core must stay usable by every host: no Office interop and no desktop UI types.
$coreRoot = Join-Path $root "oPenEfficiency.Core"
foreach ($file in Get-SourceFiles $coreRoot) {
    $text = [System.IO.File]::ReadAllText($file.FullName)
    if ($text -match '\bMicrosoft\.Office\b|\bOffice\s*=\s*Microsoft\.Office|\bSystem\.Windows\b') {
        $violations.Add("Core host/UI dependency: $($file.FullName.Substring($root.Length + 1))")
    }
}

# Existing host-bound accesses are capped per file. A new path or an increase fails.
$globalAccessAllowlist = @{
    "oPenEfficiency.Excel/UI/ExcelRibbon.cs" = 1
    "oPenEfficiency.Excel/ThisAddIn.Designer.cs" = 1
    "oPenEfficiency/ThisAddIn.Designer.cs" = 1
    "oPenEfficiency/UI/MainRibbon.cs" = 1
    "oPenEfficiency/Services/LocalAssetService.cs" = 1
    "oPenEfficiency/UI/Tabs/SlideLibraryTab.xaml.cs" = 1
    "oPenEfficiency/UI/Tabs/ShapeLibraryTab.xaml.cs" = 1
    "oPenEfficiency/UI/Tabs/ImageLibraryTab.xaml.cs" = 1
    "oPenEfficiency/UI/Tabs/FavoritesTab.xaml.cs" = 1
    "oPenEfficiency/UI/Controls/SizeAdjustControl.xaml.cs" = 1
    "oPenEfficiency/Features/Visuals/IconBrowserFeature.cs" = 1
    "oPenEfficiency/Features/Visuals/IllustrationLibraryFeature.cs" = 1
    "oPenEfficiency/UI/Dialogs/EditAssetDetailsDialog.xaml.cs" = 1
    "oPenEfficiency/UI/Dialogs/MagicResizerWindow.xaml.cs" = 1
    "oPenEfficiency/Features/Utilities/SnapToGridFeature.cs" = 2
    "oPenEfficiency/Features/Utilities/SnapToObjectsFeature.cs" = 2
}
$hostRoots = @((Join-Path $root "oPenEfficiency"), (Join-Path $root "oPenEfficiency.Excel"))
foreach ($file in Get-SourceFiles $hostRoots) {
    $text = [System.IO.File]::ReadAllText($file.FullName)
    $count = [regex]::Matches($text, 'Globals\.ThisAddIn').Count
    if ($count -eq 0) { continue }
    $relative = $file.FullName.Substring($root.Length + 1).Replace('\', '/')
    if (-not $globalAccessAllowlist.ContainsKey($relative)) {
        $violations.Add("Unapproved Globals.ThisAddIn access: $relative ($count)")
    }
    elseif ($count -gt $globalAccessAllowlist[$relative]) {
        $violations.Add("Globals.ThisAddIn access increased: $relative ($count > $($globalAccessAllowlist[$relative]))")
    }
}

$allProductionRoots = @($coreRoot) + $hostRoots
foreach ($file in Get-SourceFiles $allProductionRoots) {
    $text = [System.IO.File]::ReadAllText($file.FullName)
    if ($text -match 'catch\s*(?:\([^)]*\))?\s*\{\s*\}') {
        $violations.Add("Empty catch block: $($file.FullName.Substring($root.Length + 1))")
    }
}

Get-ChildItem -LiteralPath $root -Directory -Filter "*.Tests" | ForEach-Object {
    $project = Join-Path $_.FullName ($_.Name + ".csproj")
    $text = [System.IO.File]::ReadAllText($project)
    if ($text -match '<Compile\s+Include="\.\.\\oPenEfficiency(?:\\|\.)') {
        $violations.Add("Linked production source in test project: $($_.Name)")
    }
}

# QA-P0 Office-host stability: the add-in lives inside PowerPoint/Excel, so no timer may be
# backed by a Win32 WM_TIMER callback and the automation surface may not pump messages.
foreach ($file in Get-SourceFiles $hostRoots) {
    $text = [System.IO.File]::ReadAllText($file.FullName)
    if ($text -match 'new\s+(System\.Windows\.Forms|WinForms)\.Timer') {
        $violations.Add("WinForms WM_TIMER timer in an Office host: $($file.FullName.Substring($root.Length + 1)) (use DispatcherTimer at DispatcherPriority.ApplicationIdle)")
    }
}
$automationSurface = Join-Path $root "oPenEfficiency/ThisAddIn.Automation.cs"
if (Test-Path -LiteralPath $automationSurface) {
    if ([System.IO.File]::ReadAllText($automationSurface) -match 'Application\.DoEvents') {
        $violations.Add("Message pumping on the automation surface: oPenEfficiency/ThisAddIn.Automation.cs")
    }
}

# Public source builds must validate deployment identity without changing the user's Office registration.
# The explicit registration repair utility remains tested against a unique scratch key only.
$buildScript = Join-Path $root "eng\Build.ps1"
$resetScript = Join-Path $root "eng\Reset-AddInRegistration.ps1"
if (-not (Test-Path -LiteralPath $buildScript -PathType Leaf)) {
    $violations.Add("Build script missing: eng\Build.ps1")
}
else {
    $buildText = [System.IO.File]::ReadAllText($buildScript)
    if ($buildText -match 'Reset-AddInRegistration\.ps1' -or
        $buildText -notmatch 'Test-VstoDeploymentIdentity\.ps1' -or
        $buildText -notmatch 'Restore-VstoBuildState' -or
        $buildText -notmatch 'OPENEFFICIENCY_SKIP_OFFICE_REGISTRATION') {
        $violations.Add("Public build must validate deployment identity and must not automatically change Office registration.")
    }
}
if (-not (Test-Path -LiteralPath $resetScript -PathType Leaf)) {
    $violations.Add("Registration self-heal missing: eng\Reset-AddInRegistration.ps1 (REG-ADDIN-STARTUP-001).")
}
else {
    $probeKey = "HKCU:\Software\openEfficiencyPublicSourceProbe\" + [Guid]::NewGuid().ToString("N")
    $probeVsto = Join-Path ([System.IO.Path]::GetTempPath()) ("ope-selfheal-probe-" + $PID + ".vsto")
    try {
        Set-Content -LiteralPath $probeVsto -Value "self-heal probe"
        if (Test-Path -LiteralPath $probeKey) {
            Remove-Item -LiteralPath $probeKey -Recurse -Force
        }
        & $resetScript -RegistryKey $probeKey -VstoPath $probeVsto | Out-Null
        $probeManifest = [string](Get-ItemProperty -LiteralPath $probeKey -Name "Manifest").Manifest
        $probeLoadBehavior = [int](Get-ItemProperty -LiteralPath $probeKey -Name "LoadBehavior").LoadBehavior
        $expectedManifest = "file:///" + $probeVsto.Replace('\', '/') + "|vstolocal"
        if ($probeManifest -cne $expectedManifest -or $probeLoadBehavior -ne 3) {
            $violations.Add("Registration self-heal probe measured Manifest '$probeManifest' / LoadBehavior $probeLoadBehavior instead of '$expectedManifest' / 3 (REG-ADDIN-STARTUP-001).")
        }
    }
    catch {
        $violations.Add("Registration self-heal probe threw: $($_.Exception.Message) (REG-ADDIN-STARTUP-001)")
    }
    finally {
        if (Test-Path -LiteralPath $probeKey) {
            Remove-Item -LiteralPath $probeKey -Recurse -Force
        }
        if (Test-Path -LiteralPath $probeVsto) {
            Remove-Item -LiteralPath $probeVsto -Force
        }
    }
}

if ($violations.Count -gt 0) {
    $violations | ForEach-Object { Write-Error $_ }
    throw "Architecture validation failed with $($violations.Count) violation(s)."
}

Write-Host "ARCHITECTURE_BOUNDARIES_PASS"
Write-Host "Architecture validation passed: Core host-neutrality, global access ceilings, empty catches, Office-host timer and automation-surface pumping rules, and test assembly boundaries."
