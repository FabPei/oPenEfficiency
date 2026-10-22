[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    # Test seam: the scratch-key probe in eng/Test-Architecture.ps1 runs this script against its own
    # key instead of the live PowerPoint registration.
    [string]$RegistryKey = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\oPenEfficiency',

    # Test seam: defaults to the output this build produced.
    [string]$VstoPath
)

# REG-ADDIN-STARTUP-001 self-heal: every build of the VSTO add-in registers it for its own
# configuration (Microsoft.VisualStudio.Tools.Office.targets writes the Addins key), so a Debug
# build can leave HKCU pointing at bin\Debug and stall PowerPoint's startup on a stale manifest.
# This reset points the registration at the freshly built output and re-reads what it wrote, so a
# wrong value fails the build instead of being reported as success.
$ErrorActionPreference = "Stop"
$repositoryRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$stageLocalManifest = [string]::IsNullOrWhiteSpace($VstoPath)
$removedStaleMetadata = 0
$canonicalInclusionCreated = $false
if ([string]::IsNullOrWhiteSpace($VstoPath)) {
    $VstoPath = Join-Path $repositoryRoot "oPenEfficiency\bin\$Configuration\oPenEfficiency.vsto"
}
if (-not (Test-Path -LiteralPath $VstoPath -PathType Leaf)) {
    throw "The built VSTO manifest does not exist at '$VstoPath' (configuration $Configuration); registering a missing manifest is the REG-ADDIN-STARTUP-001 startup hang again."
}

$canonicalManifestUri = "file:///" + $VstoPath.Replace('\', '/')

# Visual Studio adds a SolutionMetadata and trust-inclusion record for every worktree that builds a
# VSTO project. All PowerPoint worktrees use the same add-in name, and VSTO can resolve that name to
# an older metadata record even while Office's Addins key points at this repository. Remove only the
# stale PowerPoint/oPenEfficiency development records; Excel and the canonical manifest are kept.
if ($stageLocalManifest) {
    $solutionMetadataRoot = 'HKCU:\Software\Microsoft\VSTO\SolutionMetadata'
    if (Test-Path -LiteralPath $solutionMetadataRoot) {
        $metadataRootValues = Get-ItemProperty -LiteralPath $solutionMetadataRoot
        $metadataValues = $metadataRootValues.PSObject.Properties |
            Where-Object { $_.Name -like 'file:///*' }
        $canonicalMetadata = $metadataValues |
            Where-Object { [string]::Equals($_.Name, $canonicalManifestUri, [StringComparison]::OrdinalIgnoreCase) } |
            Select-Object -First 1
        $canonicalMetadataGuid = if ($null -ne $canonicalMetadata) {
            [string]$canonicalMetadata.Value
        }
        else {
            '{' + [Guid]::NewGuid().ToString().ToUpperInvariant() + '}'
        }
        foreach ($metadataValue in $metadataValues) {
            $metadataKey = Join-Path $solutionMetadataRoot ([string]$metadataValue.Value)
            if (-not (Test-Path -LiteralPath $metadataKey)) { continue }
            $metadata = Get-ItemProperty -LiteralPath $metadataKey
            if ([string]$metadata.addInName -cne 'oPenEfficiency' -or
                [string]$metadata.officeApplication -cne 'PowerPoint' -or
                [string]::Equals($metadataValue.Name, $canonicalManifestUri, [StringComparison]::OrdinalIgnoreCase)) {
                continue
            }
            Remove-ItemProperty -LiteralPath $solutionMetadataRoot -Name $metadataValue.Name -Force
            if ([string]$metadataValue.Value -cne $canonicalMetadataGuid) {
                Remove-Item -LiteralPath $metadataKey -Recurse -Force
            }
            $removedStaleMetadata++
        }

        New-ItemProperty -LiteralPath $solutionMetadataRoot -Name $canonicalManifestUri `
            -PropertyType String -Value $canonicalMetadataGuid -Force | Out-Null
        $canonicalMetadataKey = Join-Path $solutionMetadataRoot $canonicalMetadataGuid
        if (-not (Test-Path -LiteralPath $canonicalMetadataKey)) {
            New-Item -Path $canonicalMetadataKey -Force | Out-Null
        }
        $compatibleFrameworks = '<compatibleFrameworks xmlns="urn:schemas-microsoft-com:clickonce.v2">' +
            "`r`n  <framework targetVersion=`"4.8`" profile=`"Full`" supportedRuntime=`"4.0.30319`" />`r`n" +
            '</compatibleFrameworks>'
        foreach ($entry in @{
            addInName = 'oPenEfficiency'
            officeApplication = 'PowerPoint'
            friendlyName = 'oPenEfficiency'
            description = 'oPenEfficiency'
            loadBehavior = '3'
            compatibleFrameworks = $compatibleFrameworks
            PreferredClr = 'v4.0.30319'
        }.GetEnumerator()) {
            New-ItemProperty -LiteralPath $canonicalMetadataKey -Name $entry.Key `
                -PropertyType String -Value $entry.Value -Force | Out-Null
        }
    }

    $inclusionRoot = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
    if (Test-Path -LiteralPath $inclusionRoot) {
        foreach ($inclusionKey in @(Get-ChildItem -LiteralPath $inclusionRoot -ErrorAction SilentlyContinue)) {
            $url = [string](Get-ItemProperty -LiteralPath $inclusionKey.PSPath -Name Url -ErrorAction SilentlyContinue).Url
            if ($url -notmatch '(?i)/oPenEfficiency/bin/(Debug|Release)/oPenEfficiency\.vsto$' -or
                [string]::Equals($url, $canonicalManifestUri, [StringComparison]::OrdinalIgnoreCase)) {
                continue
            }
            Remove-Item -LiteralPath $inclusionKey.PSPath -Recurse -Force
            $removedStaleMetadata++
        }
    }

    # "One build identity per run" (57b7b835) deletes the OTHER checkouts trust records, but the
    # cleanup never created the canonical record itself: a worktree build deleted the main
    # checkout's Inclusion entry and left NOTHING trusted, so every subsequent smoke failed with
    # "The add-in could not be installed" (VSTO 4.0 event 4096, VerifyAddInTrust - measured by the
    # batch-55-fix-R3 review). Re-create the canonical trust record from the manifest's own embedded
    # signature key whenever it is missing, so the reset target is always trusted.
    if (-not (Test-Path -LiteralPath $inclusionRoot)) {
        New-Item -Path $inclusionRoot -Force | Out-Null
    }
    $canonicalInclusion = @(Get-ChildItem -LiteralPath $inclusionRoot -ErrorAction SilentlyContinue) |
        Where-Object {
            $entryUrl = [string](Get-ItemProperty -LiteralPath $_.PSPath -Name Url -ErrorAction SilentlyContinue).Url
            [string]::Equals($entryUrl, $canonicalManifestUri, [StringComparison]::OrdinalIgnoreCase)
        }
    if ($canonicalInclusion.Count -eq 0) {
        $manifestText = Get-Content -LiteralPath $VstoPath -Raw
        $keyMatch = [regex]::Match($manifestText, '<RSAKeyValue>.*?</RSAKeyValue>', [Text.RegularExpressions.RegexOptions]::Singleline)
        if (-not $keyMatch.Success) {
            throw "The built manifest at '$VstoPath' carries no RSAKeyValue signature block; granting trust without the manifests own public key would register an add-in PowerPoint rejects with VerifyAddInTrust."
        }
        $canonicalInclusionKey = Join-Path $inclusionRoot ([Guid]::NewGuid().ToString().ToUpperInvariant())
        New-Item -Path $canonicalInclusionKey -Force | Out-Null
        New-ItemProperty -LiteralPath $canonicalInclusionKey -Name Url -PropertyType String -Value $canonicalManifestUri -Force | Out-Null
        New-ItemProperty -LiteralPath $canonicalInclusionKey -Name PublicKey -PropertyType String -Value $keyMatch.Value -Force | Out-Null
        $canonicalInclusionCreated = $true
    }
}

# REG-ADDIN-STARTUP-001 (second half): a stale registration can make VSTO actually INSTALL the
# customization (an Uninstall\* entry with DisplayName oPenEfficiency appears). That installed state
# then rejects every load from another location with AddInAlreadyInstalledException - even after the
# Addins key was reset - which leaves PowerPoint starting without the add-in at all. The canonical
# developer state is the '|vstolocal' registration WITHOUT an installed solution, so remove any
# oPenEfficiency install state first; VSTOInstaller deletes the Addins key as well, which the
# registration below then rewrites. Only the live registration triggers this (the Test-Architecture
# scratch-key seam must not uninstall anything).
$isLiveRegistration = $RegistryKey -eq 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\oPenEfficiency'
$removedInstallStates = 0
if ($isLiveRegistration) {
    $uninstallRoots = @(
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    $vstoInstallerCandidates = @(
        (Join-Path $env:CommonProgramFiles 'microsoft shared\VSTO\10.0\VSTOInstaller.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe')
    )
    $vstoInstaller = $null
    foreach ($candidate in $vstoInstallerCandidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) { $vstoInstaller = $candidate; break }
    }
    if ($vstoInstaller) {
        foreach ($uninstallRoot in $uninstallRoots) {
            if (-not (Test-Path -LiteralPath $uninstallRoot)) { continue }
            foreach ($entryKey in @(Get-ChildItem -LiteralPath $uninstallRoot -ErrorAction SilentlyContinue)) {
                $entry = Get-ItemProperty -LiteralPath $entryKey.PSPath -ErrorAction SilentlyContinue
                if ([string]$entry.DisplayName -cne 'oPenEfficiency') { continue }
                $uriMatch = [regex]::Match([string]$entry.UninstallString, 'file:///\S+')
                if (-not $uriMatch.Success) { continue }
                $uninstallProcess = Start-Process -FilePath $vstoInstaller `
                    -ArgumentList @('/Uninstall', $uriMatch.Value, '/Silent') -Wait -PassThru -NoNewWindow
                if ($uninstallProcess.ExitCode -eq 0) { $removedInstallStates++ }
            }
        }
    }
}

$manifest = "file:///" + $VstoPath.Replace('\', '/') + "|vstolocal"
if (-not (Test-Path -LiteralPath $RegistryKey)) {
    New-Item -Path $RegistryKey -Force | Out-Null
}
New-ItemProperty -Path $RegistryKey -Name "Manifest" -PropertyType String -Value $manifest -Force | Out-Null
New-ItemProperty -Path $RegistryKey -Name "LoadBehavior" -PropertyType DWord -Value 3 -Force | Out-Null

$writtenManifest = [string](Get-ItemProperty -LiteralPath $RegistryKey -Name "Manifest").Manifest
$writtenLoadBehavior = [int](Get-ItemProperty -LiteralPath $RegistryKey -Name "LoadBehavior").LoadBehavior
if ($writtenManifest -cne $manifest) {
    throw "The add-in registration did not take: '$RegistryKey' Manifest reads '$writtenManifest' instead of '$manifest'."
}
if ($writtenLoadBehavior -ne 3) {
    throw "The add-in registration did not take: '$RegistryKey' LoadBehavior reads $writtenLoadBehavior instead of 3."
}
Write-Host "ADDIN_REGISTRATION_RESET_PASS configuration=$Configuration key=$RegistryKey manifest=$manifest loadBehavior=$writtenLoadBehavior staleMetadataRemoved=$removedStaleMetadata solutionInstallStatesRemoved=$removedInstallStates canonicalInclusionCreated=$canonicalInclusionCreated"
