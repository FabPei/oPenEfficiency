[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version,

    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$ApplicationVersion,

    [switch]$Check,

    [string]$RepositoryRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = "Stop"

$versionParts = $Version.Split(".")
$installerMajor = [int]$versionParts[0]
$installerMinor = [int]$versionParts[1]
$installerPatch = [int]$versionParts[2]
if ($installerMajor -gt 255 -or
    $installerMinor -gt 255 -or
    $installerPatch -gt 65535) {
    throw "Version '$Version' is outside Windows Installer limits (major/minor <= 255, patch <= 65535)."
}

if ([string]::IsNullOrWhiteSpace($ApplicationVersion)) {
    if ($installerMajor -gt 99 -or
        $installerMinor -lt 1 -or $installerMinor -gt 12 -or
        $installerPatch -lt 1 -or $installerPatch -gt 31) {
        throw "Version '$Version' cannot be converted to VSTO CalVer. Use YY.M.D (for example, 26.7.27) or pass -ApplicationVersion explicitly."
    }

    $ApplicationVersion = "{0:D4}.{1:D2}.{2:D2}.1" -f (
        2000 + $installerMajor),
        $installerMinor,
        $installerPatch
}

$applicationParts = $ApplicationVersion.Split(".") | ForEach-Object { [int]$_ }
if ($applicationParts | Where-Object { $_ -gt 65535 }) {
    throw "ApplicationVersion '$ApplicationVersion' is outside four-part .NET version limits (each part <= 65535)."
}

$targetApplicationVersion = [System.Version]$ApplicationVersion
$projectFiles = @(
    (Join-Path $RepositoryRoot "oPenEfficiency\oPenEfficiency.csproj"),
    (Join-Path $RepositoryRoot "oPenEfficiency.Excel\oPenEfficiency.Excel.csproj")
)
$assemblyInfoFiles = @(
    (Join-Path $RepositoryRoot "oPenEfficiency\Properties\AssemblyInfo.cs"),
    (Join-Path $RepositoryRoot "oPenEfficiency.Excel\Properties\AssemblyInfo.cs")
)

# Current VSTO projects generate both attributes centrally; older checkouts retain literals.
# Validate every input before writing any version update.
$pendingUpdates = [ordered]@{}
$generatedAssemblyIdentity = $false
$propsPath = Join-Path $RepositoryRoot 'Directory.Build.props'
if (Test-Path -LiteralPath $propsPath) {
    [xml]$props = [IO.File]::ReadAllText($propsPath)
    $stamp = $props.SelectSingleNode("/Project/Target[@Name='StampBuildIdentity']")
    if ($null -ne $stamp) {
        $lines = @($stamp.SelectNodes('ItemGroup/_BuildStampLine') | ForEach-Object { $_.GetAttribute('Include') })
        $writer = $stamp.SelectSingleNode('WriteLinesToFile')
        $compile = $stamp.SelectSingleNode('ItemGroup/Compile')
        $generatedAssemblyIdentity = $stamp.GetAttribute('BeforeTargets') -eq 'CoreCompile' -and
            @($lines | Where-Object { $_ -match '^\[assembly: System\.Reflection\.AssemblyVersion\(' }).Count -eq 1 -and
            @($lines | Where-Object { $_ -match '^\[assembly: System\.Reflection\.AssemblyFileVersion\(' }).Count -eq 1 -and
            $null -ne $writer -and $writer.GetAttribute('File') -eq '$(_BuildStampInfoFile)' -and
            $null -ne $compile -and $compile.GetAttribute('Include') -eq '$(_BuildStampInfoFile)'
    }
}

foreach ($projectFile in $projectFiles) {
    $content = [System.IO.File]::ReadAllText($projectFile)
    $match = [regex]::Match(
        $content,
        "<ApplicationVersion>(.*?)</ApplicationVersion>")
    if (-not $match.Success) {
        throw "ApplicationVersion was not found in '$projectFile'."
    }

    $currentApplicationVersion = [System.Version]$match.Groups[1].Value
    if ($Check -and $targetApplicationVersion -ne $currentApplicationVersion) {
        throw "Committed ApplicationVersion '$currentApplicationVersion' in '$projectFile' does not match required release version '$ApplicationVersion'."
    }
    if (-not $Check -and $targetApplicationVersion -lt $currentApplicationVersion) {
        throw "ApplicationVersion '$ApplicationVersion' is older than '$currentApplicationVersion' in '$projectFile'."
    }

    $updated = [regex]::Replace(
        $content,
        "<ApplicationVersion>.*?</ApplicationVersion>",
        "<ApplicationVersion>$ApplicationVersion</ApplicationVersion>")
    if (-not $Check -and $updated -ne $content) {
        $pendingUpdates[$projectFile] = $updated
    }
}

foreach ($assemblyInfoFile in $assemblyInfoFiles) {
    $content = [System.IO.File]::ReadAllText($assemblyInfoFile)
    $hasAssemblyVersion = $content -match '(?m)^\[assembly:\s*AssemblyVersion\(".*?"\)\]'
    $hasFileVersion = $content -match '(?m)^\[assembly:\s*AssemblyFileVersion\(".*?"\)\]'
    if ($generatedAssemblyIdentity -and ($hasAssemblyVersion -or $hasFileVersion)) {
        throw "Central and literal assembly version producers conflict in '$assemblyInfoFile'."
    }
    if (-not $hasAssemblyVersion -and -not $hasFileVersion -and $generatedAssemblyIdentity) {
        continue # Per-build identity deliberately differs from committed VSTO CalVer.
    }
    if (-not $hasAssemblyVersion -or -not $hasFileVersion) {
        throw "Assembly version attributes or the complete central build-stamp target were not found in '$assemblyInfoFile'."
    }

    $assemblyVersionMatch = [regex]::Match(
        $content,
        '(?m)^\[assembly:\s*AssemblyVersion\("(.*?)"\)\]')
    $assemblyFileVersionMatch = [regex]::Match(
        $content,
        '(?m)^\[assembly:\s*AssemblyFileVersion\("(.*?)"\)\]')
    if ($Check -and
        ($assemblyVersionMatch.Groups[1].Value -ne $ApplicationVersion -or
         $assemblyFileVersionMatch.Groups[1].Value -ne $ApplicationVersion)) {
        throw "Committed assembly versions in '$assemblyInfoFile' do not match required release version '$ApplicationVersion'."
    }

    $updated = [regex]::Replace(
        $content,
        '(?m)^\[assembly:\s*AssemblyVersion\(".*?"\)\]',
        "[assembly: AssemblyVersion(`"$ApplicationVersion`")]")
    $updated = [regex]::Replace(
        $updated,
        '(?m)^\[assembly:\s*AssemblyFileVersion\(".*?"\)\]',
        "[assembly: AssemblyFileVersion(`"$ApplicationVersion`")]")
    if (-not $Check -and $updated -ne $content) {
        $pendingUpdates[$assemblyInfoFile] = $updated
    }
}

foreach ($entry in $pendingUpdates.GetEnumerator()) {
    $bytes = [IO.File]::ReadAllBytes($entry.Key)
    $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191
    [IO.File]::WriteAllText($entry.Key, $entry.Value, [Text.UTF8Encoding]::new($hasBom))
}

Write-Output "Mode=$(if ($Check) { 'Check' } else { 'Update' })"
Write-Output "ApplicationVersion=$ApplicationVersion"
Write-Output "InstallerVersion=$Version"
