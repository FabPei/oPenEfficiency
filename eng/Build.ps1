[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')][string]$Configuration = 'Release',
    [switch]$Rebuild,
    [switch]$Installer
)
$ErrorActionPreference = 'Stop'
$repositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$solutionPath = Join-Path $repositoryRoot 'oPenEfficiency.slnx'
$resultsPath = Join-Path $repositoryRoot 'TestResults'

function Invoke-Checked([string]$File, [string[]]$Arguments) {
    & $File @Arguments
    if ($LASTEXITCODE -ne 0) { throw "$File failed with exit code $LASTEXITCODE." }
}
function Find-MSBuild {
    $command = Get-Command msbuild -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    $locator = Join-Path ([Environment]::GetFolderPath('ProgramFilesX86')) 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (Test-Path -LiteralPath $locator) {
        $found = @(& $locator -latest -products * -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe') | Select-Object -First 1
        if ($found -and (Test-Path -LiteralPath $found)) { return $found }
    }
    throw 'Install Visual Studio with Office/VSTO tools and the .NET Framework 4.8 SDK.'
}
$msbuild = Find-MSBuild
$dotnet = (Get-Command dotnet -ErrorAction Stop).Source
$buildStamp = Get-Date -Format 'yyyy.M.d.HHmm'
$useCommittedVstoApplicationVersion = $Installer.IsPresent.ToString().ToLowerInvariant()
$target = if ($Rebuild) { 'Rebuild' } else { 'Build' }
 . (Join-Path $PSScriptRoot 'VstoBuildState.ps1')
$registryState = Get-VstoBuildState $repositoryRoot
$savedRegistrationMode = $env:OPENEFFICIENCY_SKIP_OFFICE_REGISTRATION
$env:OPENEFFICIENCY_SKIP_OFFICE_REGISTRATION = 'true'
$savedStamp = $env:OPENEFFICIENCY_BUILD_STAMP
$savedMode = $env:OPENEFFICIENCY_USE_COMMITTED_VSTO_APPLICATION_VERSION
$env:OPENEFFICIENCY_BUILD_STAMP = $buildStamp
$env:OPENEFFICIENCY_USE_COMMITTED_VSTO_APPLICATION_VERSION = $useCommittedVstoApplicationVersion
$signingProps = Join-Path $PSScriptRoot 'LocalSigning.props'
$savedProps = if (Test-Path -LiteralPath $signingProps) { [IO.File]::ReadAllBytes($signingProps) } else { $null }
$temporaryKey = Join-Path ([IO.Path]::GetTempPath()) ('openEfficiency-build-' + [guid]::NewGuid().ToString('N') + '.pfx')
$certificate = $null
$certificateAlreadyPresent = $false
Push-Location $repositoryRoot
try {
    # This certificate signs local VSTO manifests only; it is not publisher trust.
    $certificate = & (Join-Path $PSScriptRoot 'New-EphemeralSigningCertificate.ps1') -OutputPath $temporaryKey
    $certificateAlreadyPresent = Test-Path -LiteralPath ('Cert:\CurrentUser\My\' + $certificate.Thumbprint)
    & (Join-Path $PSScriptRoot 'Write-SigningProps.ps1') -Thumbprint $certificate.Thumbprint -PfxPath $certificate.PfxPath | Out-Null
    Invoke-Checked $msbuild @($solutionPath, '/restore', "/t:$target", "/p:Configuration=$Configuration", '/p:Platform=Any CPU', '/p:OpenEfficiencySkipOfficeRegistration=true', "/p:BuildStamp=$buildStamp", "/p:UseCommittedVstoApplicationVersion=$useCommittedVstoApplicationVersion", '/v:minimal', '/nologo')
    [IO.Directory]::CreateDirectory($resultsPath) | Out-Null
    $solution = [xml][IO.File]::ReadAllText($solutionPath)
    foreach ($project in @($solution.Solution.Project | Where-Object { $_.Path -like '*.Tests/*.csproj' })) {
        $name = [IO.Path]::GetFileNameWithoutExtension($project.Path)
        $trx = Join-Path $resultsPath "$name.trx"
        if (Test-Path -LiteralPath $trx) { Remove-Item -LiteralPath $trx }
        Invoke-Checked $dotnet @('test', $project.Path, '--configuration', $Configuration, '--no-build', '--no-restore', '--results-directory', $resultsPath, '--logger', "trx;LogFileName=$name.trx", '--logger', 'console;verbosity=minimal')
    }
    # Use the real safe named-argument runner. Office/release profiles require separate host tooling/evidence.
    $powershell = Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
    Invoke-Checked $powershell @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $PSScriptRoot 'Invoke-FeatureAcceptance.ps1'), '-Configuration', $Configuration, '-ResultsPath', (Join-Path $resultsPath 'feature-acceptance.json'))
    & (Join-Path $PSScriptRoot 'Test-VstoDeploymentIdentity.ps1') -Configuration $Configuration -BuildStamp $buildStamp -UseCommittedApplicationVersion:$Installer
    if (-not $?) { throw 'Deployment identity validation failed.' }
    if ($Installer) {
        $projectXml = [xml][IO.File]::ReadAllText((Join-Path $repositoryRoot 'oPenEfficiency/oPenEfficiency.csproj'))
        $version = [Version]$projectXml.Project.PropertyGroup.ApplicationVersion.Where({ $_ })[0]
        $productVersion = '{0}.{1}.{2}' -f ($version.Major % 100), $version.Minor, $version.Build
        & (Join-Path $PSScriptRoot 'Set-ReleaseVersion.ps1') -Version $productVersion -Check
        foreach ($architecture in @('x86', 'x64')) {
            Invoke-Checked $dotnet @('build', 'installer/oPenEfficiency.Installer.wixproj', '--configuration', $Configuration, "-p:InstallerPlatform=$architecture", "-p:ProductVersion=$productVersion", '--nologo')
            $msi = Join-Path $repositoryRoot "installer/bin/$architecture/$Configuration/oPenEfficiency-$architecture-$productVersion.msi"
            if (-not (Test-Path -LiteralPath $msi -PathType Leaf)) { throw "Missing installer: $msi" }
            Get-FileHash -LiteralPath $msi -Algorithm SHA256 | Format-Table Path, Hash
        }
    }
    Write-Host 'PUBLIC_SOURCE_BUILD_PASS (not Office acceptance or release approval)'
}
finally {
    try { Restore-VstoBuildState $registryState }
    finally {
    if ($null -eq $savedProps) { Remove-Item -LiteralPath $signingProps -ErrorAction SilentlyContinue }
    else { [IO.File]::WriteAllBytes($signingProps, $savedProps) }
    if ($null -ne $certificate -and -not $certificateAlreadyPresent) {
        $createdCertificate = 'Cert:\CurrentUser\My\' + $certificate.Thumbprint
        if (Test-Path -LiteralPath $createdCertificate) { Remove-Item -LiteralPath $createdCertificate -DeleteKey }
    }
    if (Test-Path -LiteralPath $temporaryKey) { Remove-Item -LiteralPath $temporaryKey }
    Remove-Item Env:\OPENEFFICIENCY_BUILD_STAMP -ErrorAction SilentlyContinue
    Remove-Item Env:\OPENEFFICIENCY_USE_COMMITTED_VSTO_APPLICATION_VERSION -ErrorAction SilentlyContinue
    if ($null -ne $savedStamp) { $env:OPENEFFICIENCY_BUILD_STAMP = $savedStamp }
    if ($null -ne $savedMode) { $env:OPENEFFICIENCY_USE_COMMITTED_VSTO_APPLICATION_VERSION = $savedMode }
    Remove-Item Env:\OPENEFFICIENCY_SKIP_OFFICE_REGISTRATION -ErrorAction SilentlyContinue
    if ($null -ne $savedRegistrationMode) { $env:OPENEFFICIENCY_SKIP_OFFICE_REGISTRATION = $savedRegistrationMode }
    Pop-Location
    }
}
