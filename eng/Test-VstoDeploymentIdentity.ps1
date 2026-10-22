[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$BuildStamp,

    [switch]$UseCommittedApplicationVersion
)

$ErrorActionPreference = "Stop"
$repositoryRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$projects = @(
    [pscustomobject]@{ Name = "oPenEfficiency"; Directory = "oPenEfficiency" },
    [pscustomobject]@{ Name = "oPenEfficiency.Excel"; Directory = "oPenEfficiency.Excel" }
)

function Get-ManifestIdentityVersion([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "VSTO identity output is missing: '$Path'."
    }

    $document = [xml][System.IO.File]::ReadAllText($Path)
    $identity = $document.SelectSingleNode(
        "/*[local-name()='assembly']/*[local-name()='assemblyIdentity']")
    if ($null -eq $identity -or [string]::IsNullOrWhiteSpace($identity.version)) {
        throw "VSTO identity output '$Path' has no assemblyIdentity version."
    }

    return [string]$identity.version
}

foreach ($project in $projects) {
    $projectPath = Join-Path $repositoryRoot ($project.Directory + "\" + $project.Name + ".csproj")
    $outputDirectory = Join-Path $repositoryRoot ($project.Directory + "\bin\" + $Configuration)
    $expectedApplicationVersion = $BuildStamp
    if ($UseCommittedApplicationVersion) {
        $projectDocument = [xml][System.IO.File]::ReadAllText($projectPath)
        $applicationVersion = $projectDocument.SelectSingleNode(
            "/*[local-name()='Project']/*[local-name()='PropertyGroup']/*[local-name()='ApplicationVersion']")
        if ($null -eq $applicationVersion -or [string]::IsNullOrWhiteSpace($applicationVersion.InnerText)) {
            throw "Committed ApplicationVersion is missing from '$projectPath'."
        }
        $expectedApplicationVersion = $applicationVersion.InnerText
    }

    $deploymentManifest = Join-Path $outputDirectory ($project.Name + ".vsto")
    $applicationManifest = Join-Path $outputDirectory ($project.Name + ".dll.manifest")
    foreach ($manifestPath in @($deploymentManifest, $applicationManifest)) {
        $actualVersion = Get-ManifestIdentityVersion $manifestPath
        if ($actualVersion -cne $expectedApplicationVersion) {
            throw "VSTO identity mismatch in '$manifestPath': expected '$expectedApplicationVersion', found '$actualVersion'. A stale ClickOnce identity can make Office load an older shadow copy."
        }
    }

    $assemblyPath = Join-Path $outputDirectory ($project.Name + ".dll")
    if (-not (Test-Path -LiteralPath $assemblyPath -PathType Leaf)) {
        throw "VSTO assembly output is missing: '$assemblyPath'."
    }
    # AssemblyVersion stores four numeric components, so zero-padded stamp parts are normalized on
    # read-back ("2026.9.28.0020" -> "2026.9.28.20"). Compare parsed Version identities; a raw string
    # compare fails for build stamps between 00:00 and 09:59 local time even though the identities are
    # equal (REG-BUILD-STAMP-HOUR-001).
    $assemblyVersion = [System.Reflection.AssemblyName]::GetAssemblyName($assemblyPath).Version
    $expectedAssemblyVersion = [System.Version]::Parse($BuildStamp)
    if ($assemblyVersion -ne $expectedAssemblyVersion) {
        throw "Assembly identity mismatch in '$assemblyPath': expected build stamp '$BuildStamp' ($expectedAssemblyVersion), found '$($assemblyVersion.ToString())'."
    }
}

$identityMode = if ($UseCommittedApplicationVersion) { "committed-release" } else { "per-build" }
Write-Host "VSTO_DEPLOYMENT_IDENTITY_PASS configuration=$Configuration buildStamp=$BuildStamp manifestMode=$identityMode"
