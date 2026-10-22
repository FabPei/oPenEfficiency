[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9A-Fa-f]{40}$')]
    [string]$Thumbprint,

    [string]$PfxPath = "",

    [string]$OutputPath = (Join-Path $PSScriptRoot "LocalSigning.props")
)

$ErrorActionPreference = "Stop"

$resolvedProps = [System.IO.Path]::GetFullPath($OutputPath)
$propsDirectory = [System.IO.Path]::GetDirectoryName($resolvedProps)
if (-not [System.IO.Directory]::Exists($propsDirectory)) {
    [System.IO.Directory]::CreateDirectory($propsDirectory) | Out-Null
}

$keyFileElement = ""
$resolvedPfx = ""
if (-not [string]::IsNullOrWhiteSpace($PfxPath)) {
    $resolvedPfx = [System.IO.Path]::GetFullPath($PfxPath)
    if (-not [System.IO.File]::Exists($resolvedPfx)) {
        throw "Signing certificate was not found at '$resolvedPfx'."
    }
    $escapedPfxPath =
        [System.Security.SecurityElement]::Escape($resolvedPfx)
    $keyFileElement =
        "    <ManifestKeyFile>$escapedPfxPath</ManifestKeyFile>$([Environment]::NewLine)"
}

$normalizedThumbprint = $Thumbprint.ToUpperInvariant()
$propsContent = @"
<Project>
  <PropertyGroup>
$keyFileElement    <ManifestCertificateThumbprint>$normalizedThumbprint</ManifestCertificateThumbprint>
  </PropertyGroup>
</Project>
"@
[System.IO.File]::WriteAllText($resolvedProps, $propsContent)

[pscustomobject]@{
    PropsPath = $resolvedProps
    PfxPath = $resolvedPfx
    Thumbprint = $normalizedThumbprint
}
