[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$Subject = "CN=oPenEfficiency CI Ephemeral",

    [ValidateRange(1, 825)]
    [int]$ValidDays = 2
)

$ErrorActionPreference = "Stop"

$resolvedOutput = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = [System.IO.Path]::GetDirectoryName($resolvedOutput)
if (-not [System.IO.Directory]::Exists($outputDirectory)) {
    [System.IO.Directory]::CreateDirectory($outputDirectory) | Out-Null
}

$rsa = [System.Security.Cryptography.RSA]::Create(2048)
$request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
    $Subject,
    $rsa,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)

$request.CertificateExtensions.Add(
    [System.Security.Cryptography.X509Certificates.X509BasicConstraintsExtension]::new(
        $false,
        $false,
        0,
        $false))
$request.CertificateExtensions.Add(
    [System.Security.Cryptography.X509Certificates.X509KeyUsageExtension]::new(
        [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature,
        $false))

$enhancedKeyUsages = [System.Security.Cryptography.OidCollection]::new()
[void]$enhancedKeyUsages.Add(
    [System.Security.Cryptography.Oid]::new("1.3.6.1.5.5.7.3.3", "Code Signing"))
$request.CertificateExtensions.Add(
    [System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]::new(
        $enhancedKeyUsages,
        $false))

$certificate = $request.CreateSelfSigned(
    (Get-Date).AddMinutes(-5),
    (Get-Date).AddDays($ValidDays))

$pfxBytes = $certificate.Export(
    [System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12,
    [string]::Empty)
[System.IO.File]::WriteAllBytes($resolvedOutput, $pfxBytes)

[pscustomobject]@{
    PfxPath = $resolvedOutput
    Thumbprint = $certificate.Thumbprint
}

if (-not [string]::IsNullOrWhiteSpace($env:GITHUB_OUTPUT)) {
    Add-Content -LiteralPath $env:GITHUB_OUTPUT -Value "pfx_path=$resolvedOutput"
    Add-Content -LiteralPath $env:GITHUB_OUTPUT -Value "thumbprint=$($certificate.Thumbprint)"
}
