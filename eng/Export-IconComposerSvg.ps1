[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$InputPath,
    [Parameter(Mandatory)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'
# Use the same .NET Framework WPF geometry implementation as the add-in.
if ($PSVersionTable.PSEdition -eq 'Core') {
    & "$env:WINDIR/System32/WindowsPowerShell/v1.0/powershell.exe" -NoProfile -STA -File $PSCommandPath -InputPath $InputPath -OutputPath $OutputPath
    if ($LASTEXITCODE -ne 0) { throw "WPF icon export failed (exit $LASTEXITCODE)." }
    return
}

$inputFile = [IO.Path]::GetFullPath($InputPath)
$outputFile = [IO.Path]::GetFullPath($OutputPath)
if ($inputFile.Equals($outputFile, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Input and output paths must differ; retain the editable Composer source.'
}
Add-Type -Path (Join-Path $PSScriptRoot 'IconComposerSvgExporter.cs') -ReferencedAssemblies @(
    'System.Xml', 'System.Xml.Linq', 'System.Core', 'WindowsBase', 'PresentationCore')
$source = [IO.File]::ReadAllText($inputFile)
$svg = [IconComposerSvgExporter]::Export($source)
# No output is written until the entire document has passed conversion.
$directory = [IO.Path]::GetDirectoryName($outputFile)
[IO.Directory]::CreateDirectory($directory) | Out-Null
[IO.File]::WriteAllText($outputFile, $svg, [Text.UTF8Encoding]::new($false))
$hasher = [Security.Cryptography.SHA256]::Create()
try { $digest = ([BitConverter]::ToString($hasher.ComputeHash([Text.Encoding]::UTF8.GetBytes($svg)))).Replace('-', '').ToLowerInvariant() }
finally { $hasher.Dispose() }
[pscustomobject]@{
    profile = 'openefficiency'
    input = $inputFile
    output = $outputFile
    sha256 = $digest
} | ConvertTo-Json -Compress
