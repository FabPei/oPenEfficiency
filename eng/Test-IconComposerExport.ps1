[CmdletBinding()]
param([string]$ProductionDirectory)
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$artifacts = Join-Path $root 'artifacts/icon-composer-export'
[IO.Directory]::CreateDirectory($artifacts) | Out-Null
& dotnet build (Join-Path $PSScriptRoot 'IconComposer.Export.Tests.csproj') -c Release --nologo *> (Join-Path $artifacts 'tool-build.log')
if ($LASTEXITCODE -ne 0) { throw "Icon export test build failed; see $artifacts/tool-build.log" }
$binary = Join-Path $artifacts 'bin/IconComposer.Export.Tests.exe'
& $binary $root *> (Join-Path $artifacts 'unit-tests.log')
if ($LASTEXITCODE -ne 0) { throw "Icon export tests failed; see $artifacts/unit-tests.log" }
if ($ProductionDirectory) {
    & $binary $root '--production-files' ((Resolve-Path -LiteralPath $ProductionDirectory -ErrorAction Stop).ProviderPath) *> (Join-Path $artifacts 'bundle-production-parser.log')
    if ($LASTEXITCODE -ne 0) { throw "Bundled production SVG parser checks failed; see $artifacts/bundle-production-parser.log" }
}
Write-Output 'ICON_COMPOSER_EXPORT_PASS runtime=net48 renderer=production-SvgParser-and-FeatureIconPathFactory'
