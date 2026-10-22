$ErrorActionPreference = "Stop"

$windowsKitsRoot = Join-Path `
    ${env:ProgramFiles(x86)} `
    "Windows Kits\10\UnionMetadata"

if (-not (Test-Path -LiteralPath $windowsKitsRoot -PathType Container)) {
    throw "Windows SDK union metadata was not found at $windowsKitsRoot."
}

$windowsWinMd = Get-ChildItem `
        -LiteralPath $windowsKitsRoot `
        -Directory |
    Where-Object {
        $version = $null
        [Version]::TryParse($_.Name, [ref]$version)
    } |
    Sort-Object { [Version]$_.Name } -Descending |
    ForEach-Object {
        Get-Item `
            -LiteralPath (Join-Path $_.FullName "Windows.winmd") `
            -ErrorAction SilentlyContinue
    } |
    Select-Object -First 1

if ($null -eq $windowsWinMd) {
    throw "No versioned Windows.winmd was found below $windowsKitsRoot."
}

if ([string]::IsNullOrWhiteSpace($env:GITHUB_ENV)) {
    $env:WindowsWinMdPath = $windowsWinMd.FullName
}
else {
    "WindowsWinMdPath=$($windowsWinMd.FullName)" |
        Add-Content -LiteralPath $env:GITHUB_ENV -Encoding utf8
}

Write-Host "Using Windows SDK metadata at $($windowsWinMd.FullName)."
