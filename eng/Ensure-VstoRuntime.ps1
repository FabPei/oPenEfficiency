$ErrorActionPreference = "Stop"

$runtimeAssemblyName = "Microsoft.VisualStudio.Tools.Office.Runtime.dll"
$runtimeGacRoot = Join-Path `
    $env:windir `
    "Microsoft.NET\assembly\GAC_MSIL\Microsoft.VisualStudio.Tools.Office.Runtime"

function Find-VstoRuntimeAssembly {
    if (-not (Test-Path -LiteralPath $runtimeGacRoot -PathType Container)) {
        return $null
    }

    return Get-ChildItem `
        -LiteralPath $runtimeGacRoot `
        -Filter $runtimeAssemblyName `
        -File `
        -Recurse |
        Select-Object -First 1
}

$runtimeAssembly = Find-VstoRuntimeAssembly
if ($null -eq $runtimeAssembly) {
    $installerPath = Join-Path $env:RUNNER_TEMP "vstor_redist.exe"
    $downloadUrl = "https://download.microsoft.com/download/5/d/2/5d24f8f8-efbb-4b63-aa33-3785e3104713/vstor_redist.exe"

    try {
        Write-Host "Installing the Microsoft Visual Studio Tools for Office runtime."
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath
        $installer = Start-Process `
            -FilePath $installerPath `
            -ArgumentList "/quiet", "/norestart" `
            -WindowStyle Hidden `
            -Wait `
            -PassThru
        if ($installer.ExitCode -notin @(0, 3010)) {
            throw "VSTO runtime installation failed with exit code $($installer.ExitCode)."
        }
    }
    finally {
        if (Test-Path -LiteralPath $installerPath -PathType Leaf) {
            Remove-Item -LiteralPath $installerPath -Force
        }
    }

    $runtimeAssembly = Find-VstoRuntimeAssembly
}

if ($null -eq $runtimeAssembly) {
    throw "$runtimeAssemblyName is unavailable after VSTO runtime setup."
}

$runtimeVersion = [System.Reflection.AssemblyName]::GetAssemblyName(
    $runtimeAssembly.FullName).Version
Write-Host "VSTO runtime $runtimeVersion is available at $($runtimeAssembly.FullName)."
