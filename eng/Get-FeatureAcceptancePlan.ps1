[CmdletBinding()]
param(
    [string[]]$FeatureId,
    [string[]]$ChangedPath,
    [string[]]$AdditionalScenarioId,
    [string]$ManifestDirectory,
    [string]$ResultsPath
)

$ErrorActionPreference = "Stop"
$repositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
if ([string]::IsNullOrWhiteSpace($ManifestDirectory)) {
    $ManifestDirectory = Join-Path $repositoryRoot "docs\testing\feature-acceptance"
}
if ([string]::IsNullOrWhiteSpace($ResultsPath)) {
    $ResultsPath = Join-Path $repositoryRoot "TestResults\feature-acceptance-plan.json"
}
$catalogManifestDirectory = Join-Path $ManifestDirectory "catalog"
$inventoryDocumentNames = @("manifest.schema.json", "catalog-coverage.json", "catalog-exclusions.json", "shared-risk-routes.json", "manual-review-migration.json", "regression-ledger.json")

function Normalize-Path([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $value = $Path -replace '\\', '/'
    if ($value.StartsWith("./", [StringComparison]::Ordinal)) { $value = $value.Substring(2) }
    return $value
}

function Matches-Any([string]$Path, $Patterns) {
    foreach ($pattern in @($Patterns)) {
        $normalizedPattern = Normalize-Path ([string]$pattern)
        if ($Path -like $normalizedPattern) { return $true }
    }
    return $false
}

function Read-TextFile([string]$Path) {
    return [IO.File]::ReadAllText($Path, [Text.UTF8Encoding]::new($false))
}

$requestedFeatures = @($FeatureId | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
$changed = @($ChangedPath | ForEach-Object { Normalize-Path $_ } | Where-Object { $_ } | Sort-Object -Unique)
$requestedScenarioIds = @($AdditionalScenarioId | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
$manifestFiles = [Collections.Generic.List[object]]::new()
foreach ($file in @(Get-ChildItem -LiteralPath $ManifestDirectory -Filter "*.json" -File |
    Where-Object { $inventoryDocumentNames -notcontains $_.Name } |
    Sort-Object Name)) {
    $manifestFiles.Add($file)
}
if (Test-Path -LiteralPath $catalogManifestDirectory -PathType Container) {
    foreach ($file in @(Get-ChildItem -LiteralPath $catalogManifestDirectory -Filter "*.json" -File | Sort-Object Name)) {
        $manifestFiles.Add($file)
    }
}

$selected = [Collections.Generic.List[object]]::new()
$resolvedRequestedScenarios = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach ($file in $manifestFiles) {
    $manifest = (Read-TextFile -Path $file.FullName) | ConvertFrom-Json
    $featureExplicit = $requestedFeatures.Count -gt 0 -and ($requestedFeatures -contains [string]$manifest.featureId)
    $featureByPath = $changed.Count -gt 0 -and @($changed | Where-Object { Matches-Any $_ $manifest.sourcePaths }).Count -gt 0
    $manifestRequestedScenarios = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($scenario in @($manifest.scenarios)) {
        if ($requestedScenarioIds -contains [string]$scenario.id) {
            $null = $manifestRequestedScenarios.Add([string]$scenario.id)
        }
    }

    if ($requestedFeatures.Count -gt 0 -and -not $featureExplicit) { continue }
    if ($requestedFeatures.Count -eq 0 -and $changed.Count -gt 0 -and
        -not $featureByPath -and $manifestRequestedScenarios.Count -eq 0) { continue }

    foreach ($scenario in @($manifest.scenarios)) {
        $scenarioId = [string]$scenario.id
        $scenarioByPath = $changed.Count -eq 0 -or @($changed | Where-Object { Matches-Any $_ $scenario.changedPathPatterns }).Count -gt 0
        $scenarioByRoute = $manifestRequestedScenarios.Contains($scenarioId)
        if (-not $featureExplicit -and $changed.Count -gt 0 -and -not $scenarioByPath -and -not $scenarioByRoute) { continue }
        if (-not $featureExplicit -and $changed.Count -eq 0 -and $requestedScenarioIds.Count -gt 0 -and -not $scenarioByRoute) { continue }
        if ($scenarioByRoute) { $null = $resolvedRequestedScenarios.Add($scenarioId) }

        $selected.Add([pscustomobject]@{
            FeatureId = [string]$manifest.featureId
            FeatureTitle = [string]$manifest.title
            Id = $scenarioId
            Title = [string]$scenario.title
            Tier = [string]$scenario.tier
            AutomationStatus = [string]$scenario.automationStatus
            Expected = [string]$scenario.expected
            Notes = [string]$scenario.notes
            OwnerSteps = @($scenario.ownerSteps)
            Runner = $scenario.runner
            SharedRiskRoute = [bool]$scenarioByRoute
            ManifestFile = $file.Name
        })
    }
}

$missingRequestedScenarios = @($requestedScenarioIds | Where-Object { -not $resolvedRequestedScenarios.Contains($_) })
if ($missingRequestedScenarios.Count -gt 0) {
    throw ("Shared-risk routing requested acceptance scenario(s) that no manifest declares: " +
        "$($missingRequestedScenarios -join ', '). A route must never select an unknown scenario.")
}

$items = @($selected | Sort-Object Id)
$result = [pscustomobject]@{
    SchemaVersion = 1
    CapturedUtc = [DateTime]::UtcNow.ToString("o")
    RequestedFeatures = $requestedFeatures
    ChangedPaths = $changed
    ScenarioCount = $items.Count
    AutomatedReady = @($items | Where-Object { $_.Tier -in @("core-automated","office-automated") -and $_.AutomationStatus -eq "ready" }).Count
    AutomatedPlannedOrBlocked = @($items | Where-Object { $_.Tier -in @("core-automated","office-automated") -and $_.AutomationStatus -ne "ready" }).Count
    OwnerOrVisual = @($items | Where-Object { $_.Tier -in @("owner-manual","visual-review") }).Count
    OwnerManualRequired = @($items | Where-Object { $_.Tier -in @("owner-manual","visual-review") }).Count -gt 0
    SharedRiskScenarioCount = @($items | Where-Object { $_.SharedRiskRoute }).Count
    Scenarios = $items
}
$fullResultsPath = [IO.Path]::GetFullPath($ResultsPath)
[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($fullResultsPath)) | Out-Null
$result | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $fullResultsPath -Encoding UTF8
Write-Output ("FEATURE_ACCEPTANCE_PLAN scenarios={0} automated_ready={1} automation_open={2} owner_or_visual={3} shared_risk={4} results={5}" -f
    $result.ScenarioCount, $result.AutomatedReady, $result.AutomatedPlannedOrBlocked, $result.OwnerOrVisual, $result.SharedRiskScenarioCount, $fullResultsPath)
