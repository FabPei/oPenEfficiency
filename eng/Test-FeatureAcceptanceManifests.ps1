[CmdletBinding()]
param(
    [string]$ManifestDirectory
)

$ErrorActionPreference = "Stop"
$repositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
if ([string]::IsNullOrWhiteSpace($ManifestDirectory)) {
    $ManifestDirectory = Join-Path $repositoryRoot "docs\testing\feature-acceptance"
}
$manifestDirectoryFull = [IO.Path]::GetFullPath($ManifestDirectory)
if (-not (Test-Path -LiteralPath $manifestDirectoryFull -PathType Container)) {
    throw "Feature acceptance manifest directory was not found: $manifestDirectoryFull"
}

$allowedTiers = @("core-automated", "office-automated", "visual-review", "owner-manual")
$allowedStatuses = @("ready", "planned", "blocked")
# Inventory documents share this directory but are not acceptance manifests.
$inventoryDocumentNames = @("manifest.schema.json", "catalog-coverage.json", "catalog-exclusions.json", "shared-risk-routes.json", "manual-review-migration.json", "regression-ledger.json", "visual-evidence-reviews.json")
$unavailableOfficeRunners = 0
$featureIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$scenarioIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$manifestFiles = [Collections.Generic.List[object]]::new()
foreach ($file in @(Get-ChildItem -LiteralPath $manifestDirectoryFull -Filter "*.json" -File |
    Where-Object { $inventoryDocumentNames -notcontains $_.Name } |
    Sort-Object Name)) {
    $manifestFiles.Add($file)
}
$catalogManifestCount = 0
$catalogManifestDirectory = Join-Path $manifestDirectoryFull "catalog"
if (Test-Path -LiteralPath $catalogManifestDirectory -PathType Container) {
    foreach ($file in @(Get-ChildItem -LiteralPath $catalogManifestDirectory -Filter "*.json" -File | Sort-Object Name)) {
        $manifestFiles.Add($file)
        $catalogManifestCount++
    }
}

if ($manifestFiles.Count -eq 0) { throw "No feature acceptance manifests were found." }

foreach ($file in $manifestFiles) {
    $displayName = if ([IO.Path]::GetDirectoryName($file.FullName) -eq $catalogManifestDirectory) {
        "catalog/$($file.Name)"
    }
    else { $file.Name }
    try { $manifest = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json }
    catch { throw "Invalid JSON in '$displayName': $($_.Exception.Message)" }

    if ($manifest.schemaVersion -ne 1) { throw "${displayName}: schemaVersion must be 1." }
    if ([string]::IsNullOrWhiteSpace($manifest.featureId)) { throw "${displayName}: featureId is required." }
    if (-not $featureIds.Add([string]$manifest.featureId)) { throw "Duplicate featureId '$($manifest.featureId)'." }
    if ([string]::IsNullOrWhiteSpace($manifest.title)) { throw "${displayName}: title is required." }
    if (@($manifest.sourcePaths).Count -eq 0) { throw "${displayName}: sourcePaths must not be empty." }
    if (@($manifest.scenarios).Count -eq 0) { throw "${displayName}: scenarios must not be empty." }
    # Catalog manifests identify the stable visible Feature ID; strategic manifests may name the
    # visible Feature IDs a shared scenario covers.
    if ($null -ne $manifest.catalogFeatureIds) {
        foreach ($catalogFeatureId in @($manifest.catalogFeatureIds)) {
            if ([string]::IsNullOrWhiteSpace([string]$catalogFeatureId)) {
                throw "${displayName}: catalogFeatureIds must not contain empty entries."
            }
        }
    }

    foreach ($scenario in @($manifest.scenarios)) {
        if ([string]::IsNullOrWhiteSpace($scenario.id)) { throw "$($file.Name): every scenario needs an id." }
        if (-not $scenarioIds.Add([string]$scenario.id)) { throw "Duplicate scenario id '$($scenario.id)'." }
        if ([string]::IsNullOrWhiteSpace($scenario.title)) { throw "$($file.Name)/$($scenario.id): title is required." }
        if ($allowedTiers -notcontains [string]$scenario.tier) { throw "$($file.Name)/$($scenario.id): invalid tier '$($scenario.tier)'." }
        if ($allowedStatuses -notcontains [string]$scenario.automationStatus) { throw "$($file.Name)/$($scenario.id): invalid automationStatus '$($scenario.automationStatus)'." }
        if (@($scenario.changedPathPatterns).Count -eq 0) { throw "$($file.Name)/$($scenario.id): changedPathPatterns must not be empty." }
        if ([string]::IsNullOrWhiteSpace($scenario.expected)) { throw "$($file.Name)/$($scenario.id): expected is required." }

        $automatedTier = @("core-automated", "office-automated") -contains [string]$scenario.tier
        if ($automatedTier -and $scenario.automationStatus -eq "ready") {
            if ($null -eq $scenario.runner) { throw "$($file.Name)/$($scenario.id): ready automated scenario requires runner." }
            $runnerPath = Join-Path $repositoryRoot ([string]$scenario.runner.path -replace '/', '\')
            if (-not (Test-Path -LiteralPath $runnerPath -PathType Leaf)) {
                if ([string]$scenario.tier -eq "office-automated") {
                    $unavailableOfficeRunners++
                } else {
                    throw "$($file.Name)/$($scenario.id): core runner path does not exist: $($scenario.runner.path)"
                }
            }
            if (@("powershell-script", "dotnet-test") -notcontains [string]$scenario.runner.kind) {
                throw "$($file.Name)/$($scenario.id): unsupported runner kind '$($scenario.runner.kind)'."
            }

            if ([string]$scenario.runner.kind -eq "dotnet-test") {
                # Review round 6g/6h: a scenario must not claim that a test project asserts something while its runner
                # runs a different one. Measured defeats: the -002 scenario named oPenEfficiency.Core.Tests in its
                # expected text while its runner pointed at another project (6g), and the sentinel scenario was
                # defeated by stripping the project name out of its text and repointing the runner (6h) - so the
                # machine-readable changedPathPatterns decide which project the scenario owns, and the claim text may
                # only name that project.
                $runnerProject = [IO.Path]::GetFileNameWithoutExtension([string]$scenario.runner.path)
                $ownedTestsProjects = @()
                foreach ($pattern in @($scenario.changedPathPatterns)) {
                    $text = [string]$pattern
                    if ($text -notmatch "[*\?]" -and $text -match "Tests\.cs$" -and $text -notmatch "^oPenEfficiency(?:\.[A-Za-z]+)*\.Tests/") {
                        throw "$($file.Name)/$($scenario.id): the changedPathPattern '$text' names a test file without its test project, so nothing ties the scenario to the project it proves (review round 6h)."
                    }
                    $match = [regex]::Match($text, "^(oPenEfficiency(?:\.[A-Za-z]+)*\.Tests)/")
                    if ($match.Success) { $ownedTestsProjects += $match.Groups[1].Value }
                }
                foreach ($owned in ($ownedTestsProjects | Select-Object -Unique)) {
                    if ($owned -ne $runnerProject) {
                        throw "$($file.Name)/$($scenario.id): its changedPathPatterns own the test project '$owned' while its runner runs '$runnerProject' - a weakened runner must not keep a stronger claim."
                    }
                }
                $claimText = ("{0} {1}" -f ([string]$scenario.expected), ([string]$scenario.notes))
                foreach ($claim in [regex]::Matches($claimText, "oPenEfficiency(?:\.[A-Za-z]+)*\.Tests")) {
                    if ($claim.Value -ne $runnerProject) {
                        throw "$($file.Name)/$($scenario.id): the scenario claims '$($claim.Value)' while its runner runs '$runnerProject' - a weakened runner must not keep a stronger claim."
                    }
                }
            }
        }

        if (@("owner-manual", "visual-review") -contains [string]$scenario.tier) {
            if (@($scenario.ownerSteps).Count -eq 0) {
                throw "$($file.Name)/$($scenario.id): manual/visual scenario requires ownerSteps."
            }
            if ([string]$scenario.tier -eq "owner-manual" -and
                [string]::IsNullOrWhiteSpace([string]$scenario.technicalReason)) {
                throw "$($file.Name)/$($scenario.id): an owner-manual scenario requires a technicalReason that records why deterministic automation is not feasible."
            }
        }
    }
}

# Review round 6h: the IL binding test asserts against a member list that lives in a data file, so that list has to be
# pinned outside the test as well - otherwise extending the list inside the test re-legalizes a guard helper, and no
# gate notices. Regenerate the pin (OPENEFFICIENCY_DISTRIBUTE_GUARD_PIN=1) and update this digest with the change.
$guardPinPath = Join-Path $repositoryRoot "docs/testing/feature-acceptance/distribute-guard-pin.txt"
if (-not (Test-Path -LiteralPath $guardPinPath -PathType Leaf)) {
    throw "The distribute guard pin docs/testing/feature-acceptance/distribute-guard-pin.txt is missing."
}
$guardPinExpected = "cc052e512bd8611bbe93005310f859dade731596a8aa445f2f2e9de59f96b627"
$guardPinText = ([IO.File]::ReadAllText($guardPinPath)) -replace "`r`n", "`n"
$guardPinHash = [BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes($guardPinText))).Replace("-", "").ToLowerInvariant()
if ($guardPinHash -ne $guardPinExpected) {
    throw "docs/testing/feature-acceptance/distribute-guard-pin.txt changed (sha256 $guardPinHash, pinned $guardPinExpected) - regenerate it with OPENEFFICIENCY_DISTRIBUTE_GUARD_PIN=1 and update this digest."
}

Write-Output ("FEATURE_ACCEPTANCE_MANIFESTS_PASS manifests={0} catalog_manifests={1} scenarios={2}" -f $manifestFiles.Count, $catalogManifestCount, $scenarioIds.Count)

Write-Output ("PUBLIC_SOURCE_OFFICE_TOOLING_UNAVAILABLE count={0}; no Office or release acceptance claimed" -f $unavailableOfficeRunners)
