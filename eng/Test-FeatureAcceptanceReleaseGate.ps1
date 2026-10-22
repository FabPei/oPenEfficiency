[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ResultsPath
)

# QA-S11 release gate over one acceptance results file: every required scenario (ready, in an
# automatable tier) must actually have run and passed. A required scenario that is Failed or
# NotRun blocks the release; visual-review and owner-manual evidence is human work and is
# reported by eng/Get-FeatureAcceptanceStatus.ps1, never judged here.

$ErrorActionPreference = "Stop"
$resolvedResultsPath = [IO.Path]::GetFullPath($ResultsPath)
if (-not (Test-Path -LiteralPath $resolvedResultsPath -PathType Leaf)) {
    throw "No acceptance results file at $resolvedResultsPath. Run eng/Invoke-FeatureAcceptance.ps1 first."
}

$result = ([IO.File]::ReadAllText($resolvedResultsPath, [Text.UTF8Encoding]::new($false))) | ConvertFrom-Json
$automatableTiers = @("core-automated", "office-automated")

$required = @($result.Scenarios | Where-Object {
    $_.AutomationStatus -eq "ready" -and $_.Tier -in $automatableTiers })
$failed = @($required | Where-Object { $_.Outcome -eq "Failed" })
$notRun = @($required | Where-Object { $_.Outcome -eq "NotRun" })
$passed = @($required | Where-Object { $_.Outcome -eq "Passed" })

Write-Output ("FEATURE_ACCEPTANCE_RELEASE_GATE required={0} passed={1} failed={2} not_run={3} include_office={4} captured={5}" -f `
    $required.Count, $passed.Count, $failed.Count, $notRun.Count, $result.IncludeOffice, $result.CapturedUtc)

if ($failed.Count -gt 0) {
    Write-Output ("FEATURE_ACCEPTANCE_RELEASE_GATE_BLOCKED reason=failed first={0}" -f $failed[0].Id)
    exit 1
}
if ($notRun.Count -gt 0) {
    Write-Output ("FEATURE_ACCEPTANCE_RELEASE_GATE_BLOCKED reason=not_run first={0}" -f $notRun[0].Id)
    exit 1
}
if ($required.Count -eq 0) {
    # No required scenario exists in this results file: that is missing evidence, not success.
    Write-Output "FEATURE_ACCEPTANCE_RELEASE_GATE_BLOCKED reason=no_required_scenario_evidence"
    exit 1
}

Write-Output "FEATURE_ACCEPTANCE_RELEASE_GATE_APPROVED required_scenarios=all_run"
exit 0
