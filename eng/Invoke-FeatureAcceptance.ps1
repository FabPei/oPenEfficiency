[CmdletBinding()]
param(
    [string[]]$FeatureId,
    [string[]]$ChangedPath,
    [switch]$Plan,
    [switch]$IncludeOffice,
    [ValidateSet("Changed", "Nightly", "Release")]
    [string]$Profile = "Changed",
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [string]$ResultsPath
)

# QA-S11 execution profiles. "Changed" is the historical behavior: Core plus the affected
# ready scenarios, Office only when explicitly requested. "Nightly" runs the complete ready
# catalog including the Office legs and defaults to the nightly evidence file. "Release" is
# the nightly run plus the rule that a required (ready) scenario left NotRun blocks the
# release instead of being counted as success.
$includeOfficeEffective = $IncludeOffice -or ($Profile -ne "Changed")

$ErrorActionPreference = "Stop"
$repositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
if ([string]::IsNullOrWhiteSpace($ResultsPath)) {
    if ($Profile -eq "Changed") {
        $ResultsPath = Join-Path $repositoryRoot "TestResults\feature-acceptance.json"
    } else {
        $ResultsPath = Join-Path $repositoryRoot "TestResults\feature-acceptance-nightly.json"
    }
}
$resolvedResultsPath = [IO.Path]::GetFullPath($ResultsPath)
$resultsDirectory = [IO.Path]::GetDirectoryName($resolvedResultsPath)
[IO.Directory]::CreateDirectory($resultsDirectory) | Out-Null
$planPath = Join-Path $resultsDirectory "feature-acceptance-plan.json"

function Find-MSBuild {
    $command = Get-Command "msbuild" -ErrorAction SilentlyContinue
    if ($null -ne $command -and
        -not [string]::IsNullOrWhiteSpace($command.Source)) {
        return $command.Source
    }

    if (-not [string]::IsNullOrWhiteSpace($env:MSBUILD) -and
        (Test-Path -LiteralPath $env:MSBUILD -PathType Leaf)) {
        return [IO.Path]::GetFullPath($env:MSBUILD)
    }

    $programFilesX86 = [Environment]::GetFolderPath(
        [Environment+SpecialFolder]::ProgramFilesX86)
    $vswhereCandidates = @(
        (Join-Path $programFilesX86 "Microsoft Visual Studio\Installer\vswhere.exe"),
        (Join-Path $repositoryRoot "eng\tools\vswhere.exe")
    )

    foreach ($vswherePath in $vswhereCandidates) {
        if (-not (Test-Path -LiteralPath $vswherePath -PathType Leaf)) {
            continue
        }

        $match = @(& $vswherePath `
            -latest `
            -products * `
            -requires Microsoft.Component.MSBuild `
            -find "MSBuild\**\Bin\MSBuild.exe") |
            Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
            Select-Object -First 1
        if (-not [string]::IsNullOrWhiteSpace($match)) {
            return $match
        }
    }

    throw "MSBuild was not found. Install Visual Studio with the Office development tools or add MSBuild to PATH."
}

& (Join-Path $PSScriptRoot "Test-FeatureAcceptanceManifests.ps1") | Out-Null
$planArgs = @{}
if ($FeatureId) { $planArgs.FeatureId = $FeatureId }
if ($ChangedPath) { $planArgs.ChangedPath = $ChangedPath }
$planArgs.ResultsPath = $planPath
& (Join-Path $PSScriptRoot "Get-FeatureAcceptancePlan.ps1") @planArgs | Out-Null
$acceptancePlan = Get-Content -LiteralPath $planPath -Raw | ConvertFrom-Json

$scenarioResults = [Collections.Generic.List[object]]::new()
$overallPassed = $true
# Every distinct runner project is built before the first scenario executes. Building lazily inside the loop
# let a runner that had already executed keep oPenEfficiency.Core.dll loaded while the next project's MSBuild
# tried to copy that file, so the run went red with MSB3027/MSB3021 - a build failure that no feature
# behaviour could explain and that no assertion reported.
function Test-ScenarioExecutes {
    param($Scenario)

    if ($Plan) { return $false }
    if ($Scenario.Tier -in @("owner-manual", "visual-review")) { return $false }
    if ($Scenario.AutomationStatus -ne "ready") { return $false }
    if ($Scenario.Tier -eq "office-automated" -and -not $includeOfficeEffective) { return $false }
    return $true
}

$runnerProjects = [Collections.Generic.List[string]]::new()
foreach ($plannedScenario in @($acceptancePlan.Scenarios)) {
    if (-not (Test-ScenarioExecutes -Scenario $plannedScenario)) { continue }
    if ($plannedScenario.Runner.kind -ne "dotnet-test") { continue }

    $plannedProject = [IO.Path]::GetFullPath((Join-Path $repositoryRoot ([string]$plannedScenario.Runner.path -replace '/', '\')))
    if (-not $runnerProjects.Contains($plannedProject)) { $runnerProjects.Add($plannedProject) }
}
if ($runnerProjects.Count -gt 0) {
    $msbuildPath = Find-MSBuild
    foreach ($runnerProject in $runnerProjects) {
        & $msbuildPath $runnerProject /restore /t:Build "/p:Configuration=$Configuration" /p:Platform=AnyCPU /v:minimal /nologo
        if ($LASTEXITCODE -ne 0) {
            throw ("MSBuild exited with code $LASTEXITCODE while building '$([IO.Path]::GetFileName($runnerProject))'.")
        }
    }
}

foreach ($scenario in @($acceptancePlan.Scenarios)) {
    $outcome = "NotRun"
    $details = ""
    $started = Get-Date

    if ($Plan) {
        $details = "Plan only."
    }
    elseif ($scenario.Tier -in @("owner-manual", "visual-review")) {
        $details = "Owner/reviewer action required."
    }
    elseif ($scenario.AutomationStatus -ne "ready") {
        $details = "Automation is $($scenario.AutomationStatus)."
    }
    elseif ($scenario.Tier -eq "office-automated" -and -not $includeOfficeEffective) {
        $details = "Office automation is ready but was not requested. Re-run with -IncludeOffice on an isolated Office test environment."
    }
    else {
        try {
            $runner = $scenario.Runner
            if ($runner.kind -eq "powershell-script") {
                $scriptPath = [IO.Path]::GetFullPath((Join-Path $repositoryRoot ([string]$runner.path -replace '/', '\')))
                # Named binding: array splatting passes its elements POSITIONALLY, so a manifest argument
                # token like '-Commands' arrives as a literal value and the runner's parameter binder rejects
                # it ("not in the ValidateSet" - measured 2026-10-19, when the first scaled office run failed
                # every office leg at bind time before a single smoke body ran; REG-ACCEPTANCE-ARGS-001). The
                # tokens are parsed into named parameters and splatted as a hashtable instead; a token that is
                # not a parameter name fails loudly instead of binding positionally and under-testing the
                # scenario. Parameter tokens carry zero values (switch), one value, or a value list.
                # An omitted argument list must produce zero tokens, not one empty string.
                $argumentTokens = @(foreach ($argument in $runner.arguments) { [string]$argument })
                $namedArguments = @{}
                for ($a = 0; $a -lt $argumentTokens.Count; $a++) {
                    if ($argumentTokens[$a] -cnotmatch '^-[A-Za-z][A-Za-z0-9]*$') {
                        throw ("Runner arguments for '$([IO.Path]::GetFileName($scriptPath))' must be named parameter tokens; the positional value '$($argumentTokens[$a])' cannot be bound safely.")
                    }
                    $parameterName = $argumentTokens[$a].Substring(1)
                    $parameterValues = @()
                    while (($a + 1) -lt $argumentTokens.Count -and $argumentTokens[$a + 1] -cnotmatch '^-') {
                        $a++
                        $parameterValues += $argumentTokens[$a]
                    }
                    if ($parameterValues.Count -eq 0) { $namedArguments[$parameterName] = $true }
                    elseif ($parameterValues.Count -eq 1) { $namedArguments[$parameterName] = $parameterValues[0] }
                    else { $namedArguments[$parameterName] = [string[]]$parameterValues }
                }
                # A runner script that ends in a failing native command or calls 'exit 1' must not be reported as
                # a passed assertion, so its own exit state is checked as well.
                $LASTEXITCODE = 0
                & $scriptPath @namedArguments
                if ($LASTEXITCODE -ne 0) {
                    throw ("Runner script '$([IO.Path]::GetFileName($scriptPath))' exited with code $LASTEXITCODE.")
                }
            }
            elseif ($runner.kind -eq "dotnet-test") {
                $projectPath = [IO.Path]::GetFullPath((Join-Path $repositoryRoot ([string]$runner.path -replace '/', '\')))
                $dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
                if ($null -eq $dotnet) { throw "dotnet was not found and is required to run acceptance tests." }
                # Neither a filter that matches nothing nor a project without tests may report Passed: both
                # make 'dotnet test' exit 0 without executing anything. The test list is therefore required - with
                # the filter when one is declared, and in full when there is none - and parsed instead of localized
                # runner output.
                $listArgs = @("test", $projectPath, "--configuration", $Configuration, "--no-build", "--no-restore", "--list-tests")
                if (-not [string]::IsNullOrWhiteSpace([string]$runner.filter)) {
                    $listArgs += @("--filter", [string]$runner.filter)
                }
                $listed = @(& $dotnet.Source @listArgs 2>$null)
                $listedTests = @($listed | Where-Object { $_ -match "^\s{2,}\S" })
                if ($listedTests.Count -eq 0) {
                    $subject = $(if ([string]::IsNullOrWhiteSpace([string]$runner.filter)) {
                        "The runner project '$([IO.Path]::GetFileName($projectPath))' declares no test" } else {
                        "The filter '$($runner.filter)' matches no test in '$([IO.Path]::GetFileName($projectPath))'" })
                    throw ("${subject}: 'dotnet test' would exit 0 without executing anything, so this scenario must " +
                        "not report Passed.")
                }
                $args = @("test", $projectPath, "--configuration", $Configuration, "--no-build", "--no-restore", "--logger", "console;verbosity=minimal")
                if (-not [string]::IsNullOrWhiteSpace([string]$runner.filter)) {
                    $args += @("--filter", [string]$runner.filter)
                }
                & $dotnet.Source @args
                if ($LASTEXITCODE -ne 0) { throw "dotnet test exited with code $LASTEXITCODE." }
            }
            else {
                throw "Unsupported runner kind '$($runner.kind)'."
            }
            $outcome = "Passed"
            $details = "Runner completed successfully."
        }
        catch {
            $outcome = "Failed"
            $details = $_.Exception.Message
            $overallPassed = $false
        }
    }

    $scenarioResults.Add([pscustomobject]@{
        FeatureId = $scenario.FeatureId
        Id = $scenario.Id
        Title = $scenario.Title
        Tier = $scenario.Tier
        AutomationStatus = $scenario.AutomationStatus
        Outcome = $outcome
        DurationSeconds = [Math]::Round(((Get-Date) - $started).TotalSeconds, 3)
        Expected = $scenario.Expected
        Details = $details
        OwnerSteps = @($scenario.OwnerSteps)
    })
}

$result = [pscustomobject]@{
    SchemaVersion = 1
    CapturedUtc = [DateTime]::UtcNow.ToString("o")
    PlanOnly = [bool]$Plan
    IncludeOffice = [bool]$includeOfficeEffective
    Profile = $Profile
    Passed = $overallPassed -and @($scenarioResults | Where-Object Outcome -eq "Failed").Count -eq 0
    ScenarioCount = $scenarioResults.Count
    PassedCount = @($scenarioResults | Where-Object Outcome -eq "Passed").Count
    FailedCount = @($scenarioResults | Where-Object Outcome -eq "Failed").Count
    NotRunCount = @($scenarioResults | Where-Object Outcome -eq "NotRun").Count
    Scenarios = @($scenarioResults)
}
$result | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $resolvedResultsPath -Encoding UTF8

$markdownPath = [IO.Path]::ChangeExtension($resolvedResultsPath, ".md")
$lines = [Collections.Generic.List[string]]::new()
$lines.Add("# Feature Acceptance Report")
$lines.Add("")
$lines.Add(("- Passed: {0}; Failed: {1}; NotRun: {2}" -f $result.PassedCount, $result.FailedCount, $result.NotRunCount))
$lines.Add("")
foreach ($scenario in @($scenarioResults)) {
    $lines.Add(("## {0} - {1}" -f $scenario.Id, $scenario.Outcome))
    $lines.Add("")
    $lines.Add(("**{0}**  " -f $scenario.Title))
    $lines.Add(("Tier: `{0}`; automation: `{1}`  " -f $scenario.Tier, $scenario.AutomationStatus))
    $lines.Add(("Expected: {0}  " -f $scenario.Expected))
    $lines.Add(("Result: {0}" -f $scenario.Details))
    if (@($scenario.OwnerSteps).Count -gt 0) {
        $lines.Add("")
        $lines.Add("Owner steps:")
        $stepIndex = 1
        foreach ($step in @($scenario.OwnerSteps)) {
            $lines.Add(("{0}. {1}" -f $stepIndex, $step))
            $stepIndex++
        }
    }
    $lines.Add("")
}
$lines | Set-Content -LiteralPath $markdownPath -Encoding UTF8

Write-Output ("FEATURE_ACCEPTANCE_DONE passed={0} failed={1} not_run={2} results={3} report={4}" -f
    $result.PassedCount, $result.FailedCount, $result.NotRunCount, $resolvedResultsPath, $markdownPath)
if ($result.FailedCount -gt 0) { exit 1 }

# QA-S11 release rule: a required (ready, automatable-tier) scenario that did not run blocks
# the release instead of being counted as success. The rule lives in its own gate script so
# it can be tested directly against synthetic results files.
if ($Profile -eq "Release" -and -not $Plan) {
    & (Join-Path $PSScriptRoot "Test-FeatureAcceptanceReleaseGate.ps1") -ResultsPath $resolvedResultsPath
    if ($LASTEXITCODE -ne 0) { exit 1 }
}
