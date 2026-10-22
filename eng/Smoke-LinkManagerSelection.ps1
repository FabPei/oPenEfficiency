[CmdletBinding()]
param(
    [string]$ResultsPath
)

$ErrorActionPreference = "Stop"

# DL-LINKMGR-001: the Link Manager's selection contract on the real WPF window. No Office host is
# involved. The rows are fixtures; everything that matters is production code - the DataGrid's
# selection machinery, the row template's CheckBox driven through its toggle automation peer, the
# production "Refresh Selected" handler, and the selected-scope resolution every command uses.

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA) {
    throw "This smoke drives a WPF window and must run on an STA thread (invoke with powershell -STA)."
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$officePath = Join-Path $env:USERPROFILE ".nuget\packages\microsoftofficecore\15.0.0\lib\net35\Office.dll"
$piaPath = Join-Path $env:USERPROFILE ".nuget\packages\microsoft.office.interop.powerpoint\15.0.4420.1018\lib\net20\Microsoft.Office.Interop.PowerPoint.dll"
$coreDll = Join-Path $repoRoot "oPenEfficiency.Core\bin\Release\oPenEfficiency.Core.dll"
$addInDll = Join-Path $repoRoot "oPenEfficiency\bin\Release\oPenEfficiency.dll"

$failureStage = "load-assemblies"
$startedUtc = [DateTime]::UtcNow
$observed = [ordered]@{}
$resultWritten = $false

function Write-SmokeResult([bool]$Passed, [object]$FailureDiagnostics) {
    if ($script:resultWritten -or [string]::IsNullOrWhiteSpace($ResultsPath)) { return }
    $path = [IO.Path]::GetFullPath($ResultsPath)
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path)) | Out-Null
    $payload = [ordered]@{
        SchemaVersion = 1
        CapturedUtc = [DateTime]::UtcNow.ToString("o")
        Passed = $Passed
        Stage = $script:failureStage
        Observed = [pscustomobject]$script:observed
    }
    if ($null -ne $FailureDiagnostics) { $payload.Failure = $FailureDiagnostics }
    [pscustomobject]$payload | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $path -Encoding UTF8
    $script:resultWritten = $true
}

function Assert-Smoke([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

function Assert-Equal([string]$Expected, [string]$Actual, [string]$Message) {
    if ($Expected -ne $Actual) { throw "$Message (expected='$Expected' actual='$Actual')" }
}

function Pump-Ui([object]$Window) {
    # Drains everything the layout/render/DataGrid machinery queued, without a message loop.
    $Window.Dispatcher.Invoke([Action]{ }, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle)
}

function New-FixtureLink([string]$Name, [string]$Status) {
    $item = [oPenEfficiency.UI.ExcelLinkItem]::new()
    $item.SlideNumber = 1
    $item.ShapeName = $Name
    $item.Kind = "Table"
    $item.SourcePath = "C:\fixtures\$Name.xlsx"
    $item.Sheet = "Sheet1"
    $item.Range = "A1:B2"
    $item.Mode = "Values"
    $item.Status = $Status
    $item.LastRefresh = "-"
    $item.IsNativeTable = $true
    $item.ShapeRef = $null
    return $item
}

$window = $null
try {
    Add-Type -Path $officePath
    Add-Type -Path $piaPath
    Add-Type -Path $coreDll
    Add-Type -Path $addInDll

    $failureStage = "create-window"
    # The host-free constructor keeps LoadLinks() away from a live deck; its own try/catch then leaves
    # the collection empty, and the fixtures take over.
    $window = [oPenEfficiency.UI.ExcelLinkManagerWindow]::new()
    $window.ShowInTaskbar = $false
    $window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
    $window.Left = -20000
    $window.Top = -20000
    $window.Show()
    Pump-Ui $window
    Assert-Smoke ($window.Links.Count -eq 0) "The window loaded $($window.Links.Count) links, but the smoke expects an empty production state before the fixtures."

    $failureStage = "load-fixtures"
    [oPenEfficiency.UI.ExcelLinkItem[]]$fixtures = @(
        (New-FixtureLink "ShapeLinked70" "Linked"),
        (New-FixtureLink "ShapeMissing71" "Missing"),
        (New-FixtureLink "ShapeLinked72" "Linked"),
        (New-FixtureLink "ShapeMissing73" "Missing"),
        (New-FixtureLink "ShapeLinked74" "Linked")
    )
    $window.LoadFixtureLinksForAutomation($fixtures)
    Pump-Ui $window
    Assert-Equal "5" ([string]$window.PresentedRowCountForAutomation()) "The grid does not present every fixture row."
    Assert-Equal "" ([string]$window.DescribeSelectedScopeForAutomation()) "The scope is not empty before anything is selected."
    $observed.Baseline = "presented=$($window.PresentedRowCountForAutomation());scope=<empty>"

    # --- row click -> scope, and the row selection mirrors into the check box ---
    $failureStage = "row-click-scope"
    $window.SelectRowsForAutomation([int[]]@(0), $false)
    Pump-Ui $window
    Assert-Equal "0:True" ([string]$window.DescribeSelectedScopeForAutomation()) "A single row click must deliver exactly that row to a selected-scope command."
    Assert-Smoke ($fixtures[0].IsSelected) "The row click did not mirror into the row's check box."
    $observed.RowClick = "scope=$($window.DescribeSelectedScopeForAutomation());checkboxRow0=$($fixtures[0].IsSelected)"

    $failureStage = "row-click-additive"
    $window.SelectRowsForAutomation([int[]]@(1), $true)
    Pump-Ui $window
    Assert-Equal "0:True,1:True" ([string]$window.DescribeSelectedScopeForAutomation()) "An additive row click must extend the scope."
    $observed.RowClickAdditive = "scope=$($window.DescribeSelectedScopeForAutomation())"

    # --- check box toggle through its automation peer (no synthetic mouse input) ---
    $failureStage = "checkbox-toggle"
    $toggled = $window.ToggleCheckboxForAutomation(2)
    Pump-Ui $window
    Assert-Smoke $toggled "The row's check box did not become checked through its toggle automation peer."
    Assert-Smoke ($fixtures[2].IsSelected) "Toggling the check box did not set the row's selection state."
    Assert-Equal "0:True,1:True,2:True" ([string]$window.DescribeSelectedScopeForAutomation()) "The check box toggle did not reach the scope."
    $observed.CheckboxToggle = "checked=$toggled;scope=$($window.DescribeSelectedScopeForAutomation())"

    $failureStage = "multi-select"
    $window.SelectAllRowsForAutomation()
    Pump-Ui $window
    Assert-Equal "0:True,1:True,2:True,3:True,4:True" ([string]$window.DescribeSelectedScopeForAutomation()) "Multi-select did not deliver every presented row."
    $observed.MultiSelect = "scope=$($window.DescribeSelectedScopeForAutomation())"

    $failureStage = "unselect"
    $window.UnselectRowsForAutomation([int[]]@(1, 3))
    Pump-Ui $window
    Assert-Equal "0:True,2:True,4:True" ([string]$window.DescribeSelectedScopeForAutomation()) "Unselecting rows did not remove them from the scope."
    Assert-Smoke (-not $fixtures[1].IsSelected -and -not $fixtures[3].IsSelected) "Unselecting a row did not clear its check box state."
    $observed.Unselect = "scope=$($window.DescribeSelectedScopeForAutomation())"

    # --- the production handler receives exactly that scope (it stamps the rows it acted on) ---
    $failureStage = "command-scope"
    $touched = $window.InvokeRefreshSelectedForAutomation()
    Pump-Ui $window
    Assert-Equal "0,2,4" ([string]$touched) "Refresh Selected did not act on exactly the selected rows."
    $observed.CommandScope = "touched=$touched"

    # --- the filtered view: a scoped command can only ever act on presented rows ---
    $failureStage = "filtered-view-scope"
    $window.ApplyBrokenFilterForAutomation()
    Pump-Ui $window
    $presentedAfterFilter = [string]$window.PresentedRowCountForAutomation()
    $hiddenChecked = [int]$window.CheckedButHiddenCountForAutomation()
    $scopeAfterFilter = [string]$window.DescribeSelectedScopeForAutomation()
    $observed.FilteredView = "presented=$presentedAfterFilter;presentedRows=$($window.DescribePresentedRowsForAutomation());checkedButHidden=$hiddenChecked;scope='$scopeAfterFilter'"
    Assert-Equal "2" $presentedAfterFilter "The broken-links filter did not reduce the presented rows."
    Assert-Equal "1:ShapeMissing71,3:ShapeMissing73" ([string]$window.DescribePresentedRowsForAutomation()) "The broken-links filter presented the wrong rows."
    # Invariant: every row in the scope is a row the grid presents.
    foreach ($entry in @($scopeAfterFilter -split ',')) {
        if ([string]::IsNullOrWhiteSpace($entry)) { continue }
        $index = [int]($entry.Split(':')[0])
        Assert-Smoke (@(1, 3) -contains $index) "The scope contains row $index, which the filtered view does not present."
    }
    $touchedFiltered = $window.InvokeRefreshSelectedForAutomation()
    Pump-Ui $window
    foreach ($index in @($touchedFiltered -split ',')) {
        if ([string]::IsNullOrWhiteSpace($index)) { continue }
        Assert-Smoke (@(1, 3) -contains [int]$index) "Refresh Selected acted on row $index, which the filtered view does not present."
    }
    $observed.FilteredViewCommand = "touched='$touchedFiltered'"

    $failureStage = "restore-view"
    $window.ShowAllRowsForAutomation()
    Pump-Ui $window
    Assert-Equal "5" ([string]$window.PresentedRowCountForAutomation()) "Restoring the unfiltered view did not present every row again."
    $observed.Restored = "presented=$($window.PresentedRowCountForAutomation())"

    Write-SmokeResult -Passed $true -FailureDiagnostics $null
    Write-Output "LINK_MANAGER_SELECTION_PASS row_click_scope=true row_click_mirrors_checkbox=true checkbox_toggle_scope=true multi_select_scope=true unselect_removes_from_scope=true command_receives_exact_scope=true filtered_view_scope_presented_only=true presented_after_filter=2 restored_rows=5"
}
catch {
    $originalFailure = $_
    $failureDiagnostics = [pscustomobject]@{
        ScriptLine = $originalFailure.InvocationInfo.ScriptLineNumber
        Message = $originalFailure.Exception.Message
        StackTrace = $originalFailure.ScriptStackTrace
        UptimeSeconds = [Math]::Round(([DateTime]::UtcNow - $startedUtc).TotalSeconds, 1)
    }
    Write-SmokeResult -Passed $false -FailureDiagnostics $failureDiagnostics
    Write-Error ("Link Manager selection smoke failed at stage '{0}', line {1}: {2}`n{3}" -f
        $failureStage, $originalFailure.InvocationInfo.ScriptLineNumber, $originalFailure.Exception.Message,
        $originalFailure.ScriptStackTrace)
    throw $originalFailure
}
finally {
    if ($window) {
        try { $window.Close() } catch { Write-Warning "Could not close the smoke window: $($_.Exception.Message)" }
    }
}
