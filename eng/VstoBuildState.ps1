# Preserve only the two add-in keys and VSTO records belonging to this source checkout.
function Read-BuildRegistryKey([string]$Path) {
    $ErrorActionPreference = 'Stop'
    $exists = Test-Path -LiteralPath $Path
    $values = @()
    if ($exists) {
        $key = Get-Item -LiteralPath $Path
        if ($Path -match '\\VSTO\\(SolutionMetadata|Security\\Inclusion)\\' -and $key.GetSubKeyNames().Length -gt 0) {
            throw 'Unexpected nested VSTO checkout metadata; refusing an incomplete snapshot.'
        }
        foreach ($name in ($key.GetValueNames() | Sort-Object)) {
            $values += [pscustomobject]@{ Name=$name; Kind=$key.GetValueKind($name).ToString(); Value=$key.GetValue($name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames) }
        }
    }
    return [pscustomobject]@{ Path=$Path; Exists=$exists; Values=$values }
}
function Write-BuildRegistryKey($Snapshot) {
    $ErrorActionPreference = 'Stop'
    $current = Read-BuildRegistryKey $Snapshot.Path
    if (($current | ConvertTo-Json -Depth 12 -Compress) -ceq ($Snapshot | ConvertTo-Json -Depth 12 -Compress)) { return }
    if (-not $Snapshot.Exists) {
        if (Test-Path -LiteralPath $Snapshot.Path) {
            # This is one of the exact captured add-in/checkout-owned keys, never a computed parent hive.
            Remove-Item -LiteralPath $Snapshot.Path -Recurse -Force
        }
        return
    }
    if (-not (Test-Path -LiteralPath $Snapshot.Path)) { New-Item -Path $Snapshot.Path -Force | Out-Null }
    $key = Get-Item -LiteralPath $Snapshot.Path
    foreach ($name in @(($key.GetValueNames() | Sort-Object))) { Remove-ItemProperty -LiteralPath $Snapshot.Path -Name $name }
    foreach ($value in $Snapshot.Values) {
        New-ItemProperty -LiteralPath $Snapshot.Path -Name $value.Name -Value $value.Value -PropertyType $value.Kind -Force | Out-Null
    }
}
function Get-VstoBuildState([string]$RepositoryRoot) {
    $ErrorActionPreference = 'Stop'
    $prefix = 'file:///' + [IO.Path]::GetFullPath($RepositoryRoot).Replace('\','/').TrimEnd('/') + '/'
    $encodedPrefix = ([Uri]([IO.Path]::GetFullPath($RepositoryRoot).TrimEnd('\') + '\')).AbsoluteUri
    $metadataPath = 'HKCU:\Software\Microsoft\VSTO\SolutionMetadata'
    $inclusionPath = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
    $links = @()
    $entries = @()
    if (Test-Path -LiteralPath $metadataPath) {
        $root = Get-Item -LiteralPath $metadataPath
        foreach ($name in ($root.GetValueNames() | Sort-Object)) {
            if (($name.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) -or $name.StartsWith($encodedPrefix, [StringComparison]::OrdinalIgnoreCase))) {
                $id = [string]$root.GetValue($name)
                if ($id -notmatch '^\{[0-9a-fA-F-]{36}\}$') { throw 'Unexpected VSTO metadata key ID.' }
                $links += [pscustomobject]@{ Name=$name; Id=$id }
                $entries += Read-BuildRegistryKey (Join-Path $metadataPath $id)
            }
        }
    }
    $inclusions = @()
    if (Test-Path -LiteralPath $inclusionPath) {
        foreach ($key in (Get-ChildItem -LiteralPath $inclusionPath | Sort-Object Name)) {
            $url = [string]$key.GetValue('Url')
            if (($url.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) -or $url.StartsWith($encodedPrefix, [StringComparison]::OrdinalIgnoreCase))) {
                $inclusions += Read-BuildRegistryKey $key.PSPath
            }
        }
    }
    return [pscustomobject]@{
        RepositoryRoot=$RepositoryRoot; Prefix=$prefix; EncodedPrefix=$encodedPrefix; MetadataPath=$metadataPath; InclusionPath=$inclusionPath
        Links=$links; Entries=$entries; Inclusions=$inclusions
        AddIns=@(
            (Read-BuildRegistryKey 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\oPenEfficiency'),
            (Read-BuildRegistryKey 'HKCU:\Software\Microsoft\Office\Excel\Addins\oPenEfficiency.Excel')
        )
    }
}
function Restore-VstoBuildState($Snapshot) {
    $ErrorActionPreference = 'Stop'
    # Restore live Office keys first. Metadata refusal must never bypass this recovery.
    $registrationErrors = @()
    foreach ($entry in $Snapshot.AddIns) {
        try { Write-BuildRegistryKey $entry }
        catch { $registrationErrors += $_.Exception.Message }
    }
    if ($registrationErrors.Count -gt 0) { throw ($registrationErrors -join '; ') }
    $current = Get-VstoBuildState $Snapshot.RepositoryRoot
    if (($Snapshot | ConvertTo-Json -Depth 12 -Compress) -ceq ($current | ConvertTo-Json -Depth 12 -Compress)) {
        Write-Output 'VSTO_BUILD_STATE_RESTORED (unchanged)'
        return
    }

    if (Test-Path -LiteralPath $Snapshot.MetadataPath) {
        $root = Get-Item -LiteralPath $Snapshot.MetadataPath
        foreach ($name in @(($root.GetValueNames() | Sort-Object))) {
            if (-not ($name.StartsWith($Snapshot.Prefix, [StringComparison]::OrdinalIgnoreCase) -or $name.StartsWith($Snapshot.EncodedPrefix, [StringComparison]::OrdinalIgnoreCase))) { continue }
            $id = [string]$root.GetValue($name)
            if ($id -notmatch '^\{[0-9a-fA-F-]{36}\}$') { throw 'Unexpected VSTO metadata key ID during restoration.' }
            # Never delete a record referenced by another checkout.
            $otherReferences = @($root.GetValueNames() | Where-Object { $_ -ne $name -and [string]$root.GetValue($_) -eq $id })
            if ($otherReferences.Count -gt 0) { throw 'VSTO metadata record is shared; refusing destructive cleanup.' }
            $child = Join-Path $Snapshot.MetadataPath $id
            if (-not $child.StartsWith($Snapshot.MetadataPath + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Invalid metadata cleanup path.' }
            Remove-ItemProperty -LiteralPath $Snapshot.MetadataPath -Name $name
            if (Test-Path -LiteralPath $child) { Remove-Item -LiteralPath $child -Recurse -Force }
        }
    }
    if (Test-Path -LiteralPath $Snapshot.InclusionPath) {
        foreach ($key in (Get-ChildItem -LiteralPath $Snapshot.InclusionPath | Sort-Object Name)) {
            $url = [string]$key.GetValue('Url')
            if (($url.StartsWith($Snapshot.Prefix, [StringComparison]::OrdinalIgnoreCase) -or $url.StartsWith($Snapshot.EncodedPrefix, [StringComparison]::OrdinalIgnoreCase))) {
                $keyPath = $key.PSPath
                if (-not $keyPath.StartsWith('Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\VSTO\Security\Inclusion\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Invalid inclusion cleanup path.' }
                Remove-Item -LiteralPath $keyPath -Recurse -Force
            }
        }
    }
    foreach ($entry in @($Snapshot.Entries) + @($Snapshot.Inclusions)) { Write-BuildRegistryKey $entry }
    if ($Snapshot.Links.Count -gt 0) {
        if (-not (Test-Path -LiteralPath $Snapshot.MetadataPath)) { New-Item -Path $Snapshot.MetadataPath -Force | Out-Null }
        $root = Get-Item -LiteralPath $Snapshot.MetadataPath
        foreach ($link in $Snapshot.Links) { New-ItemProperty -LiteralPath $Snapshot.MetadataPath -Name $link.Name -Value $link.Id -PropertyType String -Force | Out-Null }
    }
    $after = Get-VstoBuildState $Snapshot.RepositoryRoot
    # Stable serialization makes any changed preserved value a build failure.
    if (($Snapshot | ConvertTo-Json -Depth 12 -Compress) -cne ($after | ConvertTo-Json -Depth 12 -Compress)) {
        throw 'VSTO build registry state was not restored exactly.'
    }
    Write-Output 'VSTO_BUILD_STATE_RESTORED'
}
