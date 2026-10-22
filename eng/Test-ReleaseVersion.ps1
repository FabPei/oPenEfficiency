param([string]$ScriptPath=(Join-Path $PSScriptRoot 'Set-ReleaseVersion.ps1'))
$ErrorActionPreference='Stop'
$checks=0
$fixtureRoot=Join-Path ([IO.Path]::GetTempPath()) ('ope-release-version-'+[guid]::NewGuid().ToString('N'))
function Assert-Version([bool]$Condition,[string]$Name){if(-not $Condition){throw "RELEASE_VERSION_FAIL $Name"};$script:checks++}
function New-VersionFixture([string]$Name,[bool]$Generated){
 $dir=Join-Path $fixtureRoot $Name
 foreach($hostName in @('oPenEfficiency','oPenEfficiency.Excel')){
  [IO.Directory]::CreateDirectory((Join-Path $dir "$hostName/Properties"))|Out-Null
  [IO.File]::WriteAllText((Join-Path $dir "$hostName/$hostName.csproj"),"<Project>`r`n<ApplicationVersion>2026.07.27.1</ApplicationVersion>`r`n</Project>`r`n",[Text.UTF8Encoding]::new($Generated))
  $info=if($Generated){'// centrally generated'}else{'[assembly: AssemblyVersion("2026.07.27.1")]'+"`r`n"+'[assembly: AssemblyFileVersion("2026.07.27.1")]'}
  [IO.File]::WriteAllText((Join-Path $dir "$hostName/Properties/AssemblyInfo.cs"),$info)
 }
 if($Generated){Copy-Item -LiteralPath (Join-Path $PSScriptRoot '../Directory.Build.props') -Destination (Join-Path $dir 'Directory.Build.props')}
 return $dir
}
function Get-FixtureBytes([string]$Directory){return (@(Get-ChildItem -LiteralPath $Directory -File -Recurse|Sort-Object FullName|ForEach-Object{[Convert]::ToBase64String([IO.File]::ReadAllBytes($_.FullName))}) -join '|')}
function Assert-VersionReject([string]$Directory,[string]$Version,[switch]$Check){
 $before=Get-FixtureBytes $Directory;$rejected=$false
 try{& $ScriptPath -RepositoryRoot $Directory -Version $Version -Check:$Check|Out-Null}catch{$rejected=$true}
 Assert-Version $rejected 'invalid fixture rejected'
 Assert-Version ((Get-FixtureBytes $Directory) -eq $before) 'validation rejection leaves every input byte unchanged'
}
try{
 $modern=New-VersionFixture 'generated' $true
 & $ScriptPath -RepositoryRoot $modern -Version 26.10.22|Out-Null
 Assert-Version ((Get-Content -Raw -LiteralPath (Join-Path $modern 'oPenEfficiency/oPenEfficiency.csproj')).Contains('2026.10.22.1')) 'generated project version updated'
 Assert-Version ((Get-Content -Raw -LiteralPath (Join-Path $modern 'oPenEfficiency/Properties/AssemblyInfo.cs')) -eq '// centrally generated') 'no duplicate generated assembly attributes'
 $projectBytes=[IO.File]::ReadAllBytes((Join-Path $modern 'oPenEfficiency/oPenEfficiency.csproj'))
 Assert-Version ($projectBytes[0] -eq 239 -and $projectBytes[1] -eq 187 -and $projectBytes[2] -eq 191) 'UTF8 BOM preserved'
 Assert-Version (-not ([Text.Encoding]::UTF8.GetString($projectBytes) -match '(?<!\r)\n')) 'CRLF preserved'
 $snapshot=Get-FixtureBytes $modern
 & $ScriptPath -RepositoryRoot $modern -Version 26.10.22 -Check|Out-Null
 Assert-Version ((Get-FixtureBytes $modern) -eq $snapshot) 'generated check is read-only'
 $legacy=New-VersionFixture 'legacy' $false
 & $ScriptPath -RepositoryRoot $legacy -Version 26.10.22|Out-Null
 & $ScriptPath -RepositoryRoot $legacy -Version 26.10.22 -Check|Out-Null
 Assert-Version ((Get-Content -Raw -LiteralPath (Join-Path $legacy 'oPenEfficiency.Excel/Properties/AssemblyInfo.cs')).Contains('AssemblyFileVersion("2026.10.22.1")')) 'legacy assembly attributes still updated'
 Assert-VersionReject $legacy 26.10.21 -Check
 Assert-VersionReject $legacy 26.7.27
 $missing=New-VersionFixture 'missing-stamp' $true;Remove-Item -LiteralPath (Join-Path $missing 'Directory.Build.props')
 Assert-VersionReject $missing 26.10.22
 $partial=New-VersionFixture 'partial-attributes' $false
 [IO.File]::WriteAllText((Join-Path $partial 'oPenEfficiency.Excel/Properties/AssemblyInfo.cs'),'[assembly: AssemblyVersion("2026.07.27.1")]')
 Assert-VersionReject $partial 26.10.22
 $broken=New-VersionFixture 'broken-stamp' $true
 $props=Join-Path $broken 'Directory.Build.props';$txt=[IO.File]::ReadAllText($props);[IO.File]::WriteAllText($props,$txt.Replace('AssemblyFileVersion','MissingFileVersion'))
 Assert-VersionReject $broken 26.10.22
 $duplicate=New-VersionFixture 'duplicate-producers' $true
 Copy-Item -LiteralPath (Join-Path $legacy 'oPenEfficiency.Excel/Properties/AssemblyInfo.cs') -Destination (Join-Path $duplicate 'oPenEfficiency.Excel/Properties/AssemblyInfo.cs')
 Assert-VersionReject $duplicate 26.10.22
 Write-Output "RELEASE_VERSION_PASS checks=$checks realRepositoryMutated=false OfficeExecuted=false"
}
finally{
 $resolved=[IO.Path]::GetFullPath($fixtureRoot);$temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath())
 if(-not $resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolved) -notlike 'ope-release-version-*'){throw 'Fixture cleanup path escaped expected temporary directory'}
 if(Test-Path -LiteralPath $resolved){Remove-Item -LiteralPath $resolved -Recurse -Force}
}
