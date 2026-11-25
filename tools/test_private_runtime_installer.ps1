param(
    [Parameter(Mandatory)][string]$BuildDirectory,
    [Parameter(Mandatory)][string]$BaselineAddin
)
$ErrorActionPreference = 'Stop'
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Save work and close Excel first' }
$record = Get-Content (Join-Path $BuildDirectory 'build-record.json') -Raw | ConvertFrom-Json
if (-not $record.smokeTest -or $record.canDistribute) { throw 'Only isolated internal installers are allowed' }
if ($record.vbaVersion -notin @('6.02.009','6.02.010')) { throw 'Unexpected VBA candidate version' }
$setup = $record.installer
if ((Get-FileHash $setup).Hash -ne $record.sha256 -or (Get-Item $setup).VersionInfo.ProductName.Trim() -ne 'NADABAS Installer Smoke') { throw 'Unexpected isolated package' }
$key = 'Software\NADABAS\InstallerSmoke'
$uninstallKey = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\NadabasLoaderInstallerSmoke_is1'
foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32,[Microsoft.Win32.RegistryView]::Registry64)) {
    $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser',$view)
    try { if ($hive.OpenSubKey($key) -or $hive.OpenSubKey($uninstallKey)) { throw 'Existing isolated state requires inspection' } } finally { $hive.Dispose() }
}
$checks = [Collections.Generic.List[string]]::new()
function Check([string]$Name,[bool]$Passed) { if (-not $Passed) { throw "Failed: $Name" }; $checks.Add($Name) }
$parent = Join-Path $env:TEMP ('nb-install-test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $parent | Out-Null
function RunSetup([string]$Root,[string]$Arch,[string]$Language) {
    $arguments = @('/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART',"/TESTEXCELARCH=$Arch",('/DIR="'+$Root+'"'),('/LOG="'+$parent+'\setup-'+[guid]::NewGuid().ToString('N')+'.log"'))
    if ($Language) { $arguments += "/NADABASLANG=$Language" }
    (Start-Process $setup -ArgumentList $arguments -WindowStyle Hidden -Wait -PassThru).ExitCode
}
foreach ($arch in @('x86','x64')) {
    $view = if ($arch -eq 'x86') { [Microsoft.Win32.RegistryView]::Registry32 } else { [Microsoft.Win32.RegistryView]::Registry64 }
    $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser',$view)
    $root = Join-Path $parent "$arch\installation"
    $profile = Join-Path $parent "$arch\profile"
    New-Item -ItemType Directory (Join-Path $profile 'AddIns') -Force | Out-Null
    $old = Join-Path $profile 'AddIns\NADABAS.XLAM'
    Copy-Item $BaselineAddin $old
    $sentinel = Join-Path $profile 'user-workbook-and-database.txt'
    'User data must survive install, repair and uninstall.' | Set-Content $sentinel
    $originalHash = (Get-FileHash $sentinel).Hash
    try {
        Check "$arch-first-silent-install-requires-language" ((RunSetup $root $arch '') -ne 0)
        Check "$arch-fresh-install" ((RunSetup $root $arch '4') -eq 0)
        Check "$arch-only-selected-runtime" ((Test-Path "$root\runtime\$arch\dotnet.exe") -and @(Get-ChildItem "$root\runtime" -Directory).Count -eq 1)
        $unexpectedRuntime = Join-Path $root "runtime\$arch\host\fxr\10.0.0"
        New-Item -ItemType Directory $unexpectedRuntime | Out-Null
        Check "$arch-runtime-patch-change-refused" ((RunSetup $root $arch '4') -ne 0)
        # Remove only the empty synthetic directory created above, never a runtime tree.
        [IO.Directory]::Delete($unexpectedRuntime,$false)
        $options = $hive.OpenSubKey("$key\Excel\Options")
        try { Check "$arch-startup-registered" ($options.GetValue('OPEN') -eq ('/R "'+$root+'\Nadabas.Startup.xll"')) } finally { $options.Dispose() }
        $selected = & "$root\runtime\$arch\dotnet.exe" "$root\tools\Nadabas.VersionStore.dll" inspect $root | ConvertFrom-Json
        Check "$arch-approved-vba" ($LASTEXITCODE -eq 0 -and $selected.ReleaseVersion -eq $record.vbaVersion)
        if ($record.vbaVersion -eq '6.02.010') {
            & "$root\runtime\$arch\dotnet.exe" "$root\tools\Nadabas.VersionStore.dll" activate $root slot-fix602009 | Out-Null
            Check "$arch-previous-vba-available" ($LASTEXITCODE -eq 0)
            & "$root\runtime\$arch\dotnet.exe" "$root\tools\Nadabas.VersionStore.dll" activate $root slot-fix602010 | Out-Null
            Check "$arch-access-repair-reselected" ($LASTEXITCODE -eq 0)
        }
        $other = if ($arch -eq 'x86') {'x64'} else {'x86'}
        Check "$arch-architecture-change-refused" ((RunSetup $root $other '3') -ne 0)
        # Model an older managed loader registration in the exact owned install root.
        $options = $hive.OpenSubKey("$key\Excel\Options",$true)
        try { $options.SetValue('OPEN',('/R "'+$root+'\Nadabas.Loader.Test.xll"')) } finally { $options.Dispose() }
        Check "$arch-repair-and-loader-upgrade" ((RunSetup $root $arch '') -eq 0)
        $language = $hive.OpenSubKey("$key\Language")
        try { Check "$arch-repair-preserves-language" ($language.GetValue('Code') -eq '4') } finally { $language.Dispose() }
        $options = $hive.OpenSubKey("$key\Excel\Options")
        try { Check "$arch-no-duplicate-loader" ($options.GetValueNames().Count -eq 1 -and $options.GetValue('OPEN') -eq ('/R "'+$root+'\Nadabas.Startup.xll"')) } finally { $options.Dispose() }
        # Replace only the synthetic registration with the approved legacy fixture.
        $options = $hive.OpenSubKey("$key\Excel\Options",$true)
        try { $options.SetValue('OPEN',('/R "'+$old+'"')) } finally { $options.Dispose() }
        Check "$arch-legacy-handoff" ((RunSetup $root $arch '5') -eq 0)
        $language = $hive.OpenSubKey("$key\Language")
        try { Check "$arch-portuguese-selection" ($language.GetValue('Code') -eq '5') } finally { $language.Dispose() }
        Check "$arch-legacy-backup-verified" ((Get-FileHash "$root-legacy-backup\NADABAS.XLAM").Hash -eq (Get-FileHash $old).Hash)
        Check "$arch-migrated-repair" ((RunSetup $root $arch '6') -eq 0)
        # Changed legacy bytes must block rollback rather than overwrite the user's edit.
        [IO.File]::AppendAllText($old,'changed-test-fixture')
        Check "$arch-changed-legacy-refused" ((RunSetup $root $arch '') -ne 0)
        Copy-Item $BaselineAddin $old -Force
        $remove = Start-Process "$root\unins000.exe" -ArgumentList '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -WindowStyle Hidden -Wait -PassThru
        Check "$arch-uninstall" ($remove.ExitCode -eq 0)
        $options = $hive.OpenSubKey("$key\Excel\Options")
        try { Check "$arch-legacy-registration-restored" ($options.GetValue('OPEN') -eq ('/R "'+$old+'"')) } finally { $options.Dispose() }
        Check "$arch-user-data-preserved" ((Get-FileHash $sentinel).Hash -eq $originalHash)
        Check "$arch-backup-retained" (Test-Path "$root-legacy-backup\registration.txt")
        # This exact synthetic registry root was absent at preflight; never touch Office.
        $hive.DeleteSubKeyTree($key)
    } finally { $hive.Dispose() }
}
[ordered]@{passed=$true; scope='isolated installer with real x86/x64 runtimes, not Excel acceptance'; checks=$checks.ToArray(); installerSha256=$record.sha256} |
    ConvertTo-Json -Depth 4 | Set-Content (Join-Path $parent 'report.json')
Write-Output "Passed $($checks.Count) checks. Report: $parent\report.json"
