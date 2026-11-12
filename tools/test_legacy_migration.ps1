param(
    [Parameter(Mandatory)][string]$SmokeInstaller,
    [string]$ReleaseAddin = 'C:\NADABAS\NADABAS.XLAM',
    [string]$ExpectedSlot = 'slot-demo602006',
    [string]$ExpectedVbaVersion = '6.02.006'
)
$ErrorActionPreference = 'Stop'
$baseline = '2F7432B9B04C414337EB1AFEB07E2B93F05E4A5A6D151073FA967A2F7EA89A85'
$publisher = 'A70F7882B2F9F5C6D23A22C8AEC70211854B9EAF'
$registry = 'HKCU:\Software\NADABAS\InstallerSmoke'
$uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\NadabasLoaderInstallerSmoke_is1'
$options = "$registry\Excel\Options"
$manager = "$registry\Excel\Add-in Manager"
function Assert-ClosedExcel {
    if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close Excel before the isolated migration rehearsal.' }
}
Assert-ClosedExcel
if (-not [Environment]::Is64BitProcess) { throw 'Run in 64-bit PowerShell.' }
if ((Test-Path $registry) -or (Test-Path $uninstallKey)) { throw 'Existing smoke state. Do not reuse it.' }
$sig = Get-AuthenticodeSignature -LiteralPath $SmokeInstaller
$info = (Get-Item -LiteralPath $SmokeInstaller).VersionInfo
if ($sig.Status -ne 'Valid' -or $sig.SignerCertificate.Thumbprint -ne $publisher -or
    $info.ProductName.Trim() -ne 'NADABAS Installer Smoke') { throw 'Use only the signed isolated smoke installer.' }
if ((Get-FileHash -LiteralPath $ReleaseAddin).Hash -ne $baseline) { throw 'Use the approved legacy baseline.' }
$parent = Join-Path (Split-Path $PSScriptRoot) ('artifacts/legacy-migration/run-' + [guid]::NewGuid().ToString('N'))
$old = Join-Path $parent 'profile/AddIns/NADABAS.XLAM'
$target = Join-Path $parent 'installation'
$loader = Join-Path $target 'Nadabas.Loader.Test.xll'
$backup = Join-Path ($target + '-legacy-backup') 'NADABAS.XLAM'
$manifest = Join-Path ($target + '-legacy-backup') 'registration.txt'
$oldValue = '/R "' + $old + '"'
$loaderValue = '/R "' + $loader + '"'
$checks = [Collections.Generic.List[string]]::new()
function Check([string]$label, [bool]$ok) { if (-not $ok) { throw "Failed: $label" }; $checks.Add($label) }
function Entries([string]$value) {
    $key = Get-Item -LiteralPath $options
    $key.GetValueNames() | Where-Object { $_ -match '^OPEN[0-9]*$' -and $key.GetValue($_) -ceq $value }
}
function Run-Setup([bool]$expected) {
    $p = Start-Process -FilePath $SmokeInstaller -ArgumentList @('/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART',('/DIR="' + $target + '"')) -WindowStyle Hidden -Wait -PassThru
    if (($p.ExitCode -eq 0) -ne $expected) { throw "Unexpected Setup exit code $($p.ExitCode)" }
}
function Uninstall-Test([bool]$expected = $true) {
    $uninstaller = Join-Path $target 'unins000.exe'
    if (Test-Path -LiteralPath $uninstaller) {
        $p = Start-Process -FilePath $uninstaller -ArgumentList '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -WindowStyle Hidden -Wait -PassThru
        if (($p.ExitCode -eq 0) -ne $expected) { throw "Unexpected uninstall exit code $($p.ExitCode)" }
    }
}
New-Item -ItemType Directory -Path (Split-Path $old) | Out-Null
New-Item -Path $registry -ErrorAction Stop | Out-Null
try {
    # Simulate an existing language preference in the isolated profile.
    New-Item -Path ($registry + "\Language") -Force | Out-Null
    Set-ItemProperty -Path ($registry + "\Language") -Name Code -Value "4"
    New-Item -Path $options,$manager -Force | Out-Null
    New-ItemProperty -LiteralPath $options -Name OPEN -PropertyType String -Value 'unrelated.xll' | Out-Null
    New-ItemProperty -LiteralPath $options -Name OPEN1 -PropertyType String -Value $oldValue | Out-Null
    New-ItemProperty -LiteralPath $manager -Name $old -PropertyType String -Value '' | Out-Null
    [IO.File]::WriteAllText($old, 'Local custom version')
    Run-Setup $false
    Check 'unknown-copy-preserved-before-install' ([IO.File]::ReadAllText($old) -eq 'Local custom version' -and @(Entries $oldValue).Count -eq 1 -and -not (Test-Path $loader))
    [IO.File]::Copy($ReleaseAddin, $old, $true)
    New-ItemProperty -LiteralPath $options -Name OPEN2 -PropertyType String -Value $oldValue | Out-Null
    Run-Setup $false
    Check 'duplicate-registration-refused' (@(Entries $oldValue).Count -eq 2 -and -not (Test-Path $backup))
    Remove-ItemProperty -LiteralPath $options -Name OPEN2
    $startup = Join-Path $parent 'profile/XLSTART'
    New-Item -ItemType Directory -Path $startup | Out-Null
    $autoFile = Join-Path $startup 'NADABAS.XLAM'
    [IO.File]::Copy($ReleaseAddin, $autoFile)
    Run-Setup $false
    Check 'startup-copy-requires-review' ((Get-FileHash $autoFile).Hash -eq $baseline -and @(Entries $oldValue).Count -eq 1)
    [IO.File]::Move($autoFile, (Join-Path $parent 'preserved-startup-copy.xlam'))
    Run-Setup $true
    Check 'setup-backed-up-original-bytes' ((Get-FileHash $backup).Hash -eq $baseline)
    Check 'setup-recorded-exact-slot-and-value' (((Get-Content $manifest) -join '|') -ceq ('NADABAS legacy registration v1 REG_SZ|OPEN1|' + $oldValue))
    Check 'sole-loader-reuses-original-slot' (@(Entries $oldValue).Count -eq 0 -and @(Entries $loaderValue).Count -eq 1 -and (Get-ItemPropertyValue $options OPEN1) -ceq $loaderValue)
    Check 'legacy-file-still-in-place' ((Get-FileHash $old).Hash -eq $baseline)
    $selected = & (Join-Path $target 'tools/Nadabas.VersionStore.exe') inspect $target | ConvertFrom-Json
    Check 'managed-payload-verified' ($LASTEXITCODE -eq 0 -and $selected.Slot -eq $ExpectedSlot -and $selected.ReleaseVersion -eq $ExpectedVbaVersion)
    Run-Setup $true
    Check 'repair-preserves-single-entry-and-backup' (@(Entries $loaderValue).Count -eq 1 -and (Get-FileHash $backup).Hash -eq $baseline)
    [IO.File]::WriteAllText($backup, 'Damaged disposable backup')
    Run-Setup $false
    Uninstall-Test $false
    Check 'damaged-backup-blocks-repair-and-uninstall' ((Test-Path $loader) -and @(Entries $loaderValue).Count -eq 1)
    [IO.File]::Copy($old, $backup, $true)
    Set-ItemProperty -LiteralPath $options -Name OPEN1 -Value 'later-user-addin.xll'
    Uninstall-Test $false
    Check 'later-registration-preserved' ((Get-ItemPropertyValue $options OPEN1) -ceq 'later-user-addin.xll' -and (Test-Path $loader))
    Set-ItemProperty -LiteralPath $options -Name OPEN1 -Value $loaderValue
    [IO.File]::WriteAllText($old, 'Changed after installation')
    Uninstall-Test $false
    Check 'later-file-preserved' ([IO.File]::ReadAllText($old) -eq 'Changed after installation' -and (Test-Path $loader))
    [IO.File]::Move($old, (Join-Path $parent 'changed-copy.xlam'))
    Uninstall-Test
    Check 'uninstall-restored-original-file-and-entry' ((Get-FileHash $old).Hash -eq $baseline -and (Get-ItemPropertyValue $options OPEN1) -ceq $oldValue -and -not (Test-Path $loader))
    Check 'other-addin-and-manager-preserved' ((Get-ItemPropertyValue $options OPEN) -ceq 'unrelated.xll' -and (Get-Item $manager).GetValueNames() -contains $old)
    $report = Join-Path $parent 'migration-report.json'
    [ordered]@{ passed=$true; scope='Setup and Uninstall, isolated files and registry, no Excel session'; checks=$checks.ToArray(); backup=$backup; registrationBackup=$manifest } |
        ConvertTo-Json | Set-Content -LiteralPath $report -Encoding utf8
    Write-Output "Passed $($checks.Count) Setup migration checks. Report: $report"
} finally {
    # Preserve a conflicted installation for diagnosis; never force its removal.
    if (-not (Test-Path (Join-Path $target 'unins000.exe'))) {
        Remove-Item -LiteralPath $registry -Recurse
    }
}
