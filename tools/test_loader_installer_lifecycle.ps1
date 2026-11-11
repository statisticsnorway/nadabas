param([Parameter(Mandatory)][string]$SmokeInstaller, [string]$PreviousSmokeInstaller, [string]$PreviousExpectedSlot = "slot-test602008")
$ErrorActionPreference = 'Stop'
# Use ONLY the /DSmokeTest compiler variant: separate AppId, shortcuts and registry.
$repo = Split-Path $PSScriptRoot
$root = Join-Path $repo ('artifacts/loader-installer/lifecycle-' + [guid]::NewGuid().ToString('N'))
$registryRoot = 'HKCU:\Software\NADABAS\InstallerSmoke'
$options = "$registryRoot\Excel\Options"
$manager = "$registryRoot\Excel\Add-in Manager"
$uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\NadabasLoaderInstallerSmoke_is1'
if ((Test-Path $registryRoot) -or (Test-Path $uninstallKey)) { throw 'Existing smoke state: inspect before running' }
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close Excel before the installer test' }
if ((Get-AuthenticodeSignature $SmokeInstaller).Status -ne 'Valid') { throw 'Unsigned smoke installer' }
if ((Get-Item -LiteralPath $SmokeInstaller).VersionInfo.ProductName.Trim() -ne 'NADABAS Installer Smoke') { throw 'Use only the isolated /DSmokeTest variant, never the real installer' }
if ($PreviousSmokeInstaller) {
    if ((Get-AuthenticodeSignature $PreviousSmokeInstaller).Status -ne 'Valid' -or
        (Get-Item -LiteralPath $PreviousSmokeInstaller).VersionInfo.ProductName.Trim() -ne 'NADABAS Installer Smoke') { throw 'Previous package must also be a signed isolated smoke installer' }
}
$checks = [Collections.Generic.List[string]]::new()
function Check([string]$name, [bool]$ok) { if (-not $ok) { throw "Failed: $name" }; $checks.Add($name) }
function Install([bool]$expectSuccess = $true, [string]$installer = $SmokeInstaller, [string]$language = "") {
    $args = @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', ('/DIR="' + $root + '"'))
    if ($language) { $args += "/NADABASLANG=$language" }
    $p = Start-Process -FilePath $installer -ArgumentList $args -WindowStyle Hidden -Wait -PassThru
    if ($expectSuccess -and $p.ExitCode -ne 0) { throw "Install failed: $($p.ExitCode)" }
    if (-not $expectSuccess -and $p.ExitCode -eq 0) { throw 'Unexpected installation success' }
}
function Uninstall {
    $file = Join-Path $root 'unins000.exe'
    if (Test-Path -LiteralPath $file) {
        $p = Start-Process -FilePath $file -ArgumentList '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -WindowStyle Hidden -Wait -PassThru
        if ($p.ExitCode -ne 0) { throw 'Uninstall failed' }
    }
}
try {
    New-Item -ItemType Directory -Path $root | Out-Null
    # Simulate an existing language preference in the isolated profile.
    New-Item -Path ($registryRoot + "\Language") -Force | Out-Null
    Set-ItemProperty -Path ($registryRoot + "\Language") -Name Code -Value "4"
    New-Item -Path $options -Force | Out-Null
    New-Item -Path $manager -Force | Out-Null
    Set-ItemProperty $options -Name OPEN -Value 'unrelated-addin.xll'
    $unknown = Join-Path $root 'unrelated.txt'
    Set-Content -LiteralPath $unknown -Value 'keep me'
    Install $false
    Check 'unknown-directory-preserved' ((Get-Content $unknown) -eq 'keep me' -and -not (Test-Path (Join-Path $root 'Nadabas.Loader.Test.xll')))
    Remove-Item -LiteralPath $unknown
    Remove-ItemProperty -Path ($registryRoot + "\Language") -Name Code
    Install $false
    Check 'fresh-silent-install-requires-language' (-not (Test-Path $uninstallKey))
    Install $false -language '99'
    Check 'invalid-language-rejected-before-install' (-not (Test-Path $uninstallKey))
    Set-ItemProperty -Path ($registryRoot + "\Language") -Name Code -Value '4'
    $previousLoaderHash = $null
    if ($PreviousSmokeInstaller) {
        Install -installer $PreviousSmokeInstaller
        $oldTool = Join-Path $root 'tools/Nadabas.VersionStore.exe'
        $oldSelection = (& $oldTool inspect $root | ConvertFrom-Json)
        Check 'previous-real-package-installed' ($oldSelection.Slot -eq $PreviousExpectedSlot)
        $previousLoaderHash = (Get-FileHash -LiteralPath (Join-Path $root 'Nadabas.Loader.Test.xll')).Hash
    }
    Install
    if ($previousLoaderHash) {
        Check 'new-package-replaces-loader-and-retains-old-payload' (
            (Get-FileHash -LiteralPath (Join-Path $root 'Nadabas.Loader.Test.xll')).Hash -ne $previousLoaderHash -and
            (Get-FileHash -LiteralPath (Join-Path $root 'versions/slot-test602008/NADABAS.TEST.slot-test602008.xlam')).Hash -eq 'C2F70BBA9F93CBE78FA695C6E48057C9092B1FC976C7F788718590B26D92D126')
    }
    Check 'fresh-install-registered' ((Get-ItemPropertyValue $options OPEN1) -eq ('/R "' + (Join-Path $root 'Nadabas.Loader.Test.xll') + '"'))
    Check 'other-startup-entry-preserved' ((Get-ItemPropertyValue $options OPEN) -eq 'unrelated-addin.xll')
    $tool = Join-Path $root 'tools/Nadabas.VersionStore.exe'
    $installed = & $tool inspect $root | ConvertFrom-Json
    Check 'silent-install-preserves-french' ((Get-ItemPropertyValue ($registryRoot + "\Language") Code) -eq '4')
    foreach ($code in @('3','4','5','6')) {
        Install -language $code
        Check "explicit-language-$code-saved" ((Get-ItemPropertyValue ($registryRoot + "\Language") Code) -eq $code)
        Install
        Check "silent-upgrade-preserves-language-$code" ((Get-ItemPropertyValue ($registryRoot + "\Language") Code) -eq $code)
    }
    Check 'multilingual-demo-selected' ($installed.Slot -eq 'slot-demo602006' -and $installed.ReleaseVersion -eq '6.02.006')
    & $tool activate $root slot-a
    if ($LASTEXITCODE -ne 0) { throw 'Baseline selection failed' }
    & $tool queue $root slot-demo602006
    if ($LASTEXITCODE -ne 0) { throw 'Queue failed' }
    $queued = & $tool inspect $root | ConvertFrom-Json
    Check 'installed-queue-keeps-current-release' ($queued.ReleaseVersion -eq '6.02.005' -and (Test-Path (Join-Path $root 'pending-update.json')))
    & $tool cancel-update $root
    if ($LASTEXITCODE -ne 0) { throw 'Cancel failed' }
    Check 'installed-cancellation-removes-pending-request' (-not (Test-Path (Join-Path $root 'pending-update.json')))
    & $tool activate $root slot-b
    if ($LASTEXITCODE -ne 0) { throw 'Pointer setup failed' }
    $active = Join-Path $root 'active.json'
    $sessions = Join-Path $root 'sessions'
    New-Item -ItemType Directory -Path $sessions | Out-Null
    $report = Join-Path $sessions 'test.json'
    Set-Content $report '{"test":"preserve"}'
    $reportHash = (Get-FileHash $report).Hash
    $slot = Join-Path $root 'versions/slot-a/NADABAS.TEST.slot-a.xlam'
    Set-ItemProperty $options -Name OPEN2 -Value ('/R "' + $slot + '"')
    Set-ItemProperty $manager -Name $slot -Value ''
    Install
    $installed = & $tool inspect $root | ConvertFrom-Json
    Check 'reinstall-selects-packaged-release' ($installed.Slot -eq 'slot-demo602006' -and $installed.ReleaseVersion -eq '6.02.006')
    Check 'session-report-preserved' ((Get-FileHash $report).Hash -eq $reportHash)
    Check 'old-payload-registration-removed' (-not (Get-Item $options).GetValueNames().Contains('OPEN2') -and -not (Get-Item $manager).GetValueNames().Contains($slot))
    Check 'no-duplicate-loader-registration' (@((Get-Item $options).GetValueNames()).Count -eq 2)
    $originalHash = (Get-FileHash $slot).Hash
    $before = (Get-FileHash $active).Hash
    # Corrupt only this disposable copy to verify fail-closed reinstall behavior.
    [IO.File]::SetAttributes($slot, [IO.FileAttributes]::Normal)
    [IO.File]::AppendAllText($slot, 'tamper')
    Install $false
    Check 'modified-slot-not-overwritten' ((Get-FileHash $slot).Hash -ne $originalHash)
    Check 'rejected-upgrade-keeps-selection' ((Get-FileHash $active).Hash -eq $before)
    Uninstall
    Check 'uninstall-keeps-unrelated-registration' ((Get-ItemPropertyValue $options OPEN) -eq 'unrelated-addin.xll')
    Check 'uninstall-removes-own-registration' (@((Get-Item $options).GetValueNames()).Count -eq 1)
    Check 'uninstall-preserves-report' ((Get-FileHash $report).Hash -eq $reportHash)
    Install
    Check 'reinstall-over-residual-reports' ((Get-FileHash $report).Hash -eq $reportHash -and (Test-Path (Join-Path $root 'Nadabas.Loader.Test.xll')))
    Uninstall
    Check 'uninstall-preserves-language' ((Get-ItemPropertyValue ($registryRoot + "\Language") Code) -eq '6')
    Check 'smoke-registration-removed' (-not (Test-Path $uninstallKey))
    [ordered]@{ passed = $true; scope = 'isolated-installer-registry-not-live-excel'; checks = $checks.ToArray(); root = $root } |
        ConvertTo-Json -Depth 4 | Set-Content (Join-Path $root 'lifecycle-report.json')
    Write-Output "Passed $($checks.Count) checks. Report: $root\lifecycle-report.json"
} finally {
    Uninstall
    # This exact isolated key was absent before this test. Never touch Office keys.
    if (Test-Path $registryRoot) { Remove-Item -LiteralPath $registryRoot -Recurse }
}
