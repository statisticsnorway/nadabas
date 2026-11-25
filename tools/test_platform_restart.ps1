param([Parameter(Mandatory)][string]$PreviousSmokeInstaller, [Parameter(Mandatory)][string]$NextSmokeInstaller)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot
$registry = 'HKCU:\Software\NADABAS\InstallerSmoke'
$uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\NadabasLoaderInstallerSmoke_is1'
if ((Test-Path $registry) -or (Test-Path $uninstallKey)) { throw 'Existing smoke installation; inspect first.' }
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Save and close Excel first.' }
foreach ($file in @($PreviousSmokeInstaller, $NextSmokeInstaller)) {
    $signature = Get-AuthenticodeSignature -LiteralPath $file
    if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Thumbprint -ne 'A70F7882B2F9F5C6D23A22C8AEC70211854B9EAF' -or
        (Get-Item $file).VersionInfo.ProductName.Trim() -ne 'NADABAS Installer Smoke') { throw 'Use only signed isolated smoke packages.' }
}
$run = Join-Path $repo ('artifacts/platform-restart/run-' + [guid]::NewGuid().ToString('N'))
$root = Join-Path $run 'installation'
New-Item -ItemType Directory -Path $run | Out-Null
$checks = [Collections.Generic.List[string]]::new()
function Check([string]$name, [bool]$passed) { if (-not $passed) { throw "Failed: $name" }; $checks.Add($name) }
$owner = $null
try {
    $setup = Start-Process -FilePath $PreviousSmokeInstaller -ArgumentList @('/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART','/NADABASLANG=4',('/DIR="'+$root+'"')) -WindowStyle Hidden -PassThru -Wait
    Check 'previous-package-installed' ($setup.ExitCode -eq 0)
    $helper = Join-Path $root 'updater/Nadabas.PackageUpdate.exe'
    $cancelled = & $helper stage-for-excel $root $NextSmokeInstaller --smoke | ConvertFrom-Json
    Check 'first-job-staged' ($LASTEXITCODE -eq 0)
    & $helper cancel-for-excel $root $cancelled.Id
    Check 'cancel-preserves-installation-and-removes-pointer' ($LASTEXITCODE -eq 0 -and -not (Test-Path (Join-Path $root 'pending-platform-update')))
    $job = & $helper stage-for-excel $root $NextSmokeInstaller --smoke | ConvertFrom-Json
    Check 'new-job-has-distinct-identity' ($LASTEXITCODE -eq 0 -and $job.Id -ne $cancelled.Id)
    & $helper cancel-for-excel $root $cancelled.Id
    Check 'stale-cancel-cannot-remove-new-job' ($LASTEXITCODE -ne 0 -and (Get-Content (Join-Path $root 'pending-platform-update') -Raw) -eq $job.Id)
    $cache = Join-Path ($root + '-package-cache') $job.Id
    $cachedHelper = Join-Path $cache 'Nadabas.PackageUpdate.exe'
    Check 'external-helper-matches-installed-helper' ((Get-FileHash $cachedHelper).Hash -eq (Get-FileHash $helper).Hash)
    & $cachedHelper wait-apply $root $job.Id 0 0
    Check 'invalid-parent-does-not-install' ($LASTEXITCODE -ne 0 -and -not (Test-Path (Join-Path $cache 'apply-started.json')))

    # Only this COM-created, otherwise empty Excel process is closed by the test.
    if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Excel was opened during preparation; leave it unchanged.' }
    $owner = Start-Process -FilePath (Join-Path $repo '.venv/Scripts/python.exe') -ArgumentList @(('"'+(Join-Path $PSScriptRoot 'restart_test_excel_owner.py')+'"'),('"'+$run+'"')) -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $run 'excel-owner.stderr.log')
    $ownerDeadline = [DateTime]::UtcNow.AddSeconds(30)
    do {
        if (Test-Path (Join-Path $run 'excel-owner.json')) { break }
        if ($owner.HasExited) { throw 'Excel owner failed to start.' }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $ownerDeadline)
    $ownerInfo = Get-Content (Join-Path $run 'excel-owner.json') -Raw | ConvertFrom-Json
    $parent = Get-Process -Id $ownerInfo.pid
    if ($parent.ProcessName -ne 'EXCEL') { throw 'Unexpected test process identity.' }
    $ticks = $parent.StartTime.ToUniversalTime().Ticks
    $stdout = Join-Path $run 'waiter.stdout.log'
    $stderr = Join-Path $run 'waiter.stderr.log'
    $waiter = Start-Process -FilePath $cachedHelper -ArgumentList @('wait-apply',('"'+$root+'"'),$job.Id,$parent.Id,$ticks) -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr -PassThru
    $deadline = [DateTime]::UtcNow.AddSeconds(15)
    do {
        if (Test-Path $stdout) { $ack = Get-Content $stdout -Raw }
        if ($ack -match 'restart-helper-ready') { break }
        if ($waiter.HasExited) { throw 'Waiter exited before acknowledgement.' }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $deadline)
    Check 'helper-ready-before-excel-exit' ($ack -match 'restart-helper-ready')
    Check 'no-install-while-excel-running' (-not (Test-Path (Join-Path $cache 'apply-started.json')))
    Set-Content (Join-Path $run 'close-owned-excel') 'close'
    if (-not $owner.WaitForExit(30000) -or $owner.ExitCode -ne 0) { throw 'Owned Excel shutdown failed. Preserve the test state.' }
    if (-not $waiter.WaitForExit(60000)) { throw 'Waiter still running. Preserve test installation for inspection.' }
    Check 'waiter-completed-successfully' ($waiter.ExitCode -eq 0)
    $status = Get-Content (Join-Path $cache 'status.json') -Raw | ConvertFrom-Json
    $expectedVersion = [version]((Get-Item -LiteralPath $NextSmokeInstaller).VersionInfo.ProductVersion.Trim())
    Check 'new-package-installed-after-exit' ($status.status -eq 'installed' -and ([version]$status.version).ToString(3) -eq $expectedVersion.ToString(3))
    Check 'ready-pointer-cleared-after-install' (-not (Test-Path (Join-Path $root 'pending-platform-update')))
    $reopened = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" | Where-Object ParentProcessId -eq $waiter.Id)
    Check 'helper-reopened-excel' ($reopened.Count -eq 1)
    # Keep the reopened window for visual acceptance. The user closes it normally.
    # Avoid relying on MainWindowHandle from a background test session.
    Check 'language-preserved-through-restart-upgrade' ((Get-ItemPropertyValue ($registry+'\Language') Code) -eq '4')
    [ordered]@{passed=$true; scope='signed external updater with a real Excel exit and relaunch; not an Excel button click test'; checks=$checks.ToArray(); root=$root; cleanupPending=$true} |
        ConvertTo-Json | Set-Content (Join-Path $run 'restart-report.json')
    Write-Output "Passed $($checks.Count) restart checks. Report: $run\restart-report.json. Close the reopened test Excel normally, then uninstall this isolated test root."
} finally {
    if ($null -ne $owner -and -not $owner.HasExited) {
        Set-Content (Join-Path $run 'close-owned-excel') 'close'
        [void]$owner.WaitForExit(30000)
    }
    if (-not @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) {
        $uninstaller = Join-Path $root 'unins000.exe'
        if (Test-Path $uninstaller) {
            $remove = Start-Process -FilePath $uninstaller -ArgumentList '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -WindowStyle Hidden -Wait -PassThru
            if ($remove.ExitCode -ne 0) { throw 'Test uninstall failed; preserve smoke state.' }
        }
        if (-not (Test-Path $uninstallKey) -and (Test-Path $registry)) { Remove-Item -LiteralPath $registry -Recurse }
    }
}
