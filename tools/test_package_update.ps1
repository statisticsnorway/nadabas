param(
    [Parameter(Mandatory)][string]$Updater,
    [Parameter(Mandatory)][string]$PreviousSmokeInstaller,
    [Parameter(Mandatory)][string]$NextSmokeInstaller,
    [string]$ExpectedLoaderSha256
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot
$registryRoot = 'HKCU:\Software\NADABAS\InstallerSmoke'
$uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\NadabasLoaderInstallerSmoke_is1'
if ((Test-Path $registryRoot) -or (Test-Path $uninstallKey)) { throw 'Existing isolated smoke state; inspect before testing.' }
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close Excel before isolated package update testing.' }
foreach ($file in @($Updater, $PreviousSmokeInstaller, $NextSmokeInstaller)) {
    $signature = Get-AuthenticodeSignature -LiteralPath $file
    if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Thumbprint -ne 'A70F7882B2F9F5C6D23A22C8AEC70211854B9EAF') { throw 'Unapproved test executable.' }
}
foreach ($file in @($PreviousSmokeInstaller, $NextSmokeInstaller)) {
    if ((Get-Item -LiteralPath $file).VersionInfo.ProductName.Trim() -ne 'NADABAS Installer Smoke') { throw 'Use only isolated smoke installers.' }
}
$directory = Join-Path $repo ('artifacts/package-update/test-' + [guid]::NewGuid().ToString('N'))
$root = Join-Path $directory 'installation'
$cache = Join-Path $directory 'cache'
$badCache = Join-Path $directory 'tampered-cache'
New-Item -ItemType Directory -Path $directory | Out-Null
$checks = [Collections.Generic.List[string]]::new()
function Check([string]$name, [bool]$passed) { if (-not $passed) { throw "Failed: $name" }; $checks.Add($name) }
try {
    # Simulate an existing language preference in the isolated profile.
    New-Item -Path ($registryRoot + "\Language") -Force | Out-Null
    Set-ItemProperty -Path ($registryRoot + "\Language") -Name Code -Value "4"
    $process = Start-Process -FilePath $PreviousSmokeInstaller -ArgumentList @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', ('/DIR="' + $root + '"')) -WindowStyle Hidden -Wait -PassThru
    Check 'previous-package-installed' ($process.ExitCode -eq 0)
    $stageHelper = Join-Path $root 'updater/Nadabas.PackageUpdate.exe'
    if (-not (Test-Path -LiteralPath $stageHelper)) { $stageHelper = $Updater }
    $before = (Get-FileHash -LiteralPath (Join-Path $root 'Nadabas.Loader.Test.xll')).Hash

    & $stageHelper stage $root $NextSmokeInstaller (Join-Path $root 'bad-cache') --smoke
    Check 'cache-inside-installation-rejected' ($LASTEXITCODE -ne 0 -and -not (Test-Path (Join-Path $root 'bad-cache')))
    & $stageHelper stage $root $PreviousSmokeInstaller (Join-Path $directory 'same-version') --smoke
    Check 'same-version-rejected' ($LASTEXITCODE -ne 0 -and -not (Test-Path (Join-Path $directory 'same-version')))

    & $stageHelper stage $root $NextSmokeInstaller $badCache --smoke
    Check 'tamper-test-staged' ($LASTEXITCODE -eq 0)
    $badInstaller = Join-Path $badCache 'setup.exe'
    $stream = [IO.File]::Open($badInstaller, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    try { $stream.Position = 4096; $value = $stream.ReadByte(); $stream.Position = 4096; $stream.WriteByte($value -bxor 1) } finally { $stream.Dispose() }
    & (Join-Path $badCache 'Nadabas.PackageUpdate.exe') apply $badCache
    Check 'modified-cached-package-not-executed' ($LASTEXITCODE -ne 0 -and -not (Test-Path (Join-Path $badCache 'apply-started.json')) -and (Get-FileHash -LiteralPath (Join-Path $root 'Nadabas.Loader.Test.xll')).Hash -eq $before)

    $failureCache = Join-Path $directory 'installer-rejection-cache'
    & $stageHelper stage $root $NextSmokeInstaller $failureCache --smoke
    Check 'installer-failure-test-staged' ($LASTEXITCODE -eq 0)
    $ownedSlot = Join-Path $root 'versions/slot-a/NADABAS.TEST.slot-a.xlam'
    $originalBytes = [IO.File]::ReadAllBytes($ownedSlot)
    try {
        [IO.File]::SetAttributes($ownedSlot, [IO.FileAttributes]::Normal)
        [IO.File]::WriteAllText($ownedSlot, 'Deliberate corruption in an isolated test slot')
        & (Join-Path $failureCache 'Nadabas.PackageUpdate.exe') apply $failureCache
        $failureStatus = Get-Content -LiteralPath (Join-Path $failureCache 'status.json') -Raw | ConvertFrom-Json
        Check 'installer-rejection-records-repair-state' ($LASTEXITCODE -ne 0 -and $failureStatus.status -eq 'repair-required')
        Check 'rejected-install-keeps-old-loader' ((Get-FileHash -LiteralPath (Join-Path $root 'Nadabas.Loader.Test.xll')).Hash -eq $before)
    } finally {
        [IO.File]::WriteAllBytes($ownedSlot, $originalBytes)
        [IO.File]::SetAttributes($ownedSlot, [IO.FileAttributes]::ReadOnly)
    }

    & $stageHelper stage $root $NextSmokeInstaller $cache --smoke
    Check 'signed-next-package-staged' ($LASTEXITCODE -eq 0)
    Check 'staging-keeps-current-loader' ((Get-FileHash -LiteralPath (Join-Path $root 'Nadabas.Loader.Test.xll')).Hash -eq $before)
    & (Join-Path $cache 'Nadabas.PackageUpdate.exe') apply $cache
    Check 'external-helper-installs-whole-package' ($LASTEXITCODE -eq 0)
    $status = Get-Content -LiteralPath (Join-Path $cache 'status.json') -Raw | ConvertFrom-Json
    Check 'post-install-version-verified' ($status.status -eq 'installed' -and ([version]$status.version).ToString(3) -eq ([version]((Get-Item -LiteralPath $NextSmokeInstaller).VersionInfo.ProductVersion.Trim())).ToString(3))
    $after = (Get-FileHash -LiteralPath (Join-Path $root 'Nadabas.Loader.Test.xll')).Hash
    if ($ExpectedLoaderSha256) { Check 'loader-matches-incoming-release' ($after -eq $ExpectedLoaderSha256) }
    else { Check 'loader-was-replaced' ($after -ne $before) }
    Check 'updater-installed-without-locking-target' ((Get-FileHash -LiteralPath (Join-Path $root 'updater/Nadabas.PackageUpdate.exe')).Hash -eq (Get-FileHash -LiteralPath $Updater).Hash)
    $selection = (& (Join-Path $root 'tools/Nadabas.VersionStore.exe') inspect $root | ConvertFrom-Json)
    Check 'new-multilingual-demo-selected' ($selection.Slot -eq 'slot-demo602006')
    Check 'old-beta-retained-for-rollback' ((Get-FileHash -LiteralPath (Join-Path $root 'versions/slot-beta602006/NADABAS.TEST.slot-beta602006.xlam')).Hash -eq '82F22D19D66A6D682E5A5FBE7C15D7013929D36BF37820FAF915617687D3872D')
    & (Join-Path $cache 'Nadabas.PackageUpdate.exe') apply $cache
    Check 'completed-job-replay-rejected' ($LASTEXITCODE -ne 0)
    [ordered]@{ passed = $true; scope = 'isolated-signed-package-handoff-not-live-excel'; checks = $checks; root = $root; cache = $cache } |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'package-update-report.json') -Encoding utf8
    Write-Output "Passed $($checks.Count) checks. Report: $directory\package-update-report.json"
} finally {
    $uninstaller = Join-Path $root 'unins000.exe'
    if (Test-Path -LiteralPath $uninstaller) {
        $process = Start-Process -FilePath $uninstaller -ArgumentList '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -WindowStyle Hidden -Wait -PassThru
        if ($process.ExitCode -ne 0) { throw 'Isolated uninstall failed; preserve and inspect smoke state.' }
    }
    # Exact test-only key, preflight proved absent before this run.
    if (Test-Path -LiteralPath $registryRoot) { Remove-Item -LiteralPath $registryRoot -Recurse }
}
