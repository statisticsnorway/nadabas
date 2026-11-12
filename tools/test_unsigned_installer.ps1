param(
    [Parameter(Mandatory)][string]$BuildDirectory,
    [Parameter(Mandatory)][string]$Compiler
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot
$record = Get-Content -LiteralPath (Join-Path $BuildDirectory 'package/build-record.json') -Raw | ConvertFrom-Json
if (-not $record.unsignedPreparation -or -not $record.productionCandidate) { throw 'Use an unsigned production-candidate preparation' }
$registry = 'HKCU:\Software\NADABAS\InstallerSmoke'
$uninstall = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\NadabasLoaderInstallerSmoke_is1'
if ((Test-Path $registry) -or (Test-Path $uninstall)) { throw 'Existing isolated test state; inspect it first' }
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close Excel before isolated installation testing' }
$normal = $record.installer
if ((Get-AuthenticodeSignature -LiteralPath $normal).Status -ne 'NotSigned') { throw 'Expected an unsigned engineering artifact' }
if ((Get-FileHash -LiteralPath $normal).Hash -ne $record.installerSha256) { throw 'Prepared artifact changed' }
if ((Get-Item -LiteralPath $normal).VersionInfo.ProductName.Trim() -ne 'NADABAS Loader Test') { throw 'Backward-compatible product metadata is missing' }
$output = Join-Path $BuildDirectory ('unsigned-smoke-' + [guid]::NewGuid().ToString('N'))
$options = @('/DSmokeTest', '/DUnsignedBuild', '/DProductionCandidate', "/DPackageVersion=$($record.packageVersion)", "/DPayloadDir=$BuildDirectory\payload", "/DPackageOutput=$output")
$hashes = @{
    UpgradeHash = $record.previousTestSha256; ManagedUpgradeHash = $record.managedBridgeSha256
    DatedUpgradeHash = $record.datedTestSha256; BetaHash = $record.previousBetaSha256
    ContextBetaHash = $record.contextBetaSha256; DemoHash = $record.previousDemoSha256
    FixHash = $record.upgradeTestSha256
}
foreach ($entry in $hashes.GetEnumerator()) {
    if ($entry.Value -notmatch '^[A-Fa-f0-9]{64}$') { throw 'Invalid expected release hash' }
    $options += '/D' + $entry.Key + '=' + $entry.Value
}
& $Compiler @options (Join-Path $repo 'dotnet/packaging/loader-test.iss')
if ($LASTEXITCODE -ne 0) { throw 'Unsigned smoke compilation failed' }
$smoke = Join-Path $output "UNSIGNED-Nadabas-$($record.packageVersion)-Setup.exe"
if ((Get-AuthenticodeSignature -LiteralPath $smoke).Status -ne 'NotSigned' -or
    (Get-Item -LiteralPath $smoke).VersionInfo.ProductName.Trim() -ne 'NADABAS Installer Smoke') { throw 'Wrong isolated artifact identity' }
$parent = Join-Path $repo ('artifacts/unsigned-installer/test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $parent | Out-Null
$root = Join-Path $parent 'installation'
$blocked = Join-Path $parent 'must-not-install'
$checks = [Collections.Generic.List[string]]::new()
function Check([string]$name, [bool]$ok) { if (-not $ok) { throw "Failed: $name" }; $checks.Add($name) }
function Run-Setup([string]$file, [string]$directory, [string]$language) {
    $arguments = @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', ('/DIR="' + $directory + '"'), "/NADABASLANG=$language")
    (Start-Process -FilePath $file -ArgumentList $arguments -WindowStyle Hidden -Wait -PassThru).ExitCode
}
try {
    Check 'unsigned-normal-profile-refused' ((Run-Setup $normal $blocked '3') -ne 0)
    Check 'refused-package-writes-no-payload' (-not (Test-Path -LiteralPath $blocked))
    foreach ($language in @('3','4','5','6')) {
        Check "isolated-install-language-$language" ((Run-Setup $smoke $root $language) -eq 0)
        Check "language-$language-persisted" ((Get-ItemPropertyValue ($registry + '\Language') Code) -eq $language)
    }
    $selection = & (Join-Path $root 'tools/Nadabas.VersionStore.exe') inspect $root | ConvertFrom-Json
    Check 'approved-vba-selected' ($selection.Slot -eq 'slot-fix602009' -and $selection.ReleaseVersion -eq '6.02.009')
    $slots = @(Get-ChildItem -LiteralPath (Join-Path $root 'versions') -Directory | Select-Object -ExpandProperty Name)
    Check 'simulated-fixtures-not-distributed' (-not (Compare-Object @('slot-a','slot-b','slot-fix602009') $slots))
    Check 'user-guide-installed' ((Get-FileHash (Join-Path $root 'TESTGUIDE.txt')).Hash -eq (Get-FileHash (Join-Path $repo 'dotnet/packaging/NADABAS-USERGUIDE.txt')).Hash)
    Check 'package-version-registered' ((Get-ItemPropertyValue $uninstall DisplayVersion) -eq $record.packageVersion)
} finally {
    $remove = Join-Path $root 'unins000.exe'
    if (Test-Path -LiteralPath $remove) {
        $process = Start-Process -FilePath $remove -ArgumentList '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -WindowStyle Hidden -Wait -PassThru
        if ($process.ExitCode -ne 0) { throw 'Isolated uninstall failed; preserve state for inspection' }
    }
    # Exact synthetic key: preflight proved it absent. Never remove Office keys.
    if (Test-Path $registry) { Remove-Item -LiteralPath $registry -Recurse }
}
Check 'isolated-registration-removed' (-not (Test-Path $uninstall))
[ordered]@{ passed = $true; scope = 'unsigned-isolated-installer-not-office-trust'; checks = $checks.ToArray(); normalSha256 = $record.installerSha256; smokeSha256 = (Get-FileHash $smoke).Hash } |
    ConvertTo-Json -Depth 4 | Set-Content (Join-Path $parent 'report.json')
Write-Output "Passed $($checks.Count) checks. Report: $parent\report.json"
