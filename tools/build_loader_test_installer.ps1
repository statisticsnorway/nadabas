param(
    [Parameter(Mandatory)][string]$SignedLoader,
    [Parameter(Mandatory)][string]$SignedPackageUpdater,
    [Parameter(Mandatory)][string]$UpgradeAddin,
    [Parameter(Mandatory)][string]$ManagedUpgradeAddin,
    [Parameter(Mandatory)][string]$DatedUpgradeAddin,
    [Parameter(Mandatory)][string]$BetaAddin,
    [Parameter(Mandatory)][string]$ContextBetaAddin,
    [Parameter(Mandatory)][string]$DemoAddin,
    [Parameter(Mandatory)][string]$FixAddin,
    [Parameter(Mandatory)][string]$Compiler,
    [Parameter(Mandatory)][string]$SignTool,
    [Parameter(Mandatory)][ValidatePattern('^[A-Fa-f0-9]{40}$')][string]$CertificateThumbprint,
    [string]$ReleaseAddin = 'C:\NADABAS\NADABAS.XLAM',
    [ValidatePattern('^[0-9]+\.[0-9]+\.[0-9]+$')][string]$PackageVersion = '0.9.0',
    [string]$BuildParent,
    [switch]$CompactPayload
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot
$expected = '2F7432B9B04C414337EB1AFEB07E2B93F05E4A5A6D151073FA967A2F7EA89A85'
if ((Get-FileHash -LiteralPath $ReleaseAddin).Hash -ne $expected) { throw 'Approved XLAM baseline mismatch' }
$signature = Get-AuthenticodeSignature -LiteralPath $SignedLoader
if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Thumbprint -ne $CertificateThumbprint) { throw 'Loader must have a valid signature from the selected publisher' }
$updaterSignature = Get-AuthenticodeSignature -LiteralPath $SignedPackageUpdater
if ($updaterSignature.Status -ne 'Valid' -or $updaterSignature.SignerCertificate.Thumbprint -ne $CertificateThumbprint) { throw 'Package updater must have a valid publisher signature' }
$upgradeSignature = Get-AuthenticodeSignature -LiteralPath $UpgradeAddin
if ($upgradeSignature.Status -ne 'Valid' -or $upgradeSignature.SignerCertificate.Thumbprint -ne $CertificateThumbprint) { throw 'Upgrade VBA fixture must be signed by the selected publisher' }
$managedSignature = Get-AuthenticodeSignature -LiteralPath $ManagedUpgradeAddin
if ($managedSignature.Status -ne 'Valid' -or $managedSignature.SignerCertificate.Thumbprint -ne $CertificateThumbprint) { throw 'Managed VBA fixture must be signed by the selected publisher' }
$datedSignature = Get-AuthenticodeSignature -LiteralPath $DatedUpgradeAddin
if ($datedSignature.Status -ne 'Valid' -or $datedSignature.SignerCertificate.Thumbprint -ne $CertificateThumbprint) { throw 'Dated VBA fixture must be signed by the selected publisher' }
$betaSignature = Get-AuthenticodeSignature -LiteralPath $BetaAddin
if ($betaSignature.Status -ne 'Valid' -or $betaSignature.SignerCertificate.Thumbprint -ne $CertificateThumbprint) { throw 'Beta fixture must be signed by the selected publisher' }
$contextSignature = Get-AuthenticodeSignature -LiteralPath $ContextBetaAddin
if ($contextSignature.Status -ne 'Valid' -or $contextSignature.SignerCertificate.Thumbprint -ne $CertificateThumbprint) { throw 'Context beta must be signed by the selected publisher' }
$demoSignature = Get-AuthenticodeSignature -LiteralPath $DemoAddin
if ($demoSignature.Status -ne 'Valid' -or $demoSignature.SignerCertificate.Thumbprint -ne $CertificateThumbprint) { throw 'Demo must have a valid publisher signature' }
$fixSignature = Get-AuthenticodeSignature -LiteralPath $FixAddin
if ($fixSignature.Status -ne 'Valid' -or $fixSignature.SignerCertificate.Thumbprint -ne $CertificateThumbprint) { throw 'Issue 47 fixture must be signed by the selected publisher' }
$fixHash = (Get-FileHash -LiteralPath $FixAddin).Hash
$demoHash = (Get-FileHash -LiteralPath $DemoAddin).Hash
$contextHash = (Get-FileHash -LiteralPath $ContextBetaAddin).Hash
$betaHash = (Get-FileHash -LiteralPath $BetaAddin).Hash
$datedHash = (Get-FileHash -LiteralPath $DatedUpgradeAddin).Hash
$managedHash = (Get-FileHash -LiteralPath $ManagedUpgradeAddin).Hash
$upgradeHash = (Get-FileHash -LiteralPath $UpgradeAddin).Hash
if (-not $BuildParent) { $BuildParent = Join-Path $repo 'artifacts/loader-installer' }
if (-not [IO.Path]::IsPathFullyQualified($BuildParent)) { throw 'BuildParent must be absolute' }
$build = Join-Path $BuildParent ('build-' + [guid]::NewGuid().ToString('N'))
$payload = Join-Path $build 'payload'
$output = Join-Path $build 'package'
$env:DOTNET_CLI_HOME = Join-Path $repo '.codex-build/dotnet-home'
$env:NUGET_PACKAGES = Join-Path $repo '.codex-build/nuget'
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
New-Item -ItemType Directory -Path $build -Force | Out-Null
Push-Location (Join-Path $repo 'dotnet')
try {
    $tool = Join-Path $build 'tool'
    & dotnet publish src/Nadabas.VersionStore -c Release -r win-x64 --self-contained false -p:AppHostDotNetSearch=Global -p:DebugType=None -o $tool -m:1 --disable-build-servers
    if ($LASTEXITCODE -ne 0) { throw 'Version tool publish failed' }
    & (Join-Path $tool 'Nadabas.VersionStore.exe') prepare $payload $ReleaseAddin
    if ($LASTEXITCODE -ne 0) { throw 'Store preparation failed' }
    if (-not $CompactPayload) {
        & (Join-Path $tool 'Nadabas.VersionStore.exe') stage $payload slot-test602006 ([IO.Path]::GetFullPath($UpgradeAddin))
        if ($LASTEXITCODE -ne 0) { throw 'Upgrade fixture does not match the compiled approved release catalog' }
        & (Join-Path $tool 'Nadabas.VersionStore.exe') stage $payload slot-test602007 ([IO.Path]::GetFullPath($ManagedUpgradeAddin))
        if ($LASTEXITCODE -ne 0) { throw 'Managed fixture does not match the compiled catalog' }
        & (Join-Path $tool 'Nadabas.VersionStore.exe') stage $payload slot-test602008 ([IO.Path]::GetFullPath($DatedUpgradeAddin))
        if ($LASTEXITCODE -ne 0) { throw 'Dated fixture does not match the compiled catalog' }
        & (Join-Path $tool 'Nadabas.VersionStore.exe') stage $payload slot-beta602006 ([IO.Path]::GetFullPath($BetaAddin))
        if ($LASTEXITCODE -ne 0) { throw 'Beta fixture does not match the compiled catalog' }
        & (Join-Path $tool 'Nadabas.VersionStore.exe') stage $payload slot-beta602006b ([IO.Path]::GetFullPath($ContextBetaAddin))
        if ($LASTEXITCODE -ne 0) { throw 'Context beta does not match the compiled catalog' }
        & (Join-Path $tool 'Nadabas.VersionStore.exe') stage $payload slot-demo602006 ([IO.Path]::GetFullPath($DemoAddin))
        if ($LASTEXITCODE -ne 0) { throw 'Demo fixture does not match the compiled catalog' }
    }
    & (Join-Path $tool 'Nadabas.VersionStore.exe') stage $payload slot-fix602009 ([IO.Path]::GetFullPath($FixAddin))
    if ($LASTEXITCODE -ne 0) { throw 'Issue 47 fixture does not match the compiled catalog' }
    $expectedSlots = @('slot-a', 'slot-b', 'slot-fix602009')
    if (-not $CompactPayload) { $expectedSlots += @('slot-test602006', 'slot-test602007', 'slot-test602008', 'slot-beta602006', 'slot-beta602006b', 'slot-demo602006') }
    $actualSlots = @(Get-ChildItem -LiteralPath (Join-Path $payload 'versions') -Directory | Select-Object -ExpandProperty Name)
    if (Compare-Object ($expectedSlots | Sort-Object) ($actualSlots | Sort-Object)) { throw 'Unexpected packaged VBA slots' }
    Copy-Item -LiteralPath $tool -Destination (Join-Path $payload 'tools') -Recurse
    New-Item -ItemType Directory -Path (Join-Path $payload 'updater') | Out-Null
    Copy-Item -LiteralPath $SignedPackageUpdater -Destination (Join-Path $payload 'updater/Nadabas.PackageUpdate.exe')
    Copy-Item -LiteralPath $SignedLoader -Destination (Join-Path $payload 'Nadabas.Loader.Test.xll')
    Copy-Item -LiteralPath (Join-Path $repo 'dotnet/packaging/LOADER-TESTGUIDE.txt') -Destination (Join-Path $payload 'TESTGUIDE.txt')
    # Inno embeds file data, not Zone.Identifier. This package is intentionally
    # signed; no blanket Unblock-File or Office security modification is used.
    foreach ($slot in @('slot-a', 'slot-b')) {
        if ((Get-FileHash -LiteralPath (Join-Path $payload "versions/$slot/NADABAS.TEST.$slot.xlam")).Hash -ne $expected) { throw 'Packaged slot hash mismatch' }
    }
    $signCommand = '$q' + $SignTool + '$q sign /sha1 ' + $CertificateThumbprint + ' /s My /fd SHA256 /tr http://timestamp.globalsign.com/tsa/r45standard /td SHA256 $f'
    $payloadOptions = @()
    if ($CompactPayload) { $payloadOptions += '/DCompactPayload' }
    & $Compiler @payloadOptions "/DPackageVersion=$PackageVersion" "/Snadabas_test=$signCommand" "/DUpgradeHash=$upgradeHash" "/DManagedUpgradeHash=$managedHash" "/DDatedUpgradeHash=$datedHash" "/DBetaHash=$betaHash" "/DContextBetaHash=$contextHash" "/DDemoHash=$demoHash" "/DFixHash=$fixHash" "/DPayloadDir=$payload" "/DPackageOutput=$output" (Join-Path $repo 'dotnet/packaging/loader-test.iss')
    if ($LASTEXITCODE -ne 0) { throw 'Signed installer build failed; do not distribute intermediate output' }
    $installer = Join-Path $output "NadabasLoaderTest-$PackageVersion-Setup.exe"
    $installerSignature = Get-AuthenticodeSignature -LiteralPath $installer
    if ($installerSignature.Status -ne 'Valid' -or $installerSignature.SignerCertificate.Thumbprint -ne $CertificateThumbprint) { throw 'Installer signature verification failed' }
    [ordered]@{
        packageVersion = $PackageVersion; testOnly = $true; installer = $installer
        compactPayload = [bool]$CompactPayload; packagedSlots = $actualSlots
        installerSha256 = (Get-FileHash -LiteralPath $installer).Hash
        signedPackageUpdaterSha256 = (Get-FileHash -LiteralPath $SignedPackageUpdater).Hash
        signedLoaderSha256 = (Get-FileHash -LiteralPath $SignedLoader).Hash
        certificateThumbprint = $CertificateThumbprint; baselineSha256 = $expected
        upgradeTestVersion = '6.02.009'; upgradeTestSha256 = $fixHash; previousDemoSha256 = $demoHash; contextBetaSha256 = $contextHash; previousBetaSha256 = $betaHash; releaseChannel = 'internal-exchange-beta'; datedTestSha256 = $datedHash; managedBridgeSha256 = $managedHash; previousTestSha256 = $upgradeHash; publishedVbaRelease = $false
        acceptance = 'pending-user-session'; excelRegistration = 'automatic-current-user-loader-only'
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'build-record.json') -Encoding utf8
    Write-Output "SIGNED TEST INSTALLER: $installer"
} finally { Pop-Location }
