param(
    [Parameter(Mandatory)][string]$ComponentBuild,
    [Parameter(Mandatory)][string]$PreviousBuildRecord,
    [Parameter(Mandatory)][string]$BaselineAddin,
    [Parameter(Mandatory)][string]$FixAddin,
    [string]$AccessFixAddin,
    [Parameter(Mandatory)][string]$Compiler,
    [Parameter(Mandatory)][string]$Output,
    [switch]$SmokeTest,
    [string]$SignTool,
    [ValidatePattern('^[A-Fa-f0-9]{40}$')][string]$CertificateThumbprint
)
# Signed preparation requires the rebuilt startup manifest and verified components.
$ErrorActionPreference = 'Stop'
$signed = [bool]$SignTool
if ($signed -and -not $CertificateThumbprint) { throw 'A publisher certificate is required for signing' }
function Assert-Publisher([string]$Path) {
    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Thumbprint -ne $CertificateThumbprint) { throw "Publisher verification failed: $Path" }
}
$repo = Split-Path $PSScriptRoot
if (-not [IO.Path]::IsPathFullyQualified($Output) -or (Test-Path -LiteralPath $Output)) { throw 'Output must be a new absolute directory' }
$components = Get-Content (Join-Path $ComponentBuild 'component-build.json') -Raw | ConvertFrom-Json
$bundle = Get-Content (Join-Path $ComponentBuild 'bundle-record.json') -Raw | ConvertFrom-Json
$manifest = Join-Path $ComponentBuild 'runtime-manifest.tsv'
if ((Get-FileHash $manifest).Hash -ne $components.runtimeManifestSha256) { throw 'Runtime manifest changed' }
$runtimeVersions = @(Get-ChildItem (Join-Path $ComponentBuild 'payloads/x64/runtime/x64/host/fxr') -Directory)
if ($runtimeVersions.Count -ne 1 -or $runtimeVersions[0].Name -notmatch '^10\.0\.\d+$') { throw 'Expected one stable runtime patch' }
$runtimeVersion = $runtimeVersions[0].Name
if (@($bundle.architectures).Count -ne 2 -or @($bundle.architectures | Where-Object runtimeVersion -NE $runtimeVersion).Count) { throw 'Runtime patch versions differ' }
foreach ($line in Get-Content $manifest) {
    if (-not $line.Trim() -or $line.StartsWith('#')) { continue }
    $parts = $line.Split("`t")
    if ($parts.Count -ne 3 -or $parts[0] -notin @('x86','x64') -or $parts[1] -match '(^|[/\\])\.\.([/\\]|$)|:') { throw 'Invalid runtime inventory' }
    $file = Join-Path $ComponentBuild ("payloads/$($parts[0])/runtime/$($parts[0])/" + $parts[1])
    if ((Get-FileHash -LiteralPath $file).Hash -ne $parts[2]) { throw 'Runtime payload changed' }
}
$previous = Get-Content -LiteralPath $PreviousBuildRecord -Raw | ConvertFrom-Json
if ($components.canDistribute) { throw 'Expected an internal component build' }
if ($signed) {
    if ($components.status -ne 'signed-private-components-prepared') { throw 'Expected signed components with rebuilt startup hashes' }
    $startupManifest = Join-Path $ComponentBuild 'startup-manifest.SIGNED.tsv'
    if ((Get-FileHash -LiteralPath $startupManifest).Hash -ne $components.startupManifestSha256) { throw 'Signed startup inventory changed' }
    $startupLines = @(Get-Content -LiteralPath $startupManifest)
    foreach ($loader in $components.components | Where-Object component -EQ 'modern-loader') {
        $expectedLine = $loader.architecture + "`t" + "Nadabas.Loader.$($loader.architecture).xll" + "`t" + $loader.sha256
        if ($startupLines -cnotcontains $expectedLine) { throw 'Startup does not pin final signed loader bytes' }
    }
} elseif ($components.status -ne 'unsigned-private-components-prepared') { throw 'Expected unsigned component preparation' }
foreach ($item in $components.components) {
    if ((Get-FileHash -LiteralPath $item.path).Hash -ne $item.sha256) { throw 'Prepared component changed' }
    if ($signed) { Assert-Publisher $item.path }
}
$baseline = '2F7432B9B04C414337EB1AFEB07E2B93F05E4A5A6D151073FA967A2F7EA89A85'
$fix = 'D53D492724997C08596DBEACAE4695B30786B3695310D13035CD61541A829750'
foreach ($entry in @(@($BaselineAddin,$baseline), @($FixAddin,$fix))) {
    if ((Get-FileHash -LiteralPath $entry[0]).Hash -ne $entry[1]) { throw 'Approved VBA identity mismatch' }
}
$accessFix = '36A32EBF31DE42256B089228588F12E6A04D0481DDA3D9D8CC39E88D65C5FEF7'
if ($AccessFixAddin -and (Get-FileHash -LiteralPath $AccessFixAddin).Hash -ne $accessFix) { throw 'Access repair identity mismatch' }
if ($signed) {
    Assert-Publisher $BaselineAddin
    Assert-Publisher $FixAddin
    if ($AccessFixAddin) { Assert-Publisher $AccessFixAddin }
}
New-Item -ItemType Directory -Path $Output | Out-Null
$payload = Join-Path $Output 'payload'
$runtime = Join-Path $ComponentBuild 'payloads/x64/runtime/x64/dotnet.exe'
$tool = Join-Path $ComponentBuild 'payloads/x64/tools/Nadabas.VersionStore.dll'
& $runtime $tool prepare $payload $BaselineAddin
if ($LASTEXITCODE -ne 0) { throw 'Store preparation failed' }
& $runtime $tool stage $payload slot-fix602009 $FixAddin
if ($LASTEXITCODE -ne 0) { throw 'Approved VBA staging failed' }
if ($AccessFixAddin) {
    & $runtime $tool stage $payload slot-fix602010 $AccessFixAddin
    if ($LASTEXITCODE -ne 0) { throw 'Access repair staging failed' }
}
Copy-Item (Join-Path $repo 'dotnet/packaging/NADABAS-USERGUIDE.txt') (Join-Path $payload 'TESTGUIDE.txt')
$options = @('/DProductionCandidate','/DPrivateRuntime','/DPackageVersion=1.0.0',
    "/DPrivateRuntimeVersion=$runtimeVersion", "/DPayloadDir=$payload", "/DComponentsDir=$ComponentBuild\payloads", "/DPackageOutput=$Output\package")
if ($SmokeTest) { $options += '/DSmokeTest' }
if ($AccessFixAddin) { $options += "/DAccessFixHash=$accessFix" }
if ($signed) {
    $signCommand = '$q' + [IO.Path]::GetFullPath($SignTool) + '$q sign /sha1 ' + $CertificateThumbprint + ' /s My /fd SHA256 /tr http://timestamp.globalsign.com/tsa/r45standard /td SHA256 $f'
    $options += '/Snadabas_test=' + $signCommand
} else { $options += '/DUnsignedBuild' }
$hashes = @{
    UpgradeHash=$previous.previousTestSha256; ManagedUpgradeHash=$previous.managedBridgeSha256
    DatedUpgradeHash=$previous.datedTestSha256; BetaHash=$previous.previousBetaSha256
    ContextBetaHash=$previous.contextBetaSha256; DemoHash=$previous.previousDemoSha256; FixHash=$fix
}
foreach ($entry in $hashes.GetEnumerator()) {
    if ($entry.Value -notmatch '^[A-Fa-f0-9]{64}$') { throw 'Invalid historical slot identity' }
    $options += '/D' + $entry.Key + '=' + $entry.Value
}
& $Compiler @options (Join-Path $repo 'dotnet/packaging/loader-test.iss')
if ($LASTEXITCODE -ne 0) { throw 'Installer compilation failed' }
$fileName = if ($signed) { 'Nadabas-1.0.0-Setup.exe' } else { 'UNSIGNED-Nadabas-1.0.0-Setup.exe' }
$installer = Join-Path $Output "package/$fileName"
if ($signed) { Assert-Publisher $installer }
elseif ((Get-AuthenticodeSignature $installer).Status -ne 'NotSigned') { throw 'Expected unsigned artifact' }
$size = (Get-Item $installer).Length
[ordered]@{
    status=$(if ($signed) { 'signed-private-installer-candidate' } else { 'unsigned-private-installer' }); canDistribute=$false; smokeTest=[bool]$SmokeTest
    installer=$installer; sha256=(Get-FileHash $installer).Hash; bytes=$size
    withinUpdateLimit=($size -le 128MB); componentBuild=$ComponentBuild
    architectures=@('x86','x64'); vbaVersion=$(if ($AccessFixAddin) { '6.02.010' } else { '6.02.009' })
} | ConvertTo-Json | Set-Content (Join-Path $Output 'build-record.json')
Write-Output "Internal installer prepared: $installer ($size bytes)"
