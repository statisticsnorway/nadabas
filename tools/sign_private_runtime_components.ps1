param(
    [Parameter(Mandatory)][string]$ComponentBuild,
    [Parameter(Mandatory)][string]$SignTool,
    [Parameter(Mandatory)][ValidatePattern('^[A-Fa-f0-9]{40}$')][string]$CertificateThumbprint
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot
$recordPath = Join-Path $ComponentBuild 'component-build.json'
$record = Get-Content -LiteralPath $recordPath -Raw | ConvertFrom-Json
if ($record.status -ne 'unsigned-private-components-prepared' -or $record.canDistribute) { throw 'Use a fresh unsigned component build' }
$manifest = Join-Path $ComponentBuild 'runtime-manifest.tsv'
if ((Get-FileHash -LiteralPath $manifest).Hash -ne $record.runtimeManifestSha256) { throw 'Runtime manifest changed' }
foreach ($component in $record.components) {
    if ((Get-FileHash -LiteralPath $component.path).Hash -ne $component.sha256) { throw 'Component changed before signing' }
}
Copy-Item -LiteralPath $recordPath -Destination (Join-Path $ComponentBuild 'component-build.UNSIGNED.json') -ErrorAction Stop
function Sign-Verified([string]$Path) {
    & $SignTool sign /sha1 $CertificateThumbprint /s My /fd SHA256 /tr http://timestamp.globalsign.com/tsa/r45standard /td SHA256 $Path
    if ($LASTEXITCODE -ne 0) { throw 'Component signing failed; do not package this incomplete build' }
    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Thumbprint -ne $CertificateThumbprint) { throw 'Component publisher verification failed' }
}
$final = @()
foreach ($component in $record.components | Where-Object component -NE 'framework-startup') {
    Sign-Verified $component.path
    $component.sha256 = (Get-FileHash -LiteralPath $component.path).Hash
    $final += $component
}
# VersionStore is framework-dependent: sign the managed entry point as well.
foreach ($arch in @('x86','x64')) {
    $dll = Join-Path $ComponentBuild "payloads/$arch/tools/Nadabas.VersionStore.dll"
    Sign-Verified $dll
    $final += [pscustomobject]@{architecture=$arch; component='version-store-managed'; path=$dll; sha256=(Get-FileHash -LiteralPath $dll).Hash}
}
$lines = @(Get-Content -LiteralPath $manifest | ForEach-Object {
    $parts = $_.Split("`t")
    $parts[0] + "`t" + 'runtime/' + $parts[0] + '/' + $parts[1] + "`t" + $parts[2]
})
foreach ($component in $final | Where-Object component -EQ 'modern-loader') {
    $lines += $component.architecture + "`t" + "Nadabas.Loader.$($component.architecture).xll" + "`t" + $component.sha256
}
$startupManifest = Join-Path $ComponentBuild 'startup-manifest.SIGNED.tsv'
$lines | Set-Content -LiteralPath $startupManifest -Encoding utf8
$env:DOTNET_CLI_HOME = Join-Path $repo '.codex-build/dotnet-home'
$env:NUGET_PACKAGES = Join-Path $repo '.codex-build/nuget'
& dotnet build (Join-Path $repo 'dotnet/src/Nadabas.ExcelStartup') -c Release -p:RestoreLockedMode=true "-p:StartupManifest=$startupManifest" -m:1 --disable-build-servers
if ($LASTEXITCODE -ne 0) { throw 'Signed startup rebuild failed' }
foreach ($component in $record.components | Where-Object component -EQ 'framework-startup') {
    $suffix = if ($component.architecture -eq 'x64') { '64' } else { '' }
    Copy-Item -LiteralPath (Join-Path $repo "dotnet/src/Nadabas.ExcelStartup/bin/Release/net48/publish/Nadabas.ExcelStartup-AddIn$suffix-packed.xll") -Destination $component.path
    Sign-Verified $component.path
    $component.sha256 = (Get-FileHash -LiteralPath $component.path).Hash
    $final += $component
}
$record.status = 'signed-private-components-prepared'
$record.components = $final
$record.remaining = @('Signed installer build and desktop acceptance', 'Release source review and publication approval')
$record | Add-Member -NotePropertyName startupManifestSha256 -NotePropertyValue (Get-FileHash -LiteralPath $startupManifest).Hash
$record | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $recordPath -Encoding utf8
Write-Output 'Signed components verified; startup pins final signed loader bytes. Not approved for distribution.'
