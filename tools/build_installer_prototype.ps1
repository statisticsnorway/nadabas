param(
    [Parameter(Mandatory)][string]$Compiler,
    [string]$ReleaseAddin = 'C:\NADABAS\NADABAS.XLAM'
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot
$outputRoot = Join-Path $repoRoot 'artifacts/installer-prototype'
$buildRoot = Join-Path $outputRoot ('build-' + [guid]::NewGuid().ToString('N'))
$expected = '2F7432B9B04C414337EB1AFEB07E2B93F05E4A5A6D151073FA967A2F7EA89A85'
if ((Get-FileHash -LiteralPath $ReleaseAddin -Algorithm SHA256).Hash -ne $expected) { throw 'Release baseline mismatch' }
if (-not (Test-Path -LiteralPath $Compiler -PathType Leaf)) { throw 'Supply the Inno Setup ISCC.exe path' }
$env:DOTNET_CLI_HOME = Join-Path $repoRoot '.codex-build/dotnet-home'
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
$env:NUGET_PACKAGES = Join-Path $repoRoot '.codex-build/nuget'
Push-Location (Join-Path $repoRoot 'dotnet')
try {
    $diagnostics = Join-Path $buildRoot 'diagnostics'
    & dotnet publish src/Nadabas.Setup -c Release -r win-x64 --self-contained true -o $diagnostics -m:1 --disable-build-servers
    if ($LASTEXITCODE -ne 0) { throw 'Diagnostic publish failed' }
    foreach ($version in @('0.1.0', '0.2.0')) {
        $payload = Join-Path $buildRoot "payload/$version"
        $updater = Join-Path $payload 'updater'
        & dotnet publish src/Nadabas.UpdateProbe -c Release -r win-x64 --self-contained true "-p:Version=$version" -o $updater -m:1 --disable-build-servers
        if ($LASTEXITCODE -ne 0) { throw 'Update probe publish failed' }
        Copy-Item -LiteralPath $diagnostics -Destination (Join-Path $payload 'diagnostics') -Recurse -Force
        $releaseDir = Join-Path $payload 'releases/6.02.005'
        New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null
        Copy-Item -LiteralPath $ReleaseAddin -Destination (Join-Path $releaseDir 'NADABAS.XLAM') -Force
        if ((Get-FileHash -LiteralPath (Join-Path $releaseDir 'NADABAS.XLAM')).Hash -ne $expected) { throw 'Payload verification failed' }
        [ordered]@{ prototype = $true; packageVersion = $version; nadabasVersion = '6.02.005'; addinSha256 = $expected; excelRegistration = 'not-installed'; publisherSignature = 'installer-unsigned-test-only' } |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $payload 'prototype.json') -Encoding utf8
        $packageDir = Join-Path $outputRoot "packages/$version"
        & $Compiler "/DPackageVersion=$version" "/DPayloadDir=$payload" "/DPackageOutput=$packageDir" (Join-Path $repoRoot 'dotnet/packaging/prototype.iss')
        if ($LASTEXITCODE -ne 0) { throw 'Installer compilation failed' }
    }
}
finally { Pop-Location }
Write-Output "Prototype installers: $outputRoot\packages"
