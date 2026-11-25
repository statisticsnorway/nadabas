param(
    [Parameter(Mandatory)][string]$RuntimeX86,
    [Parameter(Mandatory)][string]$RuntimeX64,
    [Parameter(Mandatory)][string]$Output
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot
if (-not [IO.Path]::IsPathFullyQualified($Output)) { throw 'Use an absolute new output directory' }
$env:DOTNET_CLI_HOME = Join-Path $repo '.codex-build/dotnet-home'
$env:NUGET_PACKAGES = Join-Path $repo '.codex-build/nuget'
function Invoke-Build([string[]]$BuildArguments) {
    & dotnet @BuildArguments
    if ($LASTEXITCODE -ne 0) { throw 'Component build failed; retain the incomplete output for diagnosis' }
}
& (Join-Path $repo '.venv/Scripts/python.exe') (Join-Path $PSScriptRoot 'prepare_runtime_bundle.py') --runtime-x86 $RuntimeX86 --runtime-x64 $RuntimeX64 --output $Output
if ($LASTEXITCODE -ne 0) { throw 'Runtime preparation failed' }
$manifest = Join-Path $Output 'runtime-manifest.tsv'
$components = @()
foreach ($arch in @('x86', 'x64')) {
    $payload = Join-Path $Output "payloads/$arch"
    foreach ($project in @('Nadabas.PackageUpdate', 'Nadabas.VersionStore')) {
        $folder = if ($project -eq 'Nadabas.PackageUpdate') { 'updater' } else { 'tools' }
        $options = @('publish', (Join-Path $repo "dotnet/src/$project"), '-c', 'Release', '-r', "win-$arch", '--self-contained', 'false', '-p:PrivateRuntimeHelpers=true', '-p:RestoreLockedMode=true', '-p:DebugType=None', '-o', (Join-Path $payload $folder), '-m:1', '--disable-build-servers')
        if ($project -eq 'Nadabas.PackageUpdate') { $options += @('-p:PublishSingleFile=true', "-p:RuntimeManifest=$manifest") }
        Invoke-Build -BuildArguments $options
        $file = Join-Path $payload "$folder/$project.exe"
        $components += [ordered]@{ architecture = $arch; component = $project; path = $file; sha256 = (Get-FileHash -LiteralPath $file).Hash }
    }
}
Invoke-Build -BuildArguments @('build', (Join-Path $repo 'dotnet/src/Nadabas.ExcelBootstrap'), '-c', 'Release', '-p:RestoreLockedMode=true', "-p:RuntimeManifest=$manifest", '-p:PrivateRuntimeHelpers=true', '-m:1', '--disable-build-servers')
$startupLines = @(Get-Content -LiteralPath $manifest | ForEach-Object {
    $fields = $_.Split("`t")
    $fields[0] + "`t" + 'runtime/' + $fields[0] + '/' + $fields[1] + "`t" + $fields[2]
})
foreach ($arch in @('x86', 'x64')) {
    $suffix = if ($arch -eq 'x64') { '64' } else { '' }
    $loader = Join-Path $Output "payloads/$arch/Nadabas.Loader.$arch.xll"
    Copy-Item -LiteralPath (Join-Path $repo "dotnet/src/Nadabas.ExcelBootstrap/bin/Release/net10.0-windows/publish/Nadabas.ExcelBootstrap-AddIn$suffix-packed.xll") -Destination $loader
    $hash = (Get-FileHash -LiteralPath $loader).Hash
    $startupLines += $arch + "`t" + "Nadabas.Loader.$arch.xll" + "`t" + $hash
    $components += [ordered]@{ architecture = $arch; component = 'modern-loader'; path = $loader; sha256 = $hash }
}
$startupManifest = Join-Path $Output 'startup-manifest.UNSIGNED.tsv'
$startupLines | Set-Content -LiteralPath $startupManifest -Encoding utf8
Invoke-Build -BuildArguments @('build', (Join-Path $repo 'dotnet/src/Nadabas.ExcelStartup'), '-c', 'Release', '-p:RestoreLockedMode=true', "-p:StartupManifest=$startupManifest", '-m:1', '--disable-build-servers')
foreach ($arch in @('x86', 'x64')) {
    $suffix = if ($arch -eq 'x64') { '64' } else { '' }
    $startup = Join-Path $Output "payloads/$arch/Nadabas.Startup.xll"
    Copy-Item -LiteralPath (Join-Path $repo "dotnet/src/Nadabas.ExcelStartup/bin/Release/net48/publish/Nadabas.ExcelStartup-AddIn$suffix-packed.xll") -Destination $startup
    $components += [ordered]@{ architecture = $arch; component = 'framework-startup'; path = $startup; sha256 = (Get-FileHash -LiteralPath $startup).Hash }
}
[ordered]@{
    status = 'unsigned-private-components-prepared'; canDistribute = $false
    sourceCommit = (& git -C $repo rev-parse HEAD); sourceHasLocalChanges = [bool](& git -C $repo status --porcelain)
    runtimeManifestSha256 = (Get-FileHash -LiteralPath $manifest).Hash
    components = $components
    remaining = @('Review runtime provenance and redistribution files', 'Sign modern payload, regenerate startup hashes, rebuild and sign startup', 'Add approved XLAM store and installer architecture selection', 'Signed desktop installation and upgrade acceptance')
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $Output 'component-build.json') -Encoding utf8
Write-Output "Unsigned engineering components: $Output"
