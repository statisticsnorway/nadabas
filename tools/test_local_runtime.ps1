param(
    [string]$InstalledDotnet = 'C:\Program Files\dotnet',
    [string]$RuntimeVersion = '10.0.11',
    [switch]$OmitDesktopRuntime
)
$ErrorActionPreference = 'Stop'
if ($RuntimeVersion -notmatch '^10\.0\.\d+$') { throw 'Use a stable .NET 10 runtime version' }
$repo = Split-Path $PSScriptRoot
$run = Join-Path $repo ('artifacts/local-runtime/test-' + [guid]::NewGuid().ToString('N'))
$localRuntime = Join-Path $run 'runtime'
$nativeHost = Join-Path $InstalledDotnet "packs/Microsoft.NETCore.App.Host.win-x64/$RuntimeVersion/runtimes/win-x64/native/nethost.dll"
$inputs = @('dotnet.exe', "host/fxr/$RuntimeVersion", "shared/Microsoft.NETCore.App/$RuntimeVersion", "shared/Microsoft.WindowsDesktop.App/$RuntimeVersion")
foreach ($relative in $inputs) {
    if (-not (Test-Path -LiteralPath (Join-Path $InstalledDotnet $relative))) { throw "Missing source: $relative" }
}
if (-not (Test-Path -LiteralPath $nativeHost)) { throw 'The engineering probe needs the matching SDK nethost library' }
New-Item -ItemType Directory -Path $localRuntime | Out-Null
foreach ($relative in $inputs) {
    if ($OmitDesktopRuntime -and $relative.StartsWith('shared/Microsoft.WindowsDesktop.App/')) { continue }
    $destination = Join-Path $localRuntime $relative
    New-Item -ItemType Directory -Force -Path (Split-Path $destination) | Out-Null
    Copy-Item -LiteralPath (Join-Path $InstalledDotnet $relative) -Destination $destination -Recurse
}
Copy-Item -LiteralPath $nativeHost -Destination (Join-Path $run 'nethost.dll')
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'runtime_probe.cs') -Destination (Join-Path $run 'runtime_probe.cs')
@{ runtimeOptions = @{ tfm = 'net10.0'; framework = @{ name = 'Microsoft.WindowsDesktop.App'; version = '10.0.0' } } } |
    ConvertTo-Json -Depth 4 | Set-Content (Join-Path $run 'probe.runtimeconfig.json')
# All environment changes are confined to a new child process. No registry or Office writes.
$child = @'
$ErrorActionPreference = 'Stop'
try {
    $run = (Get-Location).Path
    $env:DOTNET_ROOT = Join-Path $run 'runtime'
    $env:DOTNET_ROOT_X64 = $env:DOTNET_ROOT
    $env:DOTNET_HOST_TRACE = '1'
    $env:DOTNET_HOST_TRACEFILE = Join-Path $run 'host-trace.txt'
    Add-Type -Path (Join-Path $run 'runtime_probe.cs')
    $dependency = [NadabasRuntimeProbe]::Run((Join-Path $run 'nethost.dll'), $env:DOTNET_ROOT, (Join-Path $run 'probe.runtimeconfig.json'))
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    $elevated = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($elevated) { throw 'Repeat this acceptance test without elevation' }
    @{ passed = $true; elevated = $elevated; runtimeDependency = $dependency; scope = 'native-host-discovery-and-activation-only'; excelTested = $false; cleanMachineTested = $false; persistentEnvironmentChanged = $false } |
        ConvertTo-Json | Set-Content -Encoding UTF8 (Join-Path $run 'report.json')
    exit 0
} catch {
    $_.Exception.Message | Set-Content (Join-Path (Get-Location).Path 'failure.txt')
    exit 1
}
'@
$encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($child))
$process = Start-Process -FilePath "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -WorkingDirectory $run `
    -ArgumentList @('-NoLogo', '-NoProfile', '-NonInteractive', '-EncodedCommand', $encoded) -WindowStyle Hidden -PassThru
$deadline = [DateTime]::UtcNow.AddSeconds(180)
while (-not $process.WaitForExit(5000)) {
    if ([DateTime]::UtcNow -ge $deadline) { throw "Probe still running (PID $($process.Id)); inspect $run" }
}
if ($OmitDesktopRuntime) {
    $failure = Get-Content -LiteralPath (Join-Path $run 'failure.txt') -Raw
    # FrameworkMissingFailure: the private host must not borrow the global Desktop Runtime.
    if ($process.ExitCode -eq 0 -or $failure -notlike '*Local Desktop Runtime initialization failed: -2147450730*') {
        throw "Missing-runtime rejection was not proven; inspect $run"
    }
    @{ passed = $true; scope = 'missing-private-desktop-runtime-rejected'; cleanMachineTested = $false } |
        ConvertTo-Json | Set-Content -Encoding UTF8 (Join-Path $run 'report.json')
} elseif ($process.ExitCode -ne 0) { throw "Probe failed; inspect $run\failure.txt" }
Get-Content -LiteralPath (Join-Path $run 'report.json')
Write-Output "Report: $run\report.json"
