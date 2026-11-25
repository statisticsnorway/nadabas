param(
    [Parameter(Mandatory)][string]$RuntimeDirectory,
    [Parameter(Mandatory)][string]$PreparedPayload,
    [Parameter(Mandatory)][string]$SignedLoader,
    [Parameter(Mandatory)][string]$SignedUpdater
)
$ErrorActionPreference = 'Stop'
$principal = [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Run acceptance without elevation' }
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Save and close Excel first; existing sessions are never touched' }
foreach ($path in @($SignedLoader, $SignedUpdater)) {
    $signature = Get-AuthenticodeSignature -LiteralPath $path
    if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Thumbprint -ne 'A70F7882B2F9F5C6D23A22C8AEC70211854B9EAF') {
        throw 'Use existing verified SSB-signed test components'
    }
}
$runtime = (Resolve-Path -LiteralPath $RuntimeDirectory).Path.TrimEnd('\')
if (-not (Test-Path -LiteralPath (Join-Path $runtime 'dotnet.exe'))) { throw 'Missing private runtime' }
$repo = Split-Path $PSScriptRoot
$run = Join-Path $repo ('artifacts/excel-local-runtime/test-' + [guid]::NewGuid().ToString('N'))
$root = Join-Path $run 'installation'
New-Item -ItemType Directory -Path $run | Out-Null
Copy-Item -LiteralPath $PreparedPayload -Destination $root -Recurse
Copy-Item -LiteralPath $SignedLoader -Destination (Join-Path $root 'Nadabas.Loader.Test.xll') -Force
Copy-Item -LiteralPath $SignedUpdater -Destination (Join-Path $root 'updater/Nadabas.PackageUpdate.exe') -Force
& (Join-Path $runtime 'dotnet.exe') (Join-Path $root 'tools/Nadabas.VersionStore.dll') activate $root slot-fix602009
if ($LASTEXITCODE -ne 0) { throw 'Cannot activate the approved test payload' }
$fixture = Join-Path $run 'blank.csv'
'NADABAS isolated runtime test' | Set-Content -LiteralPath $fixture
$excelPath = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe').'(default)'
$start = [Diagnostics.ProcessStartInfo]::new($excelPath)
$start.UseShellExecute = $false
# /automation avoids ordinary startup add-ins; this is not normal-launch acceptance.
$start.Arguments = '/x /automation "' + $fixture + '" "' + (Join-Path $root 'Nadabas.Loader.Test.xll') + '"'
$start.Environment['DOTNET_ROOT'] = $runtime
$start.Environment['DOTNET_ROOT_X64'] = $runtime
$start.Environment['DOTNET_HOST_TRACE'] = '1'
$start.Environment['DOTNET_HOST_TRACEFILE'] = Join-Path $run 'host-trace.txt'
$owned = [Diagnostics.Process]::Start($start)
@{ pid = $owned.Id; started = $owned.StartTime.ToUniversalTime(); scope = 'isolated-explicit-XLL-with-process-local-runtime' } |
    ConvertTo-Json | Set-Content (Join-Path $run 'owner.json')
$statePath = Join-Path $root "sessions/$($owned.Id).json"
try {
    $deadline = [DateTime]::UtcNow.AddSeconds(120)
    $state = $null
    while (-not $owned.HasExited -and [DateTime]::UtcNow -lt $deadline) {
        if (Test-Path -LiteralPath $statePath) {
            $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
            if ($state.status -in @('loaded', 'failed')) { break }
        }
        Start-Sleep -Milliseconds 500
    }
    $modules = @($owned.Modules | Where-Object { $_.ModuleName -in @('hostfxr.dll', 'coreclr.dll') } | ForEach-Object { $_.FileName })
    $privateModules = $modules.Count -eq 2 -and @($modules | Where-Object { -not $_.StartsWith($runtime + '\', [StringComparison]::OrdinalIgnoreCase) }).Count -eq 0
    $passed = $state.status -eq 'loaded' -and $state.runtimeVersion -eq '6.02.009' -and $privateModules
    @{ passed = $passed; scope = 'isolated-explicit-XLL-with-process-local-runtime'; normalStartupTested = $false; cleanMachineTested = $false; privateRuntimeModules = $modules; session = $state } |
        ConvertTo-Json -Depth 5 | Set-Content (Join-Path $run 'report.json')
    if (-not $passed) { throw "Excel local-runtime acceptance failed; inspect $run" }
    Write-Output "Passed isolated Excel loading. Report: $run\report.json"
} finally {
    if (-not $owned.HasExited) {
        if (-not $owned.CloseMainWindow() -or -not $owned.WaitForExit(15000)) {
            Write-Warning "Test Excel PID $($owned.Id) remains open. Close it normally; no process was killed. Report directory: $run"
        }
    }
}
