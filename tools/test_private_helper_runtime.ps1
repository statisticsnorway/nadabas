param(
    [Parameter(Mandatory)][string]$Updater,
    [Parameter(Mandatory)][string]$RuntimeDirectory,
    [ValidateSet('x86', 'x64')][string]$Architecture = 'x64'
)
$ErrorActionPreference = 'Stop'
$updaterPath = (Resolve-Path -LiteralPath $Updater).Path
if ([IO.Path]::GetFileName($updaterPath) -ne 'Nadabas.PackageUpdate.exe') { throw 'Use the engineering updater build' }
$runtime = (Resolve-Path -LiteralPath $RuntimeDirectory).Path
function Assert-Architecture([string]$Path) {
    $stream = [IO.File]::OpenRead($Path)
    $reader = [IO.BinaryReader]::new($stream)
    try {
        if ($reader.ReadUInt16() -ne 0x5a4d) { throw "Invalid DOS header: $Path" }
        $stream.Position = 0x3c
        $offset = $reader.ReadInt32()
        if ($offset -lt 0x40 -or $offset -gt $stream.Length - 26) { throw "Invalid PE offset: $Path" }
        $stream.Position = $offset
        if ($reader.ReadUInt32() -ne 0x4550) { throw "Invalid PE signature: $Path" }
        $expected = if ($Architecture -eq 'x64') { 0x8664 } else { 0x14c }
        if ($reader.ReadUInt16() -ne $expected) { throw "Expected $Architecture image: $Path" }
        $stream.Position = $offset + 24
        $expectedMagic = if ($Architecture -eq 'x64') { 0x20b } else { 0x10b }
        if ($reader.ReadUInt16() -ne $expectedMagic) { throw "Invalid optional header: $Path" }
    } finally { $reader.Dispose() }
}
Assert-Architecture $updaterPath
Assert-Architecture (Join-Path $runtime 'dotnet.exe')
$repo = Split-Path $PSScriptRoot
$run = Join-Path $repo ('artifacts/private-helper-runtime/test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $run | Out-Null
$results = foreach ($valid in @($true, $false)) {
    $label = if ($valid) { 'private-runtime' } else { 'missing-runtime' }
    $trace = Join-Path $run "$label.trace.txt"
    $start = [Diagnostics.ProcessStartInfo]::new($updaterPath)
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $selected = if ($valid) { $runtime } else { Join-Path $run 'absent' }
    $start.Environment['DOTNET_ROOT'] = $selected
    $start.Environment['DOTNET_ROOT_' + $Architecture.ToUpperInvariant()] = $selected
    $start.Environment['DOTNET_DISABLE_GUI_ERRORS'] = '1'
    $start.Environment['DOTNET_HOST_TRACE'] = '1'
    $start.Environment['DOTNET_HOST_TRACEFILE'] = $trace
    # No arguments: exercises runtime loading and the read-only usage path only.
    $process = [Diagnostics.Process]::Start($start)
    $output = $process.StandardOutput.ReadToEndAsync()
    $errorOutput = $process.StandardError.ReadToEndAsync()
    $deadline = [DateTime]::UtcNow.AddSeconds(120)
    while (-not $process.WaitForExit(5000)) {
        if ([DateTime]::UtcNow -ge $deadline) { throw "Probe PID $($process.Id) is still running; inspect $run" }
    }
    $errorText = $errorOutput.GetAwaiter().GetResult()
    $null = $output.GetAwaiter().GetResult()
    $traceText = Get-Content -LiteralPath $trace -Raw
    if ($valid) {
        if ($process.ExitCode -ne 2 -or $errorText -notlike '*INTERNAL TEST:*' -or
            -not $traceText.Contains((Join-Path $runtime 'shared\Microsoft.NETCore.App'))) {
            throw 'Private runtime did not reach updater usage entry point'
        }
    } elseif ($process.ExitCode -eq 0 -or $errorText -like '*INTERNAL TEST:*') {
        throw 'Missing private runtime incorrectly reached managed updater code'
    }
    [ordered]@{ name = $label; passed = $true; exitCode = $process.ExitCode }
}
[ordered]@{ passed = $true; scope = 'engineering-apphost-runtime-only'; architecture = $Architecture; excelTested = $false; installationTested = $false; updaterSha256 = (Get-FileHash -LiteralPath $updaterPath).Hash; checks = @($results) } |
    ConvertTo-Json -Depth 4 | Set-Content (Join-Path $run 'report.json')
Write-Output "Passed. Report: $run\report.json"
