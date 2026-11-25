param([string]$SignedLoader, [switch]$UnblockTestCopies, [string]$PreparedStore)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot
$source = 'C:\NADABAS\NADABAS.XLAM'
$normalAddin = Join-Path $env:APPDATA 'Microsoft\AddIns\NADABAS.XLAM'
$storeTool = Join-Path $repoRoot 'dotnet/src/Nadabas.VersionStore/bin/Release/net10.0/Nadabas.VersionStore.dll'
$builtLoader = Join-Path $repoRoot 'dotnet/src/Nadabas.ExcelBootstrap/bin/Release/net10.0-windows/publish/Nadabas.ExcelBootstrap-AddIn64-packed.xll'
if ($SignedLoader) { $builtLoader = [IO.Path]::GetFullPath($SignedLoader) }
$root = Join-Path $env:LOCALAPPDATA ('NBLoader-' + [guid]::NewGuid().ToString('N'))
if ($PreparedStore) {
    if ($UnblockTestCopies) { throw 'Prepared stores must not be unblocked again by this test.' }
    $root = [IO.Path]::GetFullPath($PreparedStore)
}
$reportRoot = Join-Path $repoRoot 'artifacts/version-loader'
$optionsPath = 'Software\Microsoft\Office\16.0\Excel\Options'
$managerPath = 'Software\Microsoft\Office\16.0\Excel\Add-in Manager'
$cases = [Collections.Generic.List[object]]::new()
$ownedProcesses = [Collections.Generic.List[Diagnostics.Process]]::new()
$testError = $null
$registrationWritten = $false
function Check([string]$name, [bool]$condition, $evidence) {
    $cases.Add([ordered]@{ name = $name; passed = $condition; evidence = $evidence })
    if (-not $condition) { throw "Check failed: $name" }
}
function Start-Session([string]$expectedSlot) {
    # A normal Excel launch, deliberately not COM /automation. Startup must find
    # our registered XLL without an explicit XLL argument or RegisterXLL call.
    $process = Start-Process -FilePath $excelPath -ArgumentList '/x' -PassThru
    $ownedProcesses.Add($process)
    $statePath = Join-Path $root "sessions/$($process.Id).json"
    for ($attempt = 0; $attempt -lt 120; $attempt++) {
        if ($process.HasExited) { throw 'Test Excel exited before reporting startup' }
        if (Test-Path -LiteralPath $statePath) {
            $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
            if ($state.status -eq 'failed') { throw "Loader failed: $($state.error)" }
            if ($state.status -eq 'loaded') {
                Check "startup-$expectedSlot-$($process.Id)" ($state.slot -eq $expectedSlot -and $state.runtimeVersion -eq '6.02.005') $state
                return $process
            }
        }
        Start-Sleep -Milliseconds 250
    }
    throw 'Loader did not report success. Inspect Excel for security/runtime messages; no trust settings were changed.'
}
function Stop-Session([Diagnostics.Process]$process, [switch]$Cleanup) {
    if (-not $process.HasExited) {
        if (-not $process.CloseMainWindow()) { throw "Could not request normal close of test Excel PID $($process.Id)" }
        if (-not $process.WaitForExit(30000)) { throw "Test Excel PID $($process.Id) remains open; no forced termination attempted" }
    }
    $stateFile = Join-Path $root "sessions/$($process.Id).json"
    if ($Cleanup -and -not (Test-Path -LiteralPath $stateFile)) { return }
    $state = Get-Content -LiteralPath $stateFile -Raw | ConvertFrom-Json
    # Excel-DNA AutoClose runs on explicit add-in removal, not normal shutdown.
    # The process exit is the shutdown evidence; the JSON is a startup snapshot.
    Check "normal-close-$($process.Id)" ($process.HasExited -and $state.status -in @('loaded', 'closed')) $state
}
function Activate([string]$slot) {
    & dotnet $storeTool activate $root $slot
    if ($LASTEXITCODE -ne 0) { throw 'Activation pointer update failed' }
}
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Save and close Excel before testing. Existing sessions are not touched.' }
$excelPath = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe').'(default)'
# Read PE machine type; only the 64-bit prototype is built.
$stream = [IO.File]::OpenRead($excelPath)
$reader = [IO.BinaryReader]::new($stream)
try { $stream.Position = 0x3c; $pe = $reader.ReadInt32(); $stream.Position = $pe + 4; if ($reader.ReadUInt16() -ne 0x8664) { throw 'This prototype requires x64 Excel' } }
finally { $reader.Dispose() }
if (-not (Test-Path -LiteralPath $builtLoader)) { throw 'Build the ExcelBootstrap project first' }
if ((Get-AuthenticodeSignature -LiteralPath $builtLoader).Status -ne 'Valid') {
    throw 'This acceptance test requires an Authenticode-signed loader. Supply -SignedLoader; no registration or Office security settings were changed.'
}
$beforeSource = (Get-FileHash -LiteralPath $source).Hash
$beforeNormal = (Get-FileHash -LiteralPath $normalAddin).Hash
$sourceZoneBefore = Get-Content -LiteralPath $source -Stream Zone.Identifier -Raw -ErrorAction SilentlyContinue
$normalZoneBefore = Get-Content -LiteralPath $normalAddin -Stream Zone.Identifier -Raw -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $reportRoot -Force | Out-Null
if ($PreparedStore) {
    & dotnet $storeTool inspect $root
} else {
    & dotnet $storeTool prepare $root $source
}
if ($LASTEXITCODE -ne 0) { throw 'Store preparation failed' }
if ($UnblockTestCopies) {
    # Explicit opt-in for the two generated, baseline-verified test copies only.
    foreach ($slot in @('slot-a', 'slot-b')) {
        $testCopy = Join-Path $root "versions/$slot/NADABAS.TEST.$slot.xlam"
        Check "verified-before-unblock-$slot" ((Get-FileHash -LiteralPath $testCopy).Hash -eq $beforeSource) $testCopy
        $attributes = [IO.File]::GetAttributes($testCopy)
        try {
            [IO.File]::SetAttributes($testCopy, ($attributes -band (-bnot [IO.FileAttributes]::ReadOnly)))
            Unblock-File -LiteralPath $testCopy
        } finally { [IO.File]::SetAttributes($testCopy, $attributes) }
        Check "unchanged-after-unblock-$slot" ((Get-FileHash -LiteralPath $testCopy).Hash -eq $beforeSource) $testCopy
        Check "test-zone-removed-$slot" (-not (Get-Item -LiteralPath $testCopy -Stream Zone.Identifier -ErrorAction SilentlyContinue)) $testCopy
    }
}
$loader = Join-Path $root 'Nadabas.Loader.Test.xll'
Copy-Item -LiteralPath $builtLoader -Destination $loader
$startupValue = '/R "' + $loader + '"'
$options = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($optionsPath)
$beforeOpen = @{}
foreach ($name in $options.GetValueNames()) { if ($name -match '^OPEN\d*$') { $beforeOpen[$name] = $options.GetValue($name) } }
try {
    $index = 0
    do { $entryName = if ($index -eq 0) { 'OPEN' } else { "OPEN$index" }; $index++ } while ($options.GetValueNames() -contains $entryName)
    $options.SetValue($entryName, $startupValue, [Microsoft.Win32.RegistryValueKind]::String)
    $registrationWritten = $true
    $first = Start-Session 'slot-a'
    Activate 'slot-b'
    $firstState = Get-Content -LiteralPath (Join-Path $root "sessions/$($first.Id).json") -Raw | ConvertFrom-Json
    Check 'open-session-keeps-slot-a' (-not $first.HasExited -and $firstState.slot -eq 'slot-a' -and $firstState.status -eq 'loaded') $firstState.addinPath
    $second = Start-Session 'slot-b'
    Check 'two-distinct-excel-processes' ($first.Id -ne $second.Id -and -not $first.HasExited -and -not $second.HasExited) @($first.Id, $second.Id)
    Stop-Session $second
    Stop-Session $first
    # Old-session shutdown must not revert the active pointer or lose the loader.
    $third = Start-Session 'slot-b'
    Stop-Session $third
    Activate 'slot-a'
    $fourth = Start-Session 'slot-a'
    Stop-Session $fourth
    foreach ($slot in @('slot-a', 'slot-b')) {
        $file = Join-Path $root "versions/$slot/NADABAS.TEST.$slot.xlam"
        Check "retained-unchanged-$slot" ((Get-FileHash -LiteralPath $file).Hash -eq $beforeSource) $file
    }
}
catch { $testError = $_.Exception.Message }
finally {
    foreach ($process in $ownedProcesses) {
        if (-not $process.HasExited) {
            try { Stop-Session $process -Cleanup } catch { $testError = "$testError; Cleanup: $($_.Exception.Message)" }
        }
    }
    $allClosed = @($ownedProcesses | Where-Object { -not $_.HasExited }).Count -eq 0
    # Remove only exact entries belonging to this unique test root. Never restore
    # whole registry keys or overwrite another add-in's settings.
    foreach ($section in @($optionsPath, $managerPath)) {
        $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($section, $true)
        if ($key) {
            try {
                foreach ($name in $key.GetValueNames()) {
                    $value = $key.GetValue($name)
                    $ownPaths = @($loader, (Join-Path $root 'versions/slot-a/NADABAS.TEST.slot-a.xlam'), (Join-Path $root 'versions/slot-b/NADABAS.TEST.slot-b.xlam'))
                    foreach ($ownPath in $ownPaths) {
                        if ($name -eq $ownPath -or ($name -match '^OPEN\d*$' -and $value -in @($ownPath, ('"' + $ownPath + '"'), ('/R "' + $ownPath + '"')))) { $key.DeleteValue($name, $false); break }
                    }
                }
            } finally { $key.Dispose() }
        }
    }
    $preserved = $true
    foreach ($name in $beforeOpen.Keys) { if ($options.GetValue($name) -ne $beforeOpen[$name]) { $preserved = $false } }
    $removed = $options.GetValueNames() | Where-Object { $options.GetValue($_) -eq $startupValue }
    $cases.Add([ordered]@{ name = 'test-registration-removed'; passed = (-not $removed -and $registrationWritten) })
    $cases.Add([ordered]@{ name = 'existing-open-entries-preserved'; passed = $preserved })
    $cases.Add([ordered]@{ name = 'test-excel-processes-closed'; passed = $allClosed })
    $cases.Add([ordered]@{ name = 'normal-addin-unchanged'; passed = ((Get-FileHash -LiteralPath $normalAddin).Hash -eq $beforeNormal) })
    $cases.Add([ordered]@{ name = 'source-unchanged'; passed = ((Get-FileHash -LiteralPath $source).Hash -eq $beforeSource) })
    $cases.Add([ordered]@{ name = 'source-zone-unchanged'; passed = ((Get-Content -LiteralPath $source -Stream Zone.Identifier -Raw -ErrorAction SilentlyContinue) -eq $sourceZoneBefore) })
    $cases.Add([ordered]@{ name = 'normal-zone-unchanged'; passed = ((Get-Content -LiteralPath $normalAddin -Stream Zone.Identifier -Raw -ErrorAction SilentlyContinue) -eq $normalZoneBefore) })
    $options.Dispose()
    $report = [ordered]@{ schemaVersion = 1; kind = 'version-loader-rehearsal'; passed = (-not $testError -and @($cases | Where-Object { -not $_.passed }).Count -eq 0); error = $testError; store = $root; startup = 'normal-excel-x-with-persistent-test-open-entry'; slots = 'identical-approved-6.02.005-not-two-releases'; cases = @($cases.ToArray()) }
    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $reportRoot 'latest-report.json') -Encoding utf8
    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $root 'test-report.json') -Encoding utf8
    Write-Output "REPORT: $reportRoot\latest-report.json"
}
if (-not $report.passed) { throw "Version loader test incomplete: $testError" }
Write-Output "Passed $($cases.Count) checks. Test startup entry removed; version files retained at $root."
