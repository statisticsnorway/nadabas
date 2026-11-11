param()
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot
$prototypeRoot = Join-Path $repoRoot 'artifacts/installer-prototype'
$runRoot = Join-Path $prototypeRoot ('run-' + [guid]::NewGuid().ToString('N'))
$installRoot = Join-Path $env:LOCALAPPDATA 'NADABAS-Installation-Prototype'
$registryPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{20788D4E-ED0E-4DBC-A2B2-3043CE45C9D7}_is1'
$normalAddin = Join-Path $env:APPDATA 'Microsoft\AddIns\NADABAS.XLAM'
$sourceAddin = 'C:\NADABAS\NADABAS.XLAM'
$generator = Join-Path $prototypeRoot 'tooling/netsparkle/netsparkle-generate-appcast.exe'
$python = Join-Path $repoRoot '.venv/Scripts/python.exe'
$cases = [Collections.Generic.List[object]]::new()
$server = $null
$ownedInstall = $false
$testError = $null
function Check([string]$name, [bool]$condition, $evidence) {
    $cases.Add([ordered]@{ name = $name; passed = $condition; evidence = $evidence })
    if (-not $condition) { throw "Check failed: $name" }
}
function Run-Installer([string]$path, [string]$logName) {
    $log = Join-Path $runRoot $logName
    $process = Start-Process -FilePath $path -ArgumentList ('/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LOG="' + $log + '"') -WindowStyle Hidden -Wait -PassThru
    if ($process.ExitCode -ne 0) { throw "Installer failed ($($process.ExitCode)); see $log" }
}
function Probe([string]$name, [int]$expectedExit, [string]$expectedStatus, [string]$keyPath) {
    $outputDir = Join-Path $runRoot $name
    $lines = & (Join-Path $installRoot 'updater/Nadabas.UpdateProbe.exe') --feed "$baseUrl/appcast.xml" --public-key $keyPath --output $outputDir
    $probeExit = $LASTEXITCODE
    $lines | Set-Content -LiteralPath (Join-Path $runRoot "$name.json") -Encoding utf8
    $result = ($lines -join "`n") | ConvertFrom-Json
    Check $name ($probeExit -eq $expectedExit -and $result.status -eq $expectedStatus) $result
    return $result
}
function Excel-Smoke([string]$name) {
    # Excel has tighter path limits than .NET. Keep the actual XLAM path short;
    # the complete JSON evidence is still retained in the main test report.
    $excelOutput = Join-Path ([IO.Path]::GetTempPath()) ('NB-' + [guid]::NewGuid().ToString('N'))
    $lines = & (Join-Path $installRoot 'diagnostics/Nadabas.Setup.exe') excel-startup-test --source (Join-Path $installRoot 'releases/6.02.005/NADABAS.XLAM') --output $excelOutput
    $smokeExit = $LASTEXITCODE
    $result = ($lines -join "`n") | ConvertFrom-Json
    Check $name ($smokeExit -eq 0 -and $result.addinActivated -and $result.registrationCleaned -and $result.sessionClosed -and $result.sourceUnchanged) $result
}
if ((Test-Path -LiteralPath $installRoot) -or (Test-Path -LiteralPath $registryPath)) {
    throw 'Prototype destination/registration already exists; inspect it before another installation test.'
}
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -ne 0) { throw 'Save and close Excel before running the prototype test.' }
foreach ($version in @('0.1.0', '0.2.0')) {
    if (-not (Test-Path -LiteralPath (Join-Path $prototypeRoot "packages/$version/NadabasPrototype-$version-Setup.exe"))) { throw 'Build both prototype packages first.' }
}
New-Item -ItemType Directory -Path $runRoot -Force | Out-Null
$beforeNormal = (Get-FileHash -LiteralPath $normalAddin).Hash
$beforeSource = (Get-FileHash -LiteralPath $sourceAddin).Hash
try {
    $feedDir = Join-Path $runRoot 'feed'
    $keys = Join-Path $runRoot 'test-keys'
    New-Item -ItemType Directory -Path $feedDir,$keys -Force | Out-Null
    $packageName = 'NadabasPrototype-0.2.0-Setup.exe'
    $package = Join-Path $feedDir $packageName
    Copy-Item -LiteralPath (Join-Path $prototypeRoot "packages/0.2.0/$packageName") -Destination $package
    $listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
    $listener.Start()
    $port = $listener.LocalEndpoint.Port
    $listener.Stop()
    $baseUrl = "http://127.0.0.1:$port"
    & $generator --generate-keys --key-path $keys
    if ($LASTEXITCODE -ne 0) { throw 'Test key generation failed' }
    & $generator --single-file $package --file-version 0.2.0 --key-path $keys --base-url $baseUrl --appcast-output-directory $feedDir --product-name 'NADABAS TEST ONLY' --human-readable true
    if ($LASTEXITCODE -ne 0) { throw 'Signed test feed generation failed' }
    $publicKey = Join-Path $keys 'NetSparkle_Ed25519.pub'
    $server = Start-Process -FilePath $python -ArgumentList ('-m http.server ' + $port + ' --bind 127.0.0.1 --directory "' + $feedDir + '"') -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $runRoot 'server.stdout.log') -RedirectStandardError (Join-Path $runRoot 'server.stderr.log')
    $ready = $false
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        if ($server.HasExited) { throw 'Local feed server exited' }
        try { $null = Invoke-WebRequest "$baseUrl/appcast.xml" -TimeoutSec 2; $ready = $true; break } catch { Start-Sleep -Milliseconds 100 }
    }
    if (-not $ready) { throw 'Local feed server did not become ready' }
    # From here cleanup owns only this newly created prototype registration/path.
    $ownedInstall = $true
    Run-Installer (Join-Path $prototypeRoot 'packages/0.1.0/NadabasPrototype-0.1.0-Setup.exe') 'install.log'
    $installed = Get-ItemProperty -LiteralPath $registryPath
    Check 'normal-windows-registration' ($installed.DisplayVersion -eq '0.1.0') $installed.DisplayName
    $versionText = & (Join-Path $installRoot 'updater/Nadabas.UpdateProbe.exe') --version
    Check 'installed-v1-runs' ($LASTEXITCODE -eq 0 -and "$versionText" -eq '0.1.0.0') "$versionText"
    $userDir = Join-Path $installRoot 'user-data'
    New-Item -ItemType Directory -Path $userDir | Out-Null
    $sentinel = Join-Path $userDir 'preserve-me.txt'
    'User data must survive upgrade and uninstall.' | Set-Content -LiteralPath $sentinel
    $sentinelHash = (Get-FileHash -LiteralPath $sentinel).Hash
    Excel-Smoke 'excel-from-installed-v1'
    $verified = Probe 'signed-update-accepted' 0 'verified-download' $publicKey
    Check 'download-matches-built-package' ($verified.sha256 -eq (Get-FileHash -LiteralPath $package).Hash) $verified.sha256
    $feed = Join-Path $feedDir 'appcast.xml'
    $feedBytes = [IO.File]::ReadAllBytes($feed)
    try {
        [IO.File]::AppendAllText($feed, "`n<!-- tampered -->")
        $null = Probe 'tampered-feed-rejected' 1 'rejected' $publicKey
        Check 'tampered-feed-no-download' (-not (Test-Path -LiteralPath (Join-Path $runRoot 'tampered-feed-rejected'))) 'No accepted package'
    } finally { [IO.File]::WriteAllBytes($feed, $feedBytes) }
    $originalLength = (Get-Item -LiteralPath $package).Length
    try {
        $stream = [IO.File]::Open($package, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try { $stream.WriteByte(1) } finally { $stream.Dispose() }
        $null = Probe 'tampered-package-rejected' 1 'rejected' $publicKey
        Check 'tampered-package-no-download' (-not (Test-Path -LiteralPath (Join-Path $runRoot 'tampered-package-rejected'))) 'No accepted package'
    } finally {
        $stream = [IO.File]::OpenWrite($package)
        try { $stream.SetLength($originalLength) } finally { $stream.Dispose() }
    }
    Run-Installer $verified.package 'upgrade.log'
    $installed = Get-ItemProperty -LiteralPath $registryPath
    $versionText = & (Join-Path $installRoot 'updater/Nadabas.UpdateProbe.exe') --version
    Check 'upgraded-to-v2' ($installed.DisplayVersion -eq '0.2.0' -and "$versionText" -eq '0.2.0.0') "$versionText"
    Check 'upgrade-preserves-user-data' ((Get-FileHash -LiteralPath $sentinel).Hash -eq $sentinelHash) $sentinel
    $null = Probe 'already-current' 0 'up-to-date' $publicKey
    Excel-Smoke 'excel-from-installed-v2'
    Run-Installer (Join-Path $installRoot 'unins000.exe') 'uninstall.log'
    Check 'uninstall-removes-registration' (-not (Test-Path -LiteralPath $registryPath)) $registryPath
    Check 'uninstall-removes-program' (-not (Test-Path -LiteralPath (Join-Path $installRoot 'updater/Nadabas.UpdateProbe.exe'))) 'Updater removed'
    Check 'uninstall-preserves-user-data' ((Get-FileHash -LiteralPath $sentinel).Hash -eq $sentinelHash) $sentinel
    $ownedInstall = $false
}
catch { $testError = $_.Exception.Message }
finally {
    if ($server -and -not $server.HasExited) { $server.Kill(); $server.WaitForExit() }
    if ($ownedInstall -and (Test-Path -LiteralPath (Join-Path $installRoot 'unins000.exe'))) {
        try { Run-Installer (Join-Path $installRoot 'unins000.exe') 'cleanup-uninstall.log' }
        catch { $testError = "$testError; Cleanup: $($_.Exception.Message)" }
    }
    $cases.Add([ordered]@{ name = 'normal-addin-unchanged'; passed = ((Get-FileHash -LiteralPath $normalAddin).Hash -eq $beforeNormal) })
    $cases.Add([ordered]@{ name = 'release-source-unchanged'; passed = ((Get-FileHash -LiteralPath $sourceAddin).Hash -eq $beforeSource) })
    $report = [ordered]@{ schemaVersion = 1; kind = 'installer-prototype'; passed = (-not $testError -and @($cases | Where-Object { -not $_.passed }).Count -eq 0); error = $testError; runDirectory = $runRoot; installer = 'Inno Setup 7.1.0'; updater = 'NetSparkle 3.1.0'; environment = 'current-user-profile-not-clean-vm'; publisherSigning = 'unsigned-test-installers-with-test-Ed25519-feed'; persistentExcelRegistration = $false; cases = @($cases.ToArray()) }
    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $runRoot 'report.json') -Encoding utf8
    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $prototypeRoot 'latest-report.json') -Encoding utf8
    Write-Output "REPORT: $runRoot\report.json"
}
if (-not $report.passed) { throw "Prototype acceptance failed: $testError" }
Write-Output "Passed $($cases.Count) checks. Prototype uninstalled; test user-data retained at $installRoot\user-data."
