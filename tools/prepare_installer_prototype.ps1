$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot
$toolRoot = Join-Path $repoRoot 'artifacts/installer-prototype/tooling'
New-Item -ItemType Directory -Path $toolRoot -Force | Out-Null
$compiler = Join-Path $toolRoot 'inno/ISCC.exe'
if (-not (Test-Path -LiteralPath $compiler)) {
    $download = Join-Path $toolRoot 'innosetup-7.1.0-x64.exe'
    Invoke-WebRequest 'https://github.com/jrsoftware/issrc/releases/download/is-7_1_0/innosetup-7.1.0-x64.exe' -OutFile $download
    if ((Get-FileHash -LiteralPath $download).Hash -ne '0362A383ED217D4C4239B5933866DD96D3EB2102737DA92F80F6057A4B40DF2F') { throw 'Compiler hash mismatch' }
    if ((Get-AuthenticodeSignature -LiteralPath $download).Status -ne 'Valid') { throw 'Compiler Authenticode verification failed' }
    $destination = Join-Path $toolRoot 'inno'
    $process = Start-Process -FilePath $download -ArgumentList ('/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /CURRENTUSER /PORTABLE=1 /DIR="' + $destination + '"') -WindowStyle Hidden -PassThru -Wait
    if ($process.ExitCode -ne 0) { throw 'Portable compiler extraction failed' }
}
$env:DOTNET_CLI_HOME = Join-Path $repoRoot '.codex-build/dotnet-home'
$env:NUGET_PACKAGES = Join-Path $repoRoot '.codex-build/nuget'
$generatorRoot = Join-Path $toolRoot 'netsparkle'
if (-not (Test-Path -LiteralPath (Join-Path $generatorRoot 'netsparkle-generate-appcast.exe'))) {
    & dotnet tool install NetSparkleUpdater.Tools.AppCastGenerator --version 2.9.0 --tool-path $generatorRoot --allow-roll-forward --add-source https://api.nuget.org/v3/index.json
    if ($LASTEXITCODE -ne 0) { throw 'NetSparkle tool installation failed' }
}
Write-Output "Compiler: $compiler"
