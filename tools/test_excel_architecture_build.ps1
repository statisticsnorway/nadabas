param([string]$Configuration = 'Release')
$ErrorActionPreference = 'Stop'
if ($Configuration -notin @('Debug', 'Release')) { throw 'Unknown build configuration' }
$repo = Split-Path $PSScriptRoot
$results = foreach ($project in @('Nadabas.ExcelStartup', 'Nadabas.ExcelBootstrap')) {
    $framework = if ($project -eq 'Nadabas.ExcelStartup') { 'net48' } else { 'net10.0-windows' }
    foreach ($architecture in @('x86', 'x64')) {
        $suffix = if ($architecture -eq 'x64') { '64' } else { '' }
        $path = Join-Path $repo "dotnet/src/$project/bin/$Configuration/$framework/publish/$project-AddIn$suffix-packed.xll"
        $stream = [IO.File]::OpenRead($path)
        $reader = [IO.BinaryReader]::new($stream)
        try {
            if ($reader.ReadUInt16() -ne 0x5a4d) { throw "Invalid DOS header: $path" }
            $stream.Position = 0x3c
            $offset = $reader.ReadInt32()
            if ($offset -lt 0x40 -or $offset -gt $stream.Length - 26) { throw "Invalid PE offset: $path" }
            $stream.Position = $offset
            if ($reader.ReadUInt32() -ne 0x4550) { throw "Invalid PE signature: $path" }
            $machine = $reader.ReadUInt16()
            $expected = if ($architecture -eq 'x64') { 0x8664 } else { 0x14c }
            if ($machine -ne $expected) { throw "Wrong architecture in $path" }
            $stream.Position = $offset + 22
            if (($reader.ReadUInt16() -band 0x2000) -eq 0) { throw "XLL is not a DLL image: $path" }
            $magic = $reader.ReadUInt16()
            $expectedMagic = if ($architecture -eq 'x64') { 0x20b } else { 0x10b }
            if ($magic -ne $expectedMagic) { throw "Wrong optional header in $path" }
        } finally { $reader.Dispose() }
        [ordered]@{ project = $project; architecture = $architecture; path = $path; sha256 = (Get-FileHash -LiteralPath $path).Hash }
    }
}
[ordered]@{ passed = $true; scope = 'compiled-XLL-architecture-only'; excelTested = $false; outputs = @($results) } | ConvertTo-Json -Depth 4
