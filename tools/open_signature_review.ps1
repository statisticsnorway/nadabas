param(
    [Parameter(Mandatory = $true)]
    [string]$ReviewDirectory,
    [Parameter(Mandatory = $true)]
    [ValidateSet('NADABAS.6.02.005.xlsm', 'NADABAS.XLAM')]
    [string]$FileName
)

$ErrorActionPreference = 'Stop'
$reviewRoot = (Resolve-Path -LiteralPath $ReviewDirectory).Path
$manifestPath = Join-Path $PSScriptRoot '../dotnet/releases/6.02.005.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$expected = @($manifest.files | Where-Object { $_.name -ceq $FileName })
if ($expected.Count -ne 1) { throw 'File is not uniquely identified in the release manifest.' }
$reviewFile = Join-Path $reviewRoot $FileName

function Assert-ReferenceFile {
    $item = Get-Item -LiteralPath $reviewFile
    if ($item.Length -ne $expected[0].length -or
        (Get-FileHash -LiteralPath $reviewFile -Algorithm SHA256).Hash -ne $expected[0].sha256) {
        throw 'Review copy differs from the recorded release. Do not open it with this launcher.'
    }
    if (-not $item.IsReadOnly) { throw 'The review copy must have its read-only attribute set.' }
}

Assert-ReferenceFile
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    throw 'Close Excel yourself after saving your work, then run this command again. No Excel processes were closed.'
}

# ForceDisable does not disable Excel 4.0 macros. Reject those package parts
# before starting Excel. This checks the supplied package, not signature validity.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::OpenRead($reviewFile)
try {
    foreach ($entry in $archive.Entries) {
        if ($entry.FullName -match '(?i)^xl/(macroSheets|intlMacroSheets)/') {
            throw 'Excel 4.0 macro sheets are outside this review setup.'
        }
        if ($entry.FullName -eq '[Content_Types].xml') {
            $reader = New-Object System.IO.StreamReader($entry.Open())
            try { $contentTypes = $reader.ReadToEnd() } finally { $reader.Dispose() }
            if ($contentTypes -match '(?i)(intlmacrosheet|macrosheet)') {
                throw 'Excel 4.0 macro content is outside this review setup.'
            }
        }
    }
}
finally { $archive.Dispose() }

$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.Visible = $true
    # UpdateLinks=0, ReadOnly=true. Do not call Run or register/install an add-in.
    $book = $excel.Workbooks.Open($reviewFile, 0, $true)
    Write-Host "Reviewing $FileName in Excel $($excel.Version)."
    Write-Host 'Press Alt+F11, select this file in Project Explorer, then Tools > Digital Signature.'
    Write-Host 'Inspect the existing signature/certificate only. Do not choose a certificate, sign, enable content or install.'
    Write-Host 'Record observations in RESULTS.md in the review directory.'
    Read-Host 'When finished, close any signature dialogs and press Enter here to close without saving' | Out-Null
}
finally {
    try {
        if ($null -ne $book) {
            try { $book.Close($false) }
            finally { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
        }
    }
    finally {
        if ($null -ne $excel) {
            try { $excel.Quit() }
            finally { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
        }
        Assert-ReferenceFile
    }
}
Write-Host 'Review finished. File hash still matches the release reference.'
