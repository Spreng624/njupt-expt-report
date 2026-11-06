param(
    [Parameter(Mandatory = $true)][string]$InputDocx,
    [Parameter(Mandatory = $true)][string]$OutputDirectory
)

$ErrorActionPreference = "Stop"
$inputPath = (Resolve-Path -LiteralPath $InputDocx).Path
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$pdfPath = Join-Path $OutputDirectory ((Split-Path -Leaf $inputPath) -replace '\.docx$', '.pdf')

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $document = $word.Documents.Open($inputPath, $false, $true)
    $document.ExportAsFixedFormat(
        $pdfPath,
        17,
        $false,
        0,
        0,
        0,
        0,
        0,
        $true,
        $true,
        0,
        $true,
        $true,
        $false
    )
    $document.Close($false)
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}

$pdftoppm = Get-Command pdftoppm -ErrorAction SilentlyContinue
if ($null -eq $pdftoppm) {
    Write-Output "PDF created: $pdfPath"
    Write-Output "Install Poppler if PNG page previews are required."
    exit 0
}

$prefix = Join-Path $OutputDirectory "page"
& $pdftoppm.Source -png -r 150 $pdfPath $prefix
Get-ChildItem -LiteralPath $OutputDirectory -Filter "page-*.png" |
    Select-Object Name, Length |
    Format-Table -AutoSize
