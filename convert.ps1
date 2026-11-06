param(
    [Parameter(Mandatory = $true)][string]$Markdown,
    [string]$Output,
    [string]$PythonExe
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

if ($PythonExe) {
    if (-not (Test-Path -LiteralPath $PythonExe)) {
        throw "Python executable not found: $PythonExe"
    }
    $pythonExe = (Resolve-Path -LiteralPath $PythonExe).Path
}
elseif ($env:PYTHON_EXE) {
    $pythonExe = $env:PYTHON_EXE
}
elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $pythonExe = (Get-Command python).Source
}
else {
    throw "Python was not found. Install Python or pass -PythonExe <path>."
}

$markdownPath = (Resolve-Path -LiteralPath $Markdown).Path
if (-not $Output) {
    $Output = Join-Path $root ("output\" + [IO.Path]::GetFileNameWithoutExtension($markdownPath) + ".docx")
}
elseif (-not [IO.Path]::IsPathRooted($Output)) {
    $Output = Join-Path $root $Output
}

$outputDirectory = Split-Path -Parent $Output
$markdownValidationJson = Join-Path $outputDirectory (
    [IO.Path]::GetFileNameWithoutExtension($markdownPath) + ".md.validation.json"
)
& $pythonExe (Join-Path $root "scripts\validate_markdown.py") $markdownPath --json $markdownValidationJson
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

& $pythonExe (Join-Path $root "scripts\build_report.py") $markdownPath -o $Output
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

$validationJson = [IO.Path]::ChangeExtension($Output, ".validation.json")
& $pythonExe (Join-Path $root "scripts\validate_docx.py") $Output --json $validationJson
exit $LASTEXITCODE
