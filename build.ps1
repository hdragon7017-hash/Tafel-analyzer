[CmdletBinding()]
param(
    [string]$PythonExe = "",
    [string]$OutputDirectory = "dist"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvPath = Join-Path $projectRoot ".build-venv"

if ([string]::IsNullOrWhiteSpace($PythonExe)) {
    $PythonExe = Join-Path $projectRoot ".venv\Scripts\python.exe"
}

if (-not (Test-Path -LiteralPath $pythonExe)) {
    py -3 -m venv $venvPath
    if ($LASTEXITCODE -ne 0) { throw "Failed to create the build virtual environment." }
    $PythonExe = Join-Path $venvPath "Scripts\python.exe"
}
$PythonExe = (Resolve-Path -LiteralPath $PythonExe).Path
$outputPath = if ([System.IO.Path]::IsPathRooted($OutputDirectory)) {
    $OutputDirectory
} else {
    Join-Path $projectRoot $OutputDirectory
}
$workPath = Join-Path $projectRoot "build"

& $PythonExe -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw "Failed to upgrade pip." }
& $PythonExe -m pip install -r (Join-Path $projectRoot "requirements-lock.txt")
if ($LASTEXITCODE -ne 0) { throw "Failed to install locked build dependencies." }
Push-Location $projectRoot
try {
    & $PythonExe -m PyInstaller --noconfirm --clean --distpath $outputPath --workpath $workPath TafelAnalyzer.spec
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller build failed." }
}
finally {
    Pop-Location
}

Write-Host "Created: $(Join-Path $outputPath 'TafelAnalyzer.exe')"
