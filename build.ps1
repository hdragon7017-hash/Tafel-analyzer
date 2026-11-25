[CmdletBinding()]
param(
    [string]$PythonExe = "",
    [string]$OutputDirectory = "dist",
    [string]$UPXDirectory = ""
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvPath = Join-Path $projectRoot ".build-venv"

if (-not [string]::IsNullOrWhiteSpace($UPXDirectory)) {
    $resolvedUPXDirectory = if ([System.IO.Path]::IsPathRooted($UPXDirectory)) {
        $UPXDirectory
    } else {
        Join-Path $projectRoot $UPXDirectory
    }
    $resolvedUPXDirectory = (Resolve-Path -LiteralPath $resolvedUPXDirectory).Path
    if (-not (Test-Path -LiteralPath (Join-Path $resolvedUPXDirectory "upx.exe"))) {
        throw "UPXDirectory must contain upx.exe: $resolvedUPXDirectory"
    }
    $env:PATH = "$resolvedUPXDirectory;$env:PATH"
}

$upxCommand = Get-Command upx.exe -ErrorAction SilentlyContinue
if ($null -ne $upxCommand) {
    Write-Host "UPX compression enabled: $($upxCommand.Source)"
} else {
    Write-Warning "UPX was not found; the program will still build, but its binary dependencies will not be UPX-compressed."
}

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
