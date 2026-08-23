param(
    [string]$Python = "python",
    [switch]$SkipTests,
    [switch]$KeepBuildVenv
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

if (-not $IsWindows -and $env:OS -ne "Windows_NT") {
    throw "WIN-1 must be built on Windows. PyInstaller cannot cross-compile a Windows executable."
}

$BuildVenv = Join-Path $RepoRoot ".build-venv"
$BuildPython = Join-Path $BuildVenv "Scripts\python.exe"

Write-Host "== Progress Studio WIN-1 portable build =="
Write-Host "Bootstrap Python: $Python"
& $Python -c "import sys; print(sys.executable); print(sys.version)"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

if (Test-Path $BuildVenv) {
    Write-Host "Removing previous isolated build environment..."
    Remove-Item $BuildVenv -Recurse -Force
}

Write-Host "Creating isolated build environment: $BuildVenv"
& $Python -m venv $BuildVenv
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
if (-not (Test-Path $BuildPython)) { throw "Build Python was not created: $BuildPython" }

try {
    Write-Host "Installing build + test dependencies inside isolated environment..."
    & $BuildPython -m pip install --disable-pip-version-check -e ".[build,dev]"
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    if (-not $SkipTests) {
        Write-Host "Running pre-build smoke gate in isolated environment..."
        & $BuildPython -m pytest -m smoke -q
        if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    }

    Remove-Item "$RepoRoot\build\win1" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "$RepoRoot\dist\ProgressStudio" -Recurse -Force -ErrorAction SilentlyContinue

    Write-Host "Building one-folder portable application..."
    & $BuildPython -m PyInstaller `
        --noconfirm `
        --clean `
        --workpath "$RepoRoot\build\win1" `
        --distpath "$RepoRoot\dist" `
        "$RepoRoot\packaging\windows\ProgressStudio.spec"
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    Write-Host "Verifying built portable application..."
    & $BuildPython "$RepoRoot\packaging\windows\verify_portable.py" "$RepoRoot\dist\ProgressStudio"
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    Write-Host ""
    Write-Host "WIN-1 PASS"
    Write-Host "Portable folder: $RepoRoot\dist\ProgressStudio"
    Write-Host "Executable:      $RepoRoot\dist\ProgressStudio\ProgressStudio.exe"
}
finally {
    if (-not $KeepBuildVenv -and (Test-Path $BuildVenv)) {
        Write-Host "Removing isolated build environment..."
        Remove-Item $BuildVenv -Recurse -Force -ErrorAction SilentlyContinue
    }
}
