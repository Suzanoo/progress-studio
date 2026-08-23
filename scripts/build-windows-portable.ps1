param(
    [string]$Python = "python",
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

if (-not $IsWindows -and $env:OS -ne "Windows_NT") {
    throw "WIN-1 must be built on Windows. PyInstaller cannot cross-compile a Windows executable."
}

Write-Host "== Progress Studio WIN-1 portable build =="
& $Python -c "import sys; print(sys.version)"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "Installing/validating build dependencies..."
& $Python -m pip install -e ".[build]"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

if (-not $SkipTests) {
    Write-Host "Running pre-build smoke gate..."
    & $Python -m pytest -m smoke -q
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

Remove-Item "$RepoRoot\build\win1" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$RepoRoot\dist\ProgressStudio" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Building one-folder portable application..."
& $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --workpath "$RepoRoot\build\win1" `
    --distpath "$RepoRoot\dist" `
    "$RepoRoot\packaging\windows\ProgressStudio.spec"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "Verifying built portable application..."
& $Python "$RepoRoot\packaging\windows\verify_portable.py" "$RepoRoot\dist\ProgressStudio"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host ""
Write-Host "WIN-1 PASS"
Write-Host "Portable folder: $RepoRoot\dist\ProgressStudio"
Write-Host "Executable:      $RepoRoot\dist\ProgressStudio\ProgressStudio.exe"
