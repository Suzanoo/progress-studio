param(
    [string]$PortableFolder = "",
    [string]$WorkRoot = "",
    [switch]$KeepCopy
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot

if (-not $IsWindows -and $env:OS -ne "Windows_NT") {
    throw "WIN-2 validation must run on Windows."
}

if ([string]::IsNullOrWhiteSpace($PortableFolder)) {
    $PortableFolder = Join-Path $RepoRoot "dist\ProgressStudio"
}
$PortableFolder = (Resolve-Path $PortableFolder).Path
$SourceExe = Join-Path $PortableFolder "ProgressStudio.exe"
if (-not (Test-Path $SourceExe -PathType Leaf)) {
    throw "ProgressStudio.exe was not found in portable folder: $PortableFolder"
}

if ([string]::IsNullOrWhiteSpace($WorkRoot)) {
    $WorkRoot = Join-Path $env:TEMP "ProgressStudio-WIN2"
}
New-Item -ItemType Directory -Path $WorkRoot -Force | Out-Null

$Stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$RunRoot = Join-Path $WorkRoot "portable-$Stamp"
$ReportPath = Join-Path $WorkRoot "WIN2-$Stamp.txt"

Write-Host "== Progress Studio WIN-2 isolated portable validation =="
Write-Host "Source portable folder: $PortableFolder"
Write-Host "Isolated run folder:    $RunRoot"
Write-Host ""
Write-Host "Copying portable bundle away from the source repository..."
Copy-Item -Path $PortableFolder -Destination $RunRoot -Recurse

$RunExe = Join-Path $RunRoot "ProgressStudio.exe"
if (-not (Test-Path $RunExe -PathType Leaf)) {
    throw "Isolated copy is missing ProgressStudio.exe: $RunExe"
}

$ForbiddenNames = @(".git", ".venv", ".build-venv", ".pytest_cache", "tests", "build")
foreach ($Name in $ForbiddenNames) {
    if (Test-Path (Join-Path $RunRoot $Name)) {
        throw "Portable bundle contains development-only path: $Name"
    }
}

$FileCount = (Get-ChildItem $RunRoot -Recurse -File | Measure-Object).Count
$BundleBytes = (Get-ChildItem $RunRoot -Recurse -File | Measure-Object -Property Length -Sum).Sum
$ExeHash = (Get-FileHash $RunExe -Algorithm SHA256).Hash

# Save and scrub Python/development environment variables. The executable must
# survive using only its own PyInstaller runtime and normal Windows system DLLs.
$SavedEnv = @{}
$NamesToClear = @(
    "PYTHONHOME",
    "PYTHONPATH",
    "VIRTUAL_ENV",
    "CONDA_PREFIX",
    "CONDA_DEFAULT_ENV",
    "PIP_REQUIRE_VIRTUALENV"
)
foreach ($Name in $NamesToClear) {
    $SavedEnv[$Name] = [Environment]::GetEnvironmentVariable($Name, "Process")
    Remove-Item "Env:$Name" -ErrorAction SilentlyContinue
}
$SavedPath = $env:PATH
$env:PATH = "$env:SystemRoot\System32;$env:SystemRoot"

try {
    Write-Host "Running packaged executable with Python/venv environment scrubbed..."
    $Process = Start-Process `
        -FilePath $RunExe `
        -ArgumentList "--win1-smoke" `
        -WorkingDirectory $RunRoot `
        -Wait `
        -PassThru

    if ($Process.ExitCode -ne 0) {
        throw "Portable executable isolation probe failed with exit code $($Process.ExitCode)."
    }
}
finally {
    $env:PATH = $SavedPath
    foreach ($Name in $NamesToClear) {
        $Value = $SavedEnv[$Name]
        if ($null -eq $Value) {
            Remove-Item "Env:$Name" -ErrorAction SilentlyContinue
        }
        else {
            [Environment]::SetEnvironmentVariable($Name, $Value, "Process")
        }
    }
}

$Report = @(
    "Progress Studio WIN-2 isolated portable validation",
    "Timestamp: $(Get-Date -Format o)",
    "OS: $([System.Environment]::OSVersion.VersionString)",
    "Source: $PortableFolder",
    "Isolated copy: $RunRoot",
    "File count: $FileCount",
    "Bundle bytes: $BundleBytes",
    "ProgressStudio.exe SHA256: $ExeHash",
    "Environment scrub: PYTHONHOME/PYTHONPATH/VIRTUAL_ENV/CONDA/PATH",
    "Executable probe: PASS",
    "",
    "Manual clean-machine acceptance is still required; see packaging/windows/WIN2_CHECKLIST.md."
)
$Report | Set-Content -Path $ReportPath -Encoding UTF8

Write-Host ""
Write-Host "WIN-2 AUTOMATED ISOLATION PROBE PASS"
Write-Host "Files:  $FileCount"
Write-Host "Size:   $BundleBytes bytes"
Write-Host "SHA256: $ExeHash"
Write-Host "Report: $ReportPath"
Write-Host ""
Write-Host "Next: copy the portable folder to a Windows user/VM with no Python environment and complete WIN2_CHECKLIST.md."

if ($KeepCopy) {
    Write-Host "Isolated copy kept at: $RunRoot"
}
else {
    Remove-Item $RunRoot -Recurse -Force -ErrorAction SilentlyContinue
}
