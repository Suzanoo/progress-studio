param(
    [string]$PortableFolder = "",
    [string]$InnoCompiler = "",
    [switch]$SkipPortableProbe
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

if (-not $IsWindows -and $env:OS -ne "Windows_NT") {
    throw "WIN-3 installer must be built on Windows."
}

if (-not $PortableFolder) {
    $PortableFolder = Join-Path $RepoRoot "dist\ProgressStudio"
}
$PortableFolder = (Resolve-Path $PortableFolder).Path
$Exe = Join-Path $PortableFolder "ProgressStudio.exe"
if (-not (Test-Path $Exe -PathType Leaf)) {
    throw "Known-good portable executable not found: $Exe"
}

Write-Host "== Progress Studio WIN-3 installer build =="
Write-Host "Portable payload: $PortableFolder"

if (-not $SkipPortableProbe) {
    Write-Host "Re-validating portable payload before installer packaging..."
    & "$RepoRoot\scripts\validate-windows-portable.ps1" -PortableFolder $PortableFolder
    if (-not $?) {
        throw "WIN-2 portable validation failed."
    }
}

# The .iss intentionally points to dist\ProgressStudio. If the caller supplies
# another known-good payload, stage a clean copy there before compiling.
$ExpectedPortable = Join-Path $RepoRoot "dist\ProgressStudio"
if ([IO.Path]::GetFullPath($PortableFolder) -ne [IO.Path]::GetFullPath($ExpectedPortable)) {
    Write-Host "Staging known-good portable payload into dist\ProgressStudio..."
    Remove-Item $ExpectedPortable -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path (Split-Path -Parent $ExpectedPortable) -Force | Out-Null
    Copy-Item -Path $PortableFolder -Destination $ExpectedPortable -Recurse
}

if (-not $InnoCompiler) {
    $Candidates = @(
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles(x86)\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
    )
    $InnoCompiler = $Candidates | Where-Object { $_ -and (Test-Path $_ -PathType Leaf) } | Select-Object -First 1
}
if (-not $InnoCompiler -or -not (Test-Path $InnoCompiler -PathType Leaf)) {
    throw "Inno Setup 6 compiler (ISCC.exe) was not found. Install Inno Setup 6, then rerun or pass -InnoCompiler."
}

$InstallerOut = Join-Path $RepoRoot "dist\installer"
Remove-Item $InstallerOut -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $InstallerOut -Force | Out-Null

Write-Host "Compiling installer..."
& $InnoCompiler "$RepoRoot\packaging\windows\ProgressStudio.iss"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$Installer = Get-ChildItem $InstallerOut -Filter "ProgressStudio-Setup-*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $Installer) { throw "Installer output was not created in $InstallerOut" }

$Hash = (Get-FileHash $Installer.FullName -Algorithm SHA256).Hash
Write-Host ""
Write-Host "WIN-3 BUILD PASS"
Write-Host "Installer: $($Installer.FullName)"
Write-Host "SHA256:   $Hash"
Write-Host ""
Write-Host "Next: install/uninstall using packaging\windows\WIN3_CHECKLIST.md."
