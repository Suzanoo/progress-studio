$ErrorActionPreference = "Stop"

$roots = @(".pytest_cache", "build", "dist")
foreach ($path in $roots) {
    if (Test-Path $path) { Remove-Item -Recurse -Force $path }
}

Get-ChildItem -Recurse -Directory -Force -Filter "__pycache__" | Remove-Item -Recurse -Force
Get-ChildItem -Recurse -File -Force -Include "*.pyc","*.pyo" | Remove-Item -Force
Get-ChildItem -Directory -Force -Filter "*.egg-info" | Remove-Item -Recurse -Force

Write-Host "Repository caches/build artifacts removed."
