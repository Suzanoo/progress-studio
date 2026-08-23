param(
    [string]$Python = "python"
)

$ErrorActionPreference = "Stop"
& $Python scripts/check-preproduction.py
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
