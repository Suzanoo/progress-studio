param(
    [string]$Python = "python"
)

& $Python "$PSScriptRoot\check-package.py"
exit $LASTEXITCODE
