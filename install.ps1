$ErrorActionPreference = 'Stop'
$Python = if ($env:PYTHON) { $env:PYTHON } else { 'python' }
& $Python (Join-Path $PSScriptRoot 'install.py') @args
exit $LASTEXITCODE
