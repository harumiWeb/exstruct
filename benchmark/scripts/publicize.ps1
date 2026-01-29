Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$benchDir = Split-Path -Parent $scriptDir

$venvPython = Join-Path $benchDir ".venv\Scripts\python"
if (Test-Path $venvPython) {
    & $venvPython (Join-Path $scriptDir "publicize.py")
    exit $LASTEXITCODE
}

python (Join-Path $scriptDir "publicize.py")
exit $LASTEXITCODE
