param(
    [string]$Case = "all",
    [string]$Method = "all",
    [string]$Model = "gpt-4o",
    [double]$Temperature = 0.0,
    [switch]$SkipAsk
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Info {
    param([string]$Message)
    Write-Host "[reproduce] $Message"
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$benchDir = Split-Path -Parent $scriptDir
$repoDir = Split-Path -Parent $benchDir

Push-Location $benchDir

try {
    if (-not (Test-Path ".env")) {
        Write-Info "Copying .env.example -> .env (remember to set OPENAI_API_KEY)."
        Copy-Item ".env.example" ".env"
    }

    if (-not (Test-Path ".venv")) {
        Write-Info "Creating virtual environment."
        python -m venv .venv
    }

    $python = ".venv\\Scripts\\python"
    if (-not (Test-Path $python)) {
        throw "Python venv not found at $python"
    }

    Write-Info "Installing dependencies."
    & $python -m pip install -e $repoDir
    & $python -m pip install -e .

    Write-Info "Extracting contexts."
    & $python -m bench.cli extract --case $Case --method $Method

    if (-not $SkipAsk) {
        Write-Info "Running LLM inference."
        & $python -m bench.cli ask --case $Case --method $Method --model $Model --temperature $Temperature
    } else {
        Write-Info "Skipping LLM inference."
    }

    Write-Info "Evaluating results."
    & $python -m bench.cli eval --case $Case --method $Method

    Write-Info "Generating reports."
    & $python -m bench.cli report
    & $python -m bench.cli report-public
} finally {
    Pop-Location
}
