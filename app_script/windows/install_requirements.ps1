$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Set-Location $ProjectRoot

$venvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$reqFile = Join-Path $ProjectRoot "requirements.txt"

if (-not (Test-Path $reqFile)) {
    Write-Error "requirements.txt not found at $reqFile"
    exit 1
}

if (-not (Test-Path $venvPython)) {
    Write-Host "Creating virtual environment (.venv)..." -ForegroundColor Cyan
    python -m venv .venv
}

Write-Host "Upgrading pip and installing dependencies from requirements.txt..." -ForegroundColor Cyan
& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install -r $reqFile

$envFile = Join-Path $ProjectRoot ".env"
$envExample = Join-Path $ProjectRoot ".env.example"
if (-not (Test-Path $envFile) -and (Test-Path $envExample)) {
    Copy-Item $envExample $envFile
    Write-Host "Created .env from .env.example. Remember to set GEMINI_API_KEY in .env!" -ForegroundColor Yellow
}

Write-Host "Dependencies installed successfully." -ForegroundColor Green
