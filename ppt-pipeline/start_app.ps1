$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$workspaceRoot = Split-Path -Parent $projectRoot
$venvPath = Join-Path $workspaceRoot '.venv312'
$pythonExe = Join-Path $venvPath 'Scripts\python.exe'
$appPath = Join-Path $projectRoot 'app.py'

# Check if venv exists
if (-not (Test-Path $pythonExe)) {
    Write-Host "Python 3.12 venv not found at $venvPath" -ForegroundColor Red
    Write-Host ""
    Write-Host "Creating venv..." -ForegroundColor Yellow
    
    # Try to find python 3.12
    $pythonCmd = $null
    
    # Method 1: Try py launcher
    $pyOutput = py -0 2>$null
    if ($pyOutput -match '3\.12') {
        $pythonCmd = 'py -3.12.0 -m venv'
    }
    # Method 2: Try python directly (common on Windows with Python installed from python.org)
    elseif ((Get-Command python -ErrorAction SilentlyContinue)) {
        $version = python -V 2>&1
        if ($version -match '3\.12') {
            $pythonCmd = 'python -m venv'
        }
    }
    
    if ($pythonCmd) {
        Write-Host "Running: $pythonCmd $venvPath" -ForegroundColor Cyan
        Invoke-Expression "$pythonCmd $venvPath"
    } else {
        throw "Python 3.12 not found. Install from https://www.python.org/downloads/ (Python 3.12)"
    }
}

# Verify venv created
if (-not (Test-Path $pythonExe)) {
    throw "Failed to create Python 3.12 venv"
}

Write-Host "Using Python: $((& $pythonExe -V 2>&1))" -ForegroundColor Green
Write-Host "Starting app from: $appPath" -ForegroundColor Green
& $pythonExe $appPath