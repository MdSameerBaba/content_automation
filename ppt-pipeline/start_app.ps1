$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$workspaceRoot = Split-Path -Parent $projectRoot
$pythonExe = Join-Path $workspaceRoot '.venv312\Scripts\python.exe'
$appPath = Join-Path $projectRoot 'app.py'

if (-not (Test-Path $pythonExe)) {
    throw "Python 3.12 environment not found at $pythonExe. Create it with: py -3.12 -m venv .venv312"
}

& $pythonExe $appPath
