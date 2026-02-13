$ErrorActionPreference = "Stop"

function Get-PythonCommand {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        return "py -3"
    }
    if (Get-Command python -ErrorAction SilentlyContinue) {
        return "python"
    }
    throw "Python is not found. Install Python 3.10+ and run setup again."
}

$pythonCmd = Get-PythonCommand
$venvPath = ".venv"

Write-Host "Python command: $pythonCmd"

if (-not (Test-Path $venvPath)) {
    Write-Host "Creating virtual environment .venv ..."
    Invoke-Expression "$pythonCmd -m venv $venvPath"
} else {
    Write-Host ".venv already exists, skipping creation."
}

$venvPython = Join-Path $venvPath "Scripts\python.exe"
if (-not (Test-Path $venvPython)) {
    throw "Cannot find $venvPython"
}

Write-Host "Upgrading pip ..."
& $venvPython -m pip install --upgrade pip

Write-Host "Installing dependencies from requirements.txt ..."
& $venvPython -m pip install -r requirements.txt

Write-Host ""
Write-Host "Done."
Write-Host "Run comparison:"
Write-Host ".\.venv\Scripts\python.exe .\compare_payments.py"
