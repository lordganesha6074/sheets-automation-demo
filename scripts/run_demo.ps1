# If script execution is blocked, run:
# powershell -ExecutionPolicy Bypass -File .\scripts\run_demo.ps1

$ErrorActionPreference = 'Stop'

if (Get-Command py -ErrorAction SilentlyContinue) {
    $PythonCmd = 'py'
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $PythonCmd = 'python'
} else {
    throw 'Python is required but was not found on PATH.'
}

if (-not (Test-Path .venv)) {
    & $PythonCmd -m venv .venv
}

. .\.venv\Scripts\Activate.ps1

& $PythonCmd -m pip install -r requirements.txt
& $PythonCmd .\scripts\generate_orders_export.py
& $PythonCmd .\src\run.py

Write-Host ""
Write-Host "Demo run complete. Outputs are in data/processed/"
Get-ChildItem .\data\processed\
