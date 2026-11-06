$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
Push-Location -LiteralPath $projectRoot
try {
    $python = Join-Path $projectRoot '.venv/Scripts/python.exe'
    if (-not (Test-Path -LiteralPath $python)) {
        & py -3.13 -m venv .venv
        if ($LASTEXITCODE -ne 0) { throw 'Python 3.13 x64 is required.' }
    }
    & $python -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) { throw 'Could not update pip.' }
    & $python -m pip install -r requirements-dev.txt
    if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed.' }
    & $python scripts/check_project.py
    if ($LASTEXITCODE -ne 0) { throw 'Project checks failed.' }
    & $python -m PyInstaller --noconfirm WordToPDF.spec
    if ($LASTEXITCODE -ne 0) { throw 'EXE build failed.' }
    Write-Host 'Build complete. See dist/.'
} finally {
    Pop-Location
}
