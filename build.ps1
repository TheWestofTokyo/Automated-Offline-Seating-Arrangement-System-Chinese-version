$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    python -m pip install -r requirements.txt pyinstaller==6.22.3
    if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed.' }
    python -m unittest discover -p 'test_*.py' -v
    if ($LASTEXITCODE -ne 0) { throw 'Regression tests failed.' }
    python -m PyInstaller --noconfirm --clean --onefile --windowed --collect-all customtkinter --add-data 'LICENSES;LICENSES' --name '自动化排座管理系统' app.py
    if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
} finally {
    Pop-Location
}
