@echo off
rem Builds PermitReader.exe (single file, no install needed on the target PC).
setlocal
cd /d "%~dp0"

if not exist .venv\Scripts\python.exe (
    py -3 -m venv .venv || goto :error
)
.venv\Scripts\python -m pip install --upgrade pip || goto :error
.venv\Scripts\python -m pip install -r requirements.txt || goto :error

.venv\Scripts\pyinstaller --noconfirm --clean --workpath build --distpath dist PermitReader.spec || goto :error

copy /y dist\PermitReader.exe PermitReader.exe >nul || goto :error
echo.
echo Built: %CD%\PermitReader.exe
exit /b 0

:error
echo.
echo Build failed.
exit /b 1
