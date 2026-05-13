@echo off
setlocal

REM Build against Python 3.12 — Python 3.14 bundled via --onefile triggers a
REM "Failed to load python314.dll" LoadLibrary error on some Windows 11 machines.
set "PY=C:\Python312\python.exe"
if not exist "%PY%" set "PY=python"

"%PY%" -m pip install --user -r requirements.txt
"%PY%" -m pip install --user pyinstaller
"%PY%" -m PyInstaller --noconfirm --windowed --name "QB Sales Order Converter" app.py
"%PY%" -m PyInstaller --noconfirm --onefile --windowed --name "QB Sales Order Converter" app.py

if exist "C:\Users\QB-PC\AppData\Local\Programs\Inno Setup 6\ISCC.exe" (
  "C:\Users\QB-PC\AppData\Local\Programs\Inno Setup 6\ISCC.exe" "installer.iss"
) else (
  echo Inno Setup not found. Skipping Setup.exe build.
)

echo.
echo Build complete. Check dist folder for portable EXE and Setup.exe.
pause
