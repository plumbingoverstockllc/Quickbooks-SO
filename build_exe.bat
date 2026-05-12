@echo off
setlocal

python -m pip install -r requirements.txt
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --windowed --name "QB Sales Order Converter" app.py
python -m PyInstaller --noconfirm --onefile --windowed --name "QB Sales Order Converter Portable v3" app.py

if exist "C:\Users\QB-PC\AppData\Local\Programs\Inno Setup 6\ISCC.exe" (
  "C:\Users\QB-PC\AppData\Local\Programs\Inno Setup 6\ISCC.exe" "installer.iss"
) else (
  echo Inno Setup not found. Skipping Setup.exe build.
)

echo.
echo Build complete. Check dist folder for portable EXE and Setup.exe.
pause
