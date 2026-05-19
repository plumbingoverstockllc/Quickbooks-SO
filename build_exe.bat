@echo off
setlocal

REM Build against Python 3.12 — Python 3.14 bundled via --onefile triggers a
REM "Failed to load python314.dll" LoadLibrary error on some Windows 11 machines.
set "PY=C:\Python312\python.exe"
if not exist "%PY%" set "PY=python"

"%PY%" -m pip install --user -r requirements.txt
"%PY%" -m pip install --user pyinstaller

REM --collect-all numpy/pandas avoids "No module named 'numpy._core._exceptions'"
REM at runtime — numpy 2.x reorganized its internals and the default PyInstaller
REM hook misses several C-extension submodules.
set "COLLECT=--collect-all numpy --collect-all pandas --collect-all openpyxl --collect-all xlrd"

REM We only ship the --onefile build (installer.iss sources dist\*.exe).
REM The earlier --onedir invocation was overwritten by --onefile anyway, so
REM dropping it cuts the build time roughly in half.
REM --add-data bundles the brand PNG next to the bootstrap in the onefile
REM build. Syntax is "<src>;<dest>" on Windows (semicolon, not colon).
REM --splash shows the logo immediately when the .exe is launched, while
REM PyInstaller's bootloader unpacks the bundle to %TEMP%. Without this
REM there's a 3-5 second window where the user sees nothing and assumes
REM the launch failed. app.py calls pyi_splash.close() once the main
REM window is up.
"%PY%" -m PyInstaller --noconfirm --onefile --windowed --name "DMQuotes" --icon "DMQuotes.ico" --splash "splash.png" %COLLECT% --add-data "Logo_New.png;." --add-data "DMQuotes.ico;." app.py

if exist "C:\Users\QB-PC\AppData\Local\Programs\Inno Setup 6\ISCC.exe" (
  "C:\Users\QB-PC\AppData\Local\Programs\Inno Setup 6\ISCC.exe" "installer.iss"
) else (
  echo Inno Setup not found. Skipping Setup.exe build.
)

echo.
echo Build complete. Check dist folder for portable EXE and Setup.exe.
pause
