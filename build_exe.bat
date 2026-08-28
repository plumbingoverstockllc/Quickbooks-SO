@echo off
setlocal

REM Build against Python 3.12 — Python 3.14 bundled via --onefile triggers a
REM "Failed to load python314.dll" LoadLibrary error on some Windows 11 machines.
set "PY=C:\Python312\python.exe"
if not exist "%PY%" set "PY=python"

REM Speed: only touch pip when something is actually missing. Repeat release
REM builds otherwise wasted 30-60s re-checking/re-downloading packages that
REM were already installed. Pass /full as the first arg to force a refresh.
if /I "%~1"=="/full" (
  echo Forcing full dependency reinstall...
  "%PY%" -m pip install --user -r requirements.txt
  "%PY%" -m pip install --user pyinstaller
  "%PY%" -m pip install --user --force-reinstall --no-binary charset-normalizer charset-normalizer
  goto deps_done
)

"%PY%" -c "import pandas, openpyxl, xlrd, win32com, pdfplumber, pdfminer, cryptography, cffi, PyInstaller" 2>nul
if errorlevel 1 (
  echo Installing missing dependencies...
  "%PY%" -m pip install --user -r requirements.txt
  "%PY%" -m pip install --user pyinstaller
) else (
  echo Dependencies present - skipping pip.
)

REM charset-normalizer (pulled in by pdfminer.six) ships a mypyc-compiled
REM build whose shared runtime is a hash-named module at the site-packages
REM ROOT (e.g. 81d243...__mypyc.pyd) that PyInstaller's --collect-all can't
REM associate with the package — so the frozen exe dies with
REM "No module named '...__mypyc'" and pdfplumber can't import. Force the
REM PURE-PYTHON build instead; it has no stray extension module to miss.
REM Only reinstall when a compiled .pyd is actually present (otherwise it's
REM already pure-python and we skip the slow sdist rebuild).
"%PY%" -c "import charset_normalizer,os,pathlib,sys; d=pathlib.Path(charset_normalizer.__file__).parent; sys.exit(1 if any(f.endswith('.pyd') for f in os.listdir(d)) else 0)" 2>nul
if errorlevel 1 (
  echo Reinstalling pure-python charset-normalizer...
  "%PY%" -m pip install --user --force-reinstall --no-binary charset-normalizer charset-normalizer
)
:deps_done

REM --collect-all numpy/pandas avoids "No module named 'numpy._core._exceptions'"
REM at runtime — numpy 2.x reorganized its internals and the default PyInstaller
REM hook misses several C-extension submodules.
REM pdfplumber (+ its pdfminer.six backend) reads the Deluxe Vanity showroom
REM PO PDFs. pdfminer ships CMap data files the default hook misses, AND it
REM imports cryptography at load time (which carries a compiled _cffi_backend
REM and a Rust binding). plus charset_normalizer (pure-python, see above).
REM All must be collected or `import pdfplumber` fails at runtime.
set "COLLECT=--collect-all numpy --collect-all pandas --collect-all openpyxl --collect-all xlrd --collect-all pdfplumber --collect-all pdfminer --collect-all cryptography --collect-all cffi --collect-all charset_normalizer"

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
"%PY%" -m PyInstaller --noconfirm --onefile --windowed --name "DMQuotes" --icon "DMQuotes.ico" --splash "splash.png" %COLLECT% --add-data "Logo_New.png;." --add-data "DMQuotes.ico;." --add-data "vendor_pricing.json;." app.py

if exist "C:\Users\QB-PC\AppData\Local\Programs\Inno Setup 6\ISCC.exe" (
  "C:\Users\QB-PC\AppData\Local\Programs\Inno Setup 6\ISCC.exe" "installer.iss"
) else (
  echo Inno Setup not found. Skipping Setup.exe build.
)

echo.
echo Build complete. Check dist folder for portable EXE and Setup.exe.
if defined GITHUB_ACTIONS goto :eof
pause
