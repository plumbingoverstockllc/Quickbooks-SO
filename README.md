# QuickBooks Sales Order Converter (Desktop)

Windows desktop app to:

1. Load your source order file.
2. Ask for customer/order fields and pricing rules.
3. Support pricing rules by `Per Brand` or `Per Item (SKU)`.
4. Optional checkbox to use typed `Actual Cost/Rate` instead of MSRP multiplier.
5. Convert to SaaSant Sales Order template format.
6. Preview and validate rows before upload.
7. One-click `Export for SaaSant` to Downloads with a ready-to-upload file name.
8. Export custom-path `.xlsx` or upload directly to QuickBooks Desktop Enterprise.

## Files

- `app.py` - Desktop UI and workflow
- `transformer.py` - Source-to-template mapping logic
- `quickbooks_client.py` - QuickBooks Desktop SDK connection, next SO number lookup, and upload
- `requirements.txt` - Python dependencies
- `build_exe.bat` - One-click EXE build

## Run locally

```powershell
cd "C:\Users\QB-PC\Downloads\qb_so_desktop_app"
python -m pip install -r requirements.txt
python app.py
```

## Build executables

```powershell
cd "C:\Users\QB-PC\Downloads\qb_so_desktop_app"
build_exe.bat
```

The packaged app is created under `dist\QB Sales Order Converter\`.

## Build Windows installer (Setup.exe)

```powershell
cd "C:\Users\QB-PC\Downloads\qb_so_desktop_app"
"C:\Users\QB-PC\AppData\Local\Programs\Inno Setup 6\ISCC.exe" "installer.iss"
```

Installer output:

- `dist\QB-Sales-Order-Converter-Setup.exe`

## QuickBooks requirements

- QuickBooks Desktop Enterprise installed on same machine.
- Company file open.
- First run requires admin approval when QuickBooks prompts to authorize this app.
- Item names in QuickBooks should match `SKU` values in your source file.
- Customer name and Terms/Ship Method should exist in QuickBooks.

## Notes

- `Fetch Next from QuickBooks` scans existing numeric sales order numbers and suggests max + 1.
- Upload sends one Sales Order with all preview lines.
- If upload fails, app shows exact QuickBooks status message.
