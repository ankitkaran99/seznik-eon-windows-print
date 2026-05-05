# Seznik EON Printer Toolkit

Windows toolkit for printing to supported 57mm BLE thermal printers over Bluetooth, with an optional Windows printer relay.

## What It Does

- Scans for a compatible BLE printer and saves its config
- Prints text, images, PDFs, and test pages directly over BLE
- Creates a normal Windows printer queue that forwards jobs to the BLE printer through a local relay

## Repository Layout

- `printer_gui.py`: GUI for scan and print operations
- `bt_scan.py`: BLE scanner and printer config saver
- `bt_print.py`: direct BLE printing
- `printer_relay.py`: local TCP relay for Windows spool jobs
- `configure_relay_printer.ps1`: full Windows relay setup and uninstall script
- `launch.vbs`: launches the GUI without a console window
- `drivers\`: bundled printer driver package used by relay setup

## Requirements

- Windows
- Administrator access for relay setup
- Internet access during setup so the toolkit can install system Python and Python packages when they are missing
- A supported BLE receipt printer powered on and available for pairing

## Quick Start

If you want a normal Windows printer that forwards to the BLE printer, run:

```powershell
powershell -ExecutionPolicy Bypass -File configure_relay_printer.ps1
```

That script handles the full setup:

1. Resolves a system Python install, or installs one with `winget` if it is missing
2. Installs required Python packages into that system Python
3. Prompts for BLE printer scan and saves printer config
4. Installs or reuses the printer driver
5. Creates or rebinds the Windows printer queue
6. Registers the relay at logon
7. Creates Desktop and Start Menu shortcuts for the GUI launcher

## GUI Usage

Launch the GUI:

```text
launch.vbs
```

Or run it directly with Python if needed:

```powershell
python printer_gui.py
```

Typical GUI flow:

1. Power on the printer and put it into advertising mode
2. Click `Scan`
3. Save the detected printer config
4. Print text, an image, a PDF, or a test page

## CLI Usage

Scan and save printer config:

```powershell
python bt_scan.py --save
```

Print a test page:

```powershell
python bt_print.py --test-page
```

Print text:

```powershell
python bt_print.py --print-text "Hello"
```

Print an image:

```powershell
python bt_print.py --print-image photo.jpg
```

Print a PDF:

```powershell
python bt_print.py --print-pdf file.pdf
```

Use PDFs designed for 57mm paper width.

## Relay Setup

Use the relay when you want applications to print through a normal Windows printer queue.

Run setup as Administrator:

```powershell
powershell -ExecutionPolicy Bypass -File configure_relay_printer.ps1
```

Default relay target:

```text
127.0.0.1:9100
```

Setup creates:

- A saved BLE printer config
- A local TCP/IP printer port
- A Windows printer queue bound to that port
- A logon startup launcher for `printer_relay.py`
- Desktop and Start Menu shortcuts to `launch.vbs`
- User environment variables used by the launchers

### Common Options

If the driver package installs only driver files and not the printer queue, supply the exact Windows driver name:

```powershell
powershell -ExecutionPolicy Bypass -File configure_relay_printer.ps1 `
  -DriverName "POS-58 Series"
```

If you want to override the Python executable or relay script path:

```powershell
powershell -ExecutionPolicy Bypass -File configure_relay_printer.ps1 `
  -PythonExecutablePath "C:\Users\ankit\AppData\Local\Programs\Python\Python312\python.exe" `
  -RelayScriptPath "printer_relay.py"
```

If you want to use a custom relay port:

```powershell
powershell -ExecutionPolicy Bypass -File configure_relay_printer.ps1 `
  -RelayPort 9200
```

## Relay Uninstall

Remove the relay setup:

```powershell
powershell -ExecutionPolicy Bypass -File configure_relay_printer.ps1 -Uninstall
```

Uninstall removes:

- The relay startup launcher from the current user's Startup folder
- The Desktop and Start Menu GUI shortcuts
- The relay printer queue if it is bound to the relay port
- The relay TCP/IP port if no queue still uses it
- The user environment variables created for the toolkit launchers

Uninstall does not remove:

- The bundled toolkit files
- The installed system Python
- The installed Windows printer driver package

## Manual Relay Start

Run the relay directly:

```powershell
python printer_relay.py
```

Or with a custom port:

```powershell
python printer_relay.py --port 9200
```

## Notes

- The toolkit stores printer config in the user profile so source runs and relay runs share the same saved printer
- `launch.vbs` and the relay startup launcher depend on environment variables written by the setup script
- Setup expects the selected system Python to include `tkinter`
- If you move the toolkit directory after setup, rerun `configure_relay_printer.ps1` so launchers and shortcuts are refreshed
