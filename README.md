# Seznik EON Printer Toolkit

Direct BLE printing for 57mm thermal printers on Windows, with both a desktop GUI and CLI tools.

Recommended flow:

1. Open the GUI.
2. Click `Scan` to detect and save the printer config.
3. Print text, PDFs, images, or a test page from the same window.

## Files

| File | Purpose |
|------|---------|
| `bt_shared.py` | Shared constants, printer detection, GATT probe, and helpers |
| `bt_scan.py` | Finds nearby BLE printers and saves the best match |
| `bt_print.py` | Direct BLE printing for text, images, PDFs, and test pages |
| `printer_gui.py` | Tkinter desktop GUI for scan and print actions |
| `launch.vbs` | Starts the GUI without showing a console window |
| `build_exe.ps1` | Rebuilds `dist/SeznikEONPrinterToolkit.exe` with PyInstaller |
| `dist/SeznikEONPrinterToolkit.exe` | Standalone Windows GUI build |

Keep all files in the same directory.

## Quick Start

### First Use

For most users, start with the packaged Windows app:

Download the latest build from the repository release page, then run:

```text
dist\SeznikEONPrinterToolkit.exe
```

First-time steps:

1. Turn the printer on and keep it nearby.
2. Open `dist\SeznikEONPrinterToolkit.exe`.
3. Click `Scan` to detect and save the printer config.
4. Choose `Text`, `PDF`, `Image`, or `Test Page`.
5. Click `Start Print`.

### Other Launch Options

If you want to run the Python GUI source instead:

```bash
python printer_gui.py
```

Or on Windows, if Python is installed:

```text
launch.vbs
```

The GUI provides:

- `Scan` to detect and save the printer config
- `Text` mode for direct text printing
- `PDF` mode with file picker
- `Image` mode with file picker
- `Test Page` mode
- A live log panel showing scanner and printer output

PDF note:
PDF content should be sized to 57mm width only for reliable output.

### Build EXE

To rebuild the standalone Windows executable:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

Build output:

```text
dist\SeznikEONPrinterToolkit.exe
```

## Python Requirements

Only needed if you want to run the source files directly instead of using `dist\SeznikEONPrinterToolkit.exe`.

Required:

```bash
pip install bleak Pillow pymupdf
```

Optional external tool:
- Tkinter: required for the Python GUI path
- Ghostscript: improves PDF/PostScript rendering fallback paths

### CLI

### 1. Scan and save printer config

Turn the printer on and make sure it is advertising, then run:

```bash
python bt_scan.py --save
```

This writes the detected printer config to:

```text
bt_printer_config.json
```

The config file is stored in the same directory as the scripts, not in your home folder.

### 2. Print directly

Print a test page:

```bash
python bt_print.py --test-page
```

Print text:

```bash
python bt_print.py --print-text "Hello"
```

Print an image:

```bash
python bt_print.py --print-image photo.jpg
```

Print a PDF:

```bash
python bt_print.py --print-pdf file.pdf
```

Use PDFs designed for 57mm width only.

## `bt_scan.py`

```bash
python bt_scan.py [options]
```

| Flag | Default | Description |
|------|---------|-------------|
| `--scan-time N` | `12` | BLE scan duration in seconds |
| `--all` | off | Show all BLE devices, not just printer candidates |
| `--no-probe` | off | Skip GATT probe |
| `--save` | off | Save the detected printer to `bt_printer_config.json` |

## `bt_print.py`

```bash
python bt_print.py [option]
```

| Flag | Description |
|------|-------------|
| `--test-page` | Print a built-in test page |
| `--print-text TEXT` | Print plain text directly over BLE |
| `--print-image FILE` | Print an image file |
| `--print-pdf FILE` | Print a PDF file |

## Supported Paths

- Cat/iPrint printers using `AE01` are sent through the native bitmap protocol.
- BLE ESC/POS printers are sent ESC/POS payloads directly.
- Images are scaled to 58mm paper width.
- PDFs are rasterized before printing.
- PDFs should be designed at 57mm content width for best results.

## Notes

- Keep the printer powered on and nearby while printing.
- Most users should download the latest packaged build from the repository release page.
- Users who do not want to install Python can run `dist\SeznikEONPrinterToolkit.exe`.
- The GUI runs the same `bt_scan.py` and `bt_print.py` logic as the CLI tools.
- `launch.vbs` uses `pythonw.exe` so the GUI opens without a terminal window.
- If `bt_print.py` says no config was found, run `python bt_scan.py --save` first.
- If PDF printing is poor or unavailable, install `pymupdf` and optionally Ghostscript.
- Browser `Ctrl+P` workflows were intentionally removed from this project.

## Example Workflow

```bash
python bt_scan.py --save
python bt_print.py --print-text "Receipt line 1"
python bt_print.py --print-image logo.png
python bt_print.py --print-pdf invoice.pdf
```
