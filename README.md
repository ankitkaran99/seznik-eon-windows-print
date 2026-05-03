# BT Thermal Printer Toolkit

Direct BLE printing for 57mm thermal printers on Windows.

1. Scan and save the printer config in the project folder.
2. Run a direct print command.

## Files

| File | Purpose |
|------|---------|
| `bt_shared.py` | Shared constants, printer detection, GATT probe, and helpers |
| `bt_scan.py` | Finds nearby BLE printers and saves the best match |
| `bt_print.py` | Direct BLE printing for text, images, PDFs, and test pages |

Keep all three files in the same directory.

## Requirements

Required:

```bash
pip install bleak
```

Optional:

```bash
pip install Pillow
pip install pymupdf
```

Optional external tool:

- Ghostscript: improves PDF/PostScript rendering fallback paths.

## Quick Start

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

## Notes

- Keep the printer powered on and nearby while printing.
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
