$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

function Find-Python {
    $candidates = @(
        "C:\Users\ankit\AppData\Local\Programs\Python\Python312\python.exe",
        "C:\Users\ankit\AppData\Local\Programs\Python\Python313\python.exe",
        "C:\Users\ankit\AppData\Local\Programs\Python\Python311\python.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Path -and (Test-Path $cmd.Path) -and ((Get-Item $cmd.Path).Length -gt 0)) {
        return $cmd.Path
    }

    throw "No usable Python installation was found. Install Python 3.11+ and rerun this script."
}

function Ensure-Package {
    param(
        [string]$PythonExe,
        [string]$ImportName,
        [string]$PackageName
    )

    & $PythonExe -c "import $ImportName" 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Installing $PackageName..."
        & $PythonExe -m pip install $PackageName
    }
}

$PythonExe = Find-Python
$PythonHome = Split-Path -Parent $PythonExe
$TclDir = Join-Path $PythonHome "tcl\tcl8.6"
$TkDir = Join-Path $PythonHome "tcl\tk8.6"
$PyInstallerExe = Join-Path $PythonHome "Scripts\pyinstaller.exe"

if (-not (Test-Path $TclDir)) {
    throw "Tcl runtime not found at $TclDir"
}

if (-not (Test-Path $TkDir)) {
    throw "Tk runtime not found at $TkDir"
}

Write-Host "Using Python: $PythonExe"
& $PythonExe --version

Write-Host "Ensuring build dependencies..."
Ensure-Package -PythonExe $PythonExe -ImportName "PyInstaller" -PackageName "pyinstaller"
Ensure-Package -PythonExe $PythonExe -ImportName "bleak" -PackageName "bleak"
Ensure-Package -PythonExe $PythonExe -ImportName "PIL" -PackageName "Pillow"
Ensure-Package -PythonExe $PythonExe -ImportName "fitz" -PackageName "pymupdf"

if (-not (Test-Path $PyInstallerExe)) {
    throw "PyInstaller executable not found at $PyInstallerExe"
}

Write-Host "Building dist\SeznikEONPrinterToolkit.exe ..."
& $PyInstallerExe `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name SeznikEONPrinterToolkit `
    --hidden-import tkinter `
    --hidden-import _tkinter `
    --add-data "${TclDir};tcl\tcl8.6" `
    --add-data "${TkDir};tcl\tk8.6" `
    --runtime-hook pyi_rth_tkfix.py `
    printer_gui.py

Write-Host ""
Write-Host "Build complete:"
Write-Host "  $ProjectRoot\dist\SeznikEONPrinterToolkit.exe"
