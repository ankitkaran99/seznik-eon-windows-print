[CmdletBinding()]
param(
    [string]$PrinterName = "POS58 Printer",
    [string]$RelayHost = "127.0.0.1",
    [uint32]$RelayPort = 9100,
    [string]$PortName = "",
    [string]$DriverPackagePath = "",
    [string]$DriversDirectory = "",
    [string]$DriverName = "",
    [string]$DriverInstallerArgs = "",
    [string]$PythonExecutablePath = "",
    [string]$RelayScriptPath = "",
    [string]$StartupLauncherPath = "",
    [string]$GuiLauncherPath = "",
    [string]$DesktopShortcutPath = "",
    [string]$StartMenuShortcutPath = "",
    [switch]$SkipDriverInstall,
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"
$SystemPythonPackageId = "Python.Python.3.12"

function Write-Info([string]$Message) {
    Write-Host "[INFO] $Message"
}

function Write-Ok([string]$Message) {
    Write-Host "[ OK ] $Message" -ForegroundColor Green
}

function Write-Fail([string]$Message) {
    Write-Host "[FAIL] $Message" -ForegroundColor Red
}

function Test-HasText {
    param([string]$Value)

    return -not [string]::IsNullOrWhiteSpace($Value)
}

function Test-PathIfSet {
    param(
        [string]$LiteralPath,
        [Microsoft.PowerShell.Commands.TestPathType]$PathType = [Microsoft.PowerShell.Commands.TestPathType]::Any
    )

    if (-not (Test-HasText -Value $LiteralPath)) {
        return $false
    }

    if ($PathType -eq [Microsoft.PowerShell.Commands.TestPathType]::Any) {
        return Test-Path -LiteralPath $LiteralPath
    }

    return Test-Path -LiteralPath $LiteralPath -PathType $PathType
}

function Test-PythonModuleInstalled {
    param(
        [string]$PythonExe,
        [string]$ImportName
    )

    $quotedImportName = $ImportName.Replace("'", "''")
    & $PythonExe -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('$quotedImportName') else 1)" *> $null
    return ($LASTEXITCODE -eq 0)
}

function Test-PythonExecutable {
    param([string]$PythonExe)

    & $PythonExe -c "import sys; sys.exit(0)" *> $null
    return ($LASTEXITCODE -eq 0)
}

function Get-CommandPathIfPresent {
    param([string]$CommandName)

    $command = Get-Command $CommandName -ErrorAction SilentlyContinue
    if (-not $command) {
        return $null
    }

    return $command.Source
}

function Resolve-PythonFromLauncher {
    param([string]$LauncherCommand)

    $launcherPath = Get-CommandPathIfPresent -CommandName $LauncherCommand
    if (-not (Test-HasText -Value $launcherPath)) {
        return $null
    }

    $pythonExe = & $launcherPath -3 -c "import sys; print(sys.executable)" 2>$null
    if ($LASTEXITCODE -ne 0) {
        return $null
    }

    $candidate = ($pythonExe | Select-Object -First 1).Trim()
    if (-not (Test-HasText -Value $candidate)) {
        return $null
    }

    $resolvedCandidate = Resolve-Path -LiteralPath $candidate -ErrorAction SilentlyContinue
    if (-not $resolvedCandidate) {
        return $null
    }

    return $resolvedCandidate.Path
}

function Get-RegisteredPythonInstallPaths {
    $roots = @(
        "HKLM:\SOFTWARE\Python\PythonCore",
        "HKCU:\SOFTWARE\Python\PythonCore",
        "HKLM:\SOFTWARE\WOW6432Node\Python\PythonCore"
    )

    $installPaths = @()
    foreach ($root in $roots) {
        if (-not (Test-Path -LiteralPath $root)) {
            continue
        }

        foreach ($versionKey in Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue) {
            $installPathKey = Join-Path $versionKey.PSPath "InstallPath"
            $installPath = (Get-ItemProperty -LiteralPath $installPathKey -ErrorAction SilentlyContinue).'(default)'
            if (Test-HasText -Value $installPath) {
                $installPaths += (Join-Path $installPath "python.exe")
            }
        }
    }

    return @($installPaths | Select-Object -Unique)
}

function Resolve-SystemPythonExecutable {
    param([string]$PreferredPythonExe)

    $candidates = @()
    if (Test-HasText -Value $PreferredPythonExe) {
        $candidates += $PreferredPythonExe
    }

    $pythonCommandPath = Get-CommandPathIfPresent -CommandName "python"
    if (Test-HasText -Value $pythonCommandPath) {
        $candidates += $pythonCommandPath
    }

    $launcherPython = Resolve-PythonFromLauncher -LauncherCommand "py"
    if (Test-HasText -Value $launcherPython) {
        $candidates += $launcherPython
    }

    $candidates += Get-RegisteredPythonInstallPaths

    foreach ($candidate in $candidates | Where-Object { Test-HasText -Value $_ } | Select-Object -Unique) {
        $resolvedCandidate = Resolve-Path -LiteralPath $candidate -ErrorAction SilentlyContinue
        if (-not $resolvedCandidate) {
            continue
        }

        if ((Test-PythonExecutable -PythonExe $resolvedCandidate.Path) -and
            (Test-PythonModuleInstalled -PythonExe $resolvedCandidate.Path -ImportName "tkinter")) {
            return $resolvedCandidate.Path
        }
    }

    return $null
}

function Install-SystemPython {
    $wingetPath = Get-CommandPathIfPresent -CommandName "winget"
    if (-not (Test-HasText -Value $wingetPath)) {
        throw "Python was not found and winget is not available to install it automatically."
    }

    Write-Info "Installing system Python with winget package '$SystemPythonPackageId'..."
    $wingetArgs = @(
        "install",
        "--exact",
        "--id", $SystemPythonPackageId,
        "--scope", "machine",
        "--silent",
        "--accept-source-agreements",
        "--accept-package-agreements",
        "--override", "InstallAllUsers=1 PrependPath=1 Include_launcher=1"
    )

    $process = Start-Process -FilePath $wingetPath -ArgumentList $wingetArgs -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        throw "winget install for '$SystemPythonPackageId' failed with exit code $($process.ExitCode)."
    }

    Write-Ok "System Python install completed."
}

function Ensure-PythonPackage {
    param(
        [string]$PythonExe,
        [string]$ImportName,
        [string]$PackageName
    )

    if (Test-PythonModuleInstalled -PythonExe $PythonExe -ImportName $ImportName) {
        Write-Ok "Python package '$PackageName' is already installed."
        return
    }

    Write-Info "Installing Python package '$PackageName'..."
    & $PythonExe -m pip install --disable-pip-version-check $PackageName | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "pip install for '$PackageName' failed with exit code $LASTEXITCODE."
    }

    Write-Ok "Installed Python package '$PackageName'."
}

function Ensure-SystemPython {
    param([string]$PreferredPythonExe)

    $pythonExe = Resolve-SystemPythonExecutable -PreferredPythonExe $PreferredPythonExe
    if (-not (Test-HasText -Value $pythonExe)) {
        Install-SystemPython
        $pythonExe = Resolve-SystemPythonExecutable -PreferredPythonExe $PreferredPythonExe
    }

    if (-not (Test-HasText -Value $pythonExe)) {
        throw "Python installation completed, but no system python.exe could be resolved."
    }

    Write-Info "Using system Python '$pythonExe'."
    & $pythonExe -m ensurepip --upgrade *> $null
    Ensure-PythonPackage -PythonExe $pythonExe -ImportName "bleak" -PackageName "bleak"
    Ensure-PythonPackage -PythonExe $pythonExe -ImportName "PIL" -PackageName "Pillow"
    Ensure-PythonPackage -PythonExe $pythonExe -ImportName "fitz" -PackageName "pymupdf"
    if (-not (Test-PythonModuleInstalled -PythonExe $pythonExe -ImportName "tkinter")) {
        throw "System Python '$pythonExe' does not include tkinter."
    }

    return $pythonExe
}

function Invoke-PrinterScanSetup {
    param([string]$PythonExe)

    Write-Host ""
    Write-Host "Before relay setup continues:" -ForegroundColor Cyan
    Write-Host "  1. Power on the printer."
    Write-Host "  2. Put it into Bluetooth advertising or pairing mode."
    Write-Host "  3. Turn On Bluetooth Adapter of this computer."
    Write-Host "  4. Keep printer close to this computer."
    Write-Host ""
    Read-Host "Press Enter when the printer is ready to scan"

    Write-Info "Running Bluetooth scan and saving printer config..."
    Push-Location $PSScriptRoot
    try {
        & $PythonExe ".\bt_scan.py" --save
        if ($LASTEXITCODE -ne 0) {
            throw "bt_scan.py exited with code $LASTEXITCODE."
        }
    }
    finally {
        Pop-Location
    }

    Write-Ok "Printer scan completed and config was saved."
}

function Confirm-DriverInstallationStep {
    Write-Host ""
    Write-Host "Next step: driver installation" -ForegroundColor Cyan
    Write-Host "  1. The POS58 driver installer will open next."
    Write-Host "  2. In the installer, choose 'Parallel port'."
    Write-Host "  3. Then choose 'LPT1' when it asks for the port."
    Write-Host "  4. Finish the driver installation and return here."
    Write-Host ""
    Read-Host "Press Enter to continue"
}

function New-HiddenRelayLauncher {
    param(
        [string]$LauncherPath,
        [string]$PythonExePath,
        [string]$RelayScriptPath,
        [string]$RelayHost,
        [uint32]$RelayPort
    )

    $launcherDir = Split-Path -Parent $LauncherPath
    if ((Test-HasText -Value $launcherDir) -and (-not (Test-Path -LiteralPath $launcherDir))) {
        New-Item -ItemType Directory -Path $launcherDir -Force | Out-Null
    }

    $resolvedPython = (Resolve-Path -LiteralPath $PythonExePath -ErrorAction Stop).Path.Replace('"', '""')
    $resolvedRelayScript = (Resolve-Path -LiteralPath $RelayScriptPath -ErrorAction Stop).Path.Replace('"', '""')
    $resolvedWorkingDirectory = (Split-Path -Parent $resolvedRelayScript).Replace('"', '""')

    $content = @"
Set oShell = CreateObject("WScript.Shell")
pythonExe = "$resolvedPython"
relayScript = "$resolvedRelayScript"
cmd = Chr(34) & pythonExe & Chr(34) & " " & Chr(34) & relayScript & Chr(34) & " --host $($RelayHost.Replace('"', '""')) --port $RelayPort"
oShell.CurrentDirectory = "$resolvedWorkingDirectory"
oShell.Run cmd, 0, False
"@

    Set-Content -LiteralPath $LauncherPath -Value $content -Encoding ASCII
    return $LauncherPath
}

function Ensure-RelayStartup {
    param(
        [string]$PythonExePath,
        [string]$RelayScriptPath,
        [string]$LauncherPath,
        [string]$RelayHost,
        [uint32]$RelayPort
    )

    $resolvedPython = Resolve-Path -LiteralPath $PythonExePath -ErrorAction SilentlyContinue
    if (-not $resolvedPython) {
        Write-Fail "Python executable '$PythonExePath' was not found."
        exit 1
    }

    $resolvedRelayScript = Resolve-Path -LiteralPath $RelayScriptPath -ErrorAction SilentlyContinue
    if (-not $resolvedRelayScript) {
        Write-Fail "Relay script '$RelayScriptPath' was not found."
        exit 1
    }

    $startupFolder = [Environment]::GetFolderPath("Startup")
    $startupLauncher = Join-Path $startupFolder "SeznikEONRelay.vbs"
    $sourceLauncher = New-HiddenRelayLauncher `
        -LauncherPath $LauncherPath `
        -PythonExePath $resolvedPython.Path `
        -RelayScriptPath $resolvedRelayScript.Path `
        -RelayHost $RelayHost `
        -RelayPort $RelayPort
    Copy-Item -LiteralPath $sourceLauncher -Destination $startupLauncher -Force
    Write-Ok "Relay startup launcher installed to '$startupLauncher'."
    return $startupLauncher
}

function Start-RelayLauncher {
    param([string]$LauncherPath)

    $resolvedLauncher = Resolve-Path -LiteralPath $LauncherPath -ErrorAction SilentlyContinue
    if (-not $resolvedLauncher) {
        Write-Fail "Relay launcher '$LauncherPath' was not found."
        exit 1
    }

    $wscriptPath = Join-Path $env:SystemRoot "System32\wscript.exe"
    $process = Start-Process -FilePath $wscriptPath -ArgumentList "`"$($resolvedLauncher.Path)`"" -WindowStyle Hidden -PassThru
    Start-Sleep -Milliseconds 500
    if ($process.HasExited -and $process.ExitCode -ne 0) {
        throw "Relay launcher failed with exit code $($process.ExitCode)."
    }

    Write-Ok "Relay launcher started."
}

function Remove-RelayStartup {
    param([string]$LauncherPath)

    $startupFolder = [Environment]::GetFolderPath("Startup")
    $startupLauncher = Join-Path $startupFolder "SeznikEONRelay.vbs"
    $removedAny = $false

    foreach ($path in @($LauncherPath, $startupLauncher) | Where-Object { $_ } | Select-Object -Unique) {
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            Remove-Item -LiteralPath $path -Force
            Write-Ok "Removed relay launcher '$path'."
            $removedAny = $true
        }
    }

    if (-not $removedAny) {
        Write-Info "No relay startup launcher was found."
    }
}

function Ensure-GuiShortcut {
    param(
        [string]$LauncherPath,
        [string]$ShortcutPath
    )

    $resolvedLauncher = Resolve-Path -LiteralPath $LauncherPath -ErrorAction SilentlyContinue
    if (-not $resolvedLauncher) {
        Write-Fail "GUI launcher '$LauncherPath' was not found."
        exit 1
    }

    $shortcutDir = Split-Path -Parent $ShortcutPath
    if ((Test-HasText -Value $shortcutDir) -and (-not (Test-Path -LiteralPath $shortcutDir))) {
        New-Item -ItemType Directory -Path $shortcutDir -Force | Out-Null
    }

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $resolvedLauncher.Path
    $shortcut.WorkingDirectory = Split-Path -Parent $resolvedLauncher.Path
    $shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,220"
    $shortcut.Save()

    Write-Ok "GUI shortcut installed to '$ShortcutPath'."
    return $ShortcutPath
}

function Remove-GuiShortcut {
    param([string]$ShortcutPath)

    if (Test-Path -LiteralPath $ShortcutPath -PathType Leaf) {
        Remove-Item -LiteralPath $ShortcutPath -Force
        Write-Ok "Removed GUI shortcut '$ShortcutPath'."
        return $true
    }

    Write-Info "No GUI shortcut was found at '$ShortcutPath'."
    return $false
}

function Ensure-GuiShortcuts {
    param(
        [string]$LauncherPath,
        [string[]]$ShortcutPaths
    )

    $createdPaths = @()
    foreach ($shortcutPath in $ShortcutPaths | Where-Object { $_ } | Select-Object -Unique) {
        $createdPaths += Ensure-GuiShortcut -LauncherPath $LauncherPath -ShortcutPath $shortcutPath
    }

    return $createdPaths
}

function Remove-GuiShortcuts {
    param([string[]]$ShortcutPaths)

    $removedAny = $false
    foreach ($shortcutPath in $ShortcutPaths | Where-Object { $_ } | Select-Object -Unique) {
        if (Remove-GuiShortcut -ShortcutPath $shortcutPath) {
            $removedAny = $true
        }
    }

    return $removedAny
}

function Set-UserEnvironmentVariable {
    param(
        [string]$Name,
        [string]$Value
    )

    [Environment]::SetEnvironmentVariable($Name, $Value, "User")
}

function Remove-UserEnvironmentVariable {
    param([string]$Name)

    [Environment]::SetEnvironmentVariable($Name, $null, "User")
}

function Ensure-ToolkitEnvironment {
    param(
        [string]$PythonExePath,
        [string]$ToolkitDir
    )

    $pythonExe = (Resolve-Path -LiteralPath $PythonExePath -ErrorAction Stop).Path
    $toolkitRoot = (Resolve-Path -LiteralPath $ToolkitDir -ErrorAction Stop).Path
    Set-UserEnvironmentVariable -Name "SEZNIK_EON_TOOLKIT_DIR" -Value $toolkitRoot
    Set-UserEnvironmentVariable -Name "SEZNIK_EON_PYTHON" -Value $pythonExe
    Write-Ok "Saved user environment for toolkit launchers."
}

function Remove-ToolkitEnvironment {
    Remove-UserEnvironmentVariable -Name "SEZNIK_EON_TOOLKIT_DIR"
    Remove-UserEnvironmentVariable -Name "SEZNIK_EON_PYTHON"
    Remove-UserEnvironmentVariable -Name "SEZNIK_EON_PYTHONW"
    Write-Ok "Removed user environment for toolkit launchers."
}

function Ensure-Admin {
    $windowsIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $windowsPrincipal = [Security.Principal.WindowsPrincipal]::new($windowsIdentity)
    $isAdmin = $windowsPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdmin) {
        Write-Fail "This script must be run as Administrator."
        exit 1
    }
}

function New-DriverWorkingDirectory {
    $root = Join-Path $env:TEMP "SeznikEONPrinterToolkit"
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $path = Join-Path $root "driver-$stamp"
    New-Item -ItemType Directory -Path $path -Force | Out-Null
    return $path
}

function Get-InstalledPrinterDriverNames {
    $drivers = Get-PrinterDriver -ErrorAction SilentlyContinue
    if (-not $drivers) {
        return @()
    }

    return @($drivers | Select-Object -ExpandProperty Name)
}

function Resolve-DriverPackage {
    param(
        [string]$PackagePath,
        [string]$FallbackDriversDirectory
    )

    if ($PackagePath) {
        $resolved = Resolve-Path -LiteralPath $PackagePath -ErrorAction Stop
        Write-Info "Using local driver package: $resolved"
        return $resolved.Path
    }

    if ($FallbackDriversDirectory) {
        $resolvedDirectory = Resolve-Path -LiteralPath $FallbackDriversDirectory -ErrorAction SilentlyContinue
        if ($resolvedDirectory) {
            Write-Info "Using local drivers directory: $resolvedDirectory"
            return $resolvedDirectory.Path
        }
    }

    return $null
}

function Expand-DriverPackageIfNeeded {
    param(
        [string]$PackagePath,
        [string]$WorkingDirectory
    )

    if (-not $PackagePath) {
        return $null
    }

    $extension = [System.IO.Path]::GetExtension($PackagePath).ToLowerInvariant()
    if ($extension -ne ".zip") {
        return $PackagePath
    }

    $extractDir = Join-Path $WorkingDirectory "expanded"
    Write-Info "Expanding driver archive..."
    Expand-Archive -LiteralPath $PackagePath -DestinationPath $extractDir -Force
    Write-Ok "Expanded archive to '$extractDir'."
    return $extractDir
}

function Get-DriverPackageMetadata {
    param([string]$SourcePath)

    if (-not $SourcePath -or -not (Test-Path -LiteralPath $SourcePath)) {
        return $null
    }

    $root = $SourcePath
    if (Test-Path -LiteralPath $SourcePath -PathType Leaf) {
        $root = Split-Path -Parent $SourcePath
    }

    $metadata = @{
        PackageRoot = $root
        InstallerPath = $null
        PrinterName = $null
        DriverName = $null
    }

    $driverSetup = Join-Path $root "DriverSetup.exe"
    if (Test-Path -LiteralPath $driverSetup -PathType Leaf) {
        $metadata.InstallerPath = $driverSetup
    }

    $configPath = Join-Path $root "config.ini"
    if (Test-Path -LiteralPath $configPath -PathType Leaf) {
        foreach ($line in Get-Content -LiteralPath $configPath) {
            if ($line -match '^\s*PrnName\s*=\s*(.+?)\s*$') {
                $metadata.PrinterName = $matches[1].Trim()
                break
            }
        }
    }

    $setupFolderName = if ([Environment]::Is64BitOperatingSystem) { "SETUP64" } else { "SETUP" }
    $candidateRoots = @(
        (Join-Path $root "$setupFolderName\ENG"),
        (Join-Path $root "$setupFolderName\CHN"),
        (Join-Path $root "SETUP64\ENG"),
        (Join-Path $root "SETUP\ENG")
    ) | Select-Object -Unique

    foreach ($candidate in $candidateRoots) {
        if (-not (Test-Path -LiteralPath $candidate -PathType Container)) {
            continue
        }

        $gpdFile = Get-ChildItem -LiteralPath $candidate -Filter *.gpd -File |
            Sort-Object Name |
            Select-Object -First 1
        if (-not $gpdFile) {
            continue
        }

        foreach ($line in Get-Content -LiteralPath $gpdFile.FullName) {
            if ($line -match '^\*ModelName:\s*"(.+?)"') {
                $metadata.DriverName = $matches[1].Trim()
                break
            }
        }

        if ($metadata.DriverName) {
            break
        }
    }

    return $metadata
}

function Find-DriverInstaller {
    param([string]$SourcePath)

    if (-not $SourcePath) {
        return $null
    }

    if (Test-Path -LiteralPath $SourcePath -PathType Leaf) {
        return Get-Item -LiteralPath $SourcePath
    }

    $patterns = @(
        "DriverSetup.exe",
        "setup.exe",
        "install.exe",
        "*.msi",
        "*.exe",
        "*.inf"
    )

    foreach ($pattern in $patterns) {
        $match = Get-ChildItem -LiteralPath $SourcePath -Recurse -File -Filter $pattern |
            Sort-Object FullName |
            Select-Object -First 1
        if ($match) {
            return $match
        }
    }

    return $null
}

function Install-DriverPackage {
    param(
        [string]$InstallerPath,
        [string]$InstallerArgs
    )

    $extension = [System.IO.Path]::GetExtension($InstallerPath).ToLowerInvariant()
    Write-Info "Installing driver package from '$InstallerPath'..."

    if ($extension -eq ".inf") {
        $null = pnputil.exe /add-driver $InstallerPath /install
        Write-Ok "INF driver imported."
        return
    }

    if ($extension -eq ".msi") {
        $arguments = "/i `"$InstallerPath`" /qn"
        if ($InstallerArgs) {
            $arguments = "$arguments $InstallerArgs"
        }

        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -PassThru -Wait
        if ($process.ExitCode -ne 0) {
            throw "Driver MSI install failed with exit code $($process.ExitCode)."
        }

        Write-Ok "MSI driver install completed."
        return
    }

    $processArgs = $InstallerArgs
    if (-not $processArgs) {
        if ([System.IO.Path]::GetFileName($InstallerPath) -ieq "DriverSetup.exe") {
            $processArgs = ""
        }
        else {
            $processArgs = "/S"
        }
    }

    $workingDirectory = Split-Path -Parent $InstallerPath
    if (Test-HasText -Value $processArgs) {
        $process = Start-Process -FilePath $InstallerPath -ArgumentList $processArgs -WorkingDirectory $workingDirectory -PassThru -Wait
    }
    else {
        $process = Start-Process -FilePath $InstallerPath -WorkingDirectory $workingDirectory -PassThru -Wait
    }
    if ($process.ExitCode -ne 0) {
        throw "Driver installer failed with exit code $($process.ExitCode)."
    }

    Write-Ok "Driver installer completed."
}

function Get-ExistingPrinter {
    param([string]$Name)

    return Get-Printer -Name $Name -ErrorAction SilentlyContinue
}

function Ensure-PrinterDriver {
    param(
        [string]$Name,
        [string]$PackagePath,
        [string]$FallbackDriversDirectory,
        [string]$InstalledDriverName,
        [string]$InstallerArgs
    )

    $printer = Get-ExistingPrinter -Name $Name
    if ($printer) {
        return $printer
    }

    if ($SkipDriverInstall) {
        Write-Fail "Printer '$Name' was not found and driver installation was skipped."
        exit 1
    }

    if (-not $PackagePath -and -not $FallbackDriversDirectory) {
        Write-Fail "Printer '$Name' was not found."
        Write-Host "Add a 'drivers' folder next to this script, or provide -DriverPackagePath." -ForegroundColor Yellow
        exit 1
    }

    Confirm-DriverInstallationStep

    $workingDirectory = New-DriverWorkingDirectory
    $driverNamesBefore = Get-InstalledPrinterDriverNames
    $package = Resolve-DriverPackage `
        -PackagePath $PackagePath `
        -FallbackDriversDirectory $FallbackDriversDirectory
    $source = Expand-DriverPackageIfNeeded -PackagePath $package -WorkingDirectory $workingDirectory
    $packageMetadata = Get-DriverPackageMetadata -SourcePath $source

    if ($packageMetadata.PrinterName -and $Name -eq "POS58 Printer") {
        Write-Info "Driver package declares printer name '$($packageMetadata.PrinterName)'."
    }

    $installer = Find-DriverInstaller -SourcePath $source

    if (-not $installer) {
        Write-Fail "No installer was found in '$source'."
        exit 1
    }

    Install-DriverPackage -InstallerPath $installer.FullName -InstallerArgs $InstallerArgs

    Start-Sleep -Seconds 2
    $printer = Get-ExistingPrinter -Name $Name
    if ($printer) {
        Write-Ok "Printer '$Name' is now installed."
        return $printer
    }

    if (-not $InstalledDriverName) {
        $driverNamesAfter = Get-InstalledPrinterDriverNames
        $newDriverNames = @($driverNamesAfter | Where-Object { $_ -notin $driverNamesBefore })
        if ($newDriverNames.Count -eq 1) {
            $InstalledDriverName = $newDriverNames[0]
            Write-Info "Detected installed printer driver '$InstalledDriverName'."
        }
        elseif ($packageMetadata.DriverName -and (Get-PrinterDriver -Name $packageMetadata.DriverName -ErrorAction SilentlyContinue)) {
            $InstalledDriverName = $packageMetadata.DriverName
            Write-Info "Using driver name from package metadata: '$InstalledDriverName'."
        }
    }

    if (-not $InstalledDriverName) {
        Write-Fail "Driver install completed, but printer '$Name' still does not exist."
        Write-Host "If the driver only installed driver files, either keep only one printer driver in the local 'drivers' folder or re-run with -DriverName." -ForegroundColor Yellow
        exit 1
    }

    return @{
        Printer = $null
        DriverName = $InstalledDriverName
    }
}

function Ensure-RelayPort {
    param(
        [string]$Name,
        [string]$HostAddress,
        [uint32]$Port
    )

    $existingPort = Get-PrinterPort -Name $Name -ErrorAction SilentlyContinue

    if (-not $existingPort) {
        Write-Info "Creating TCP/IP printer port '$Name'..."
        Add-PrinterPort -Name $Name -PrinterHostAddress $HostAddress -PortNumber $Port
        Write-Ok "Created port '$Name'."
        return
    }

    Write-Ok "Port '$Name' already exists."
}

function Ensure-PrinterQueue {
    param(
        [string]$Name,
        [string]$Driver,
        [string]$TargetPort
    )

    $printer = Get-ExistingPrinter -Name $Name
    if ($printer) {
        return $printer
    }

    if (-not $Driver) {
        Write-Fail "Printer '$Name' does not exist and no -DriverName was supplied to create it."
        exit 1
    }

    $driverExists = Get-PrinterDriver -Name $Driver -ErrorAction SilentlyContinue
    if (-not $driverExists) {
        Write-Fail "Printer driver '$Driver' is not installed."
        exit 1
    }

    Write-Info "Creating printer queue '$Name' with driver '$Driver'..."
    Add-Printer -Name $Name -DriverName $Driver -PortName $TargetPort
    Write-Ok "Created printer queue '$Name'."
    return Get-ExistingPrinter -Name $Name
}

function Remove-PrinterQueueIfBoundToPort {
    param(
        [string]$Name,
        [string]$TargetPort
    )

    $printer = Get-ExistingPrinter -Name $Name
    if (-not $printer) {
        Write-Info "Printer '$Name' was not found."
        return $false
    }

    if ($printer.PortName -ne $TargetPort) {
        Write-Info "Printer '$Name' is using port '$($printer.PortName)', not '$TargetPort'. Leaving queue in place."
        return $false
    }

    Remove-Printer -Name $Name
    Write-Ok "Removed printer queue '$Name'."
    return $true
}

function Remove-RelayPortIfUnused {
    param([string]$Name)

    $port = Get-PrinterPort -Name $Name -ErrorAction SilentlyContinue
    if (-not $port) {
        Write-Info "Port '$Name' was not found."
        return
    }

    $boundPrinters = @(Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.PortName -eq $Name })
    if ($boundPrinters.Count -gt 0) {
        $printerNames = $boundPrinters | Select-Object -ExpandProperty Name
        Write-Info "Port '$Name' is still in use by: $($printerNames -join ', '). Leaving port in place."
        return
    }

    Remove-PrinterPort -Name $Name
    Write-Ok "Removed printer port '$Name'."
}

if (-not $PortName) {
    $PortName = "Relay_$RelayHost`_$RelayPort"
}

if (-not $DriversDirectory) {
    $DriversDirectory = Join-Path $PSScriptRoot "drivers"
}

if (-not $RelayScriptPath) {
    $RelayScriptPath = Join-Path $PSScriptRoot "printer_relay.py"
}

if (-not $StartupLauncherPath) {
    $StartupLauncherPath = Join-Path $PSScriptRoot "relay_startup.vbs"
}

if (-not $GuiLauncherPath) {
    $GuiLauncherPath = Join-Path $PSScriptRoot "launch.vbs"
}

if (-not $DesktopShortcutPath) {
    $desktopDir = [Environment]::GetFolderPath("Desktop")
    $DesktopShortcutPath = Join-Path $desktopDir "Seznik EON Printer Toolkit.lnk"
}

if (-not $StartMenuShortcutPath) {
    $startMenuDir = [Environment]::GetFolderPath("Programs")
    $StartMenuShortcutPath = Join-Path $startMenuDir "Seznik EON Printer Toolkit.lnk"
}

Ensure-Admin

if ($Uninstall) {
    Write-Info "Removing relay configuration for printer '$PrinterName'..."
    Write-Info "Relay endpoint : $RelayHost`:$RelayPort"
    Write-Info "Port name      : $PortName"

    Remove-RelayStartup -LauncherPath $StartupLauncherPath
    Remove-ToolkitEnvironment
    $removedShortcut = Remove-GuiShortcuts -ShortcutPaths @($DesktopShortcutPath, $StartMenuShortcutPath)
    $removedPrinter = Remove-PrinterQueueIfBoundToPort -Name $PrinterName -TargetPort $PortName
    Remove-RelayPortIfUnused -Name $PortName

    Write-Host ""
    Write-Host "Relay uninstall summary:" -ForegroundColor Cyan
    Write-Host "  GUI shortcut removed  : $removedShortcut"
    Write-Host "  Printer queue removed : $removedPrinter"
    Write-Host "  Startup launcher path : $StartupLauncherPath"
    Write-Host "  Relay port            : $PortName"
    Write-Host ""
    Write-Host "Driver packages were not removed." -ForegroundColor Yellow
    exit 0
}

$PythonExecutablePath = Ensure-SystemPython `
    -PreferredPythonExe $PythonExecutablePath
Ensure-ToolkitEnvironment -PythonExePath $PythonExecutablePath -ToolkitDir $PSScriptRoot
Invoke-PrinterScanSetup -PythonExe $PythonExecutablePath

Write-Info "Target printer : $PrinterName"
Write-Info "Relay endpoint : $RelayHost`:$RelayPort"
Write-Info "Port name      : $PortName"
if ($DriverPackagePath) {
    Write-Info "Driver package : $DriverPackagePath"
}
elseif (Test-PathIfSet -LiteralPath $DriversDirectory) {
    Write-Info "Drivers dir    : $DriversDirectory"
}
if (Test-PathIfSet -LiteralPath $PythonExecutablePath) {
    Write-Info "Python exe     : $PythonExecutablePath"
}
if (Test-PathIfSet -LiteralPath $RelayScriptPath) {
    Write-Info "Relay script   : $RelayScriptPath"
}
if (Test-PathIfSet -LiteralPath $GuiLauncherPath) {
    Write-Info "GUI launcher   : $GuiLauncherPath"
}
Write-Info "Desktop link   : $DesktopShortcutPath"
Write-Info "Start Menu link: $StartMenuShortcutPath"
$canLaunchGui = Test-PythonModuleInstalled -PythonExe $PythonExecutablePath -ImportName "tkinter"
if (-not $canLaunchGui) {
    Write-Fail "GUI support    : tkinter is not available in the selected system Python."
    exit 1
}

$driverResult = Ensure-PrinterDriver `
    -Name $PrinterName `
    -PackagePath $DriverPackagePath `
    -FallbackDriversDirectory $DriversDirectory `
    -InstalledDriverName $DriverName `
    -InstallerArgs $DriverInstallerArgs

$printer = $driverResult
if ($driverResult -is [hashtable]) {
    $printer = $driverResult.Printer
    if (-not $DriverName -and $driverResult.DriverName) {
        $DriverName = $driverResult.DriverName
    }
}

Ensure-RelayPort -Name $PortName -HostAddress $RelayHost -Port $RelayPort

if (-not $printer) {
    $printer = Ensure-PrinterQueue -Name $PrinterName -Driver $DriverName -TargetPort $PortName
}

if ($printer.PortName -ne $PortName) {
    Write-Info "Binding printer '$PrinterName' to port '$PortName'..."
    Set-Printer -Name $PrinterName -PortName $PortName
    Write-Ok "Printer '$PrinterName' is now using '$PortName'."
}
else {
    Write-Ok "Printer '$PrinterName' is already bound to '$PortName'."
}

$startupEntry = Ensure-RelayStartup `
    -PythonExePath $PythonExecutablePath `
    -RelayScriptPath $RelayScriptPath `
    -LauncherPath $StartupLauncherPath `
    -RelayHost $RelayHost `
    -RelayPort $RelayPort
Start-RelayLauncher -LauncherPath $startupEntry
$guiShortcuts = @()
if ($canLaunchGui) {
    $guiShortcuts = Ensure-GuiShortcuts `
        -LauncherPath $GuiLauncherPath `
        -ShortcutPaths @($DesktopShortcutPath, $StartMenuShortcutPath)
}

$updatedPrinter = Get-Printer -Name $PrinterName
Write-Host ""
Write-Host "Current printer mapping:" -ForegroundColor Cyan
Write-Host "  Name : $($updatedPrinter.Name)"
Write-Host "  Port : $($updatedPrinter.PortName)"
if ($DriverName) {
    Write-Host "  Driver : $($updatedPrinter.DriverName)"
}
if ($startupEntry) {
    Write-Host "  Startup : $startupEntry"
}
foreach ($shortcutPath in $guiShortcuts) {
    Write-Host "  Shortcut : $shortcutPath"
}
