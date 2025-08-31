

# Perdanga Software Solutions

![GitHub release (latest by date)](https://img.shields.io/github/v/release/perdanger/Perdanga-Software-Solutions?color=blue) ![License](https://img.shields.io/github/license/perdanger/Perdanga-Software-Solutions?color=green) ![Chocolatey](https://img.shields.io/badge/Powered%20by-Chocolatey-brown) ![PowerShell](https://img.shields.io/badge/Powered%20by-PowerShell-blue) ![Open Source](https://img.shields.io/badge/Open%20Source-PerdangaForever-brightgreen) ![GitHub Stars](https://img.shields.io/github/stars/perdanger/Perdanga-Software-Solutions?style=social)

![Perdanga Forever](https://github.com/perdanger/Perdanga-Software-Solutions/blob/main/PSS1.5.png?raw=true)

> Perdanga Software Solutions is a PowerShell script designed to simplify the **installation**, **uninstallation**, and **management** of essential Windows software, with enhanced system cleanup and information features.

## Quick Start

1. **Run in PowerShell**:

   ```powershell
   irm https://bit.ly/PerdangaSoftwareSolutions | iex
   ```

2. **[Direct Download the Latest Version](https://github.com/perdanger/Perdanga-Software-Solutions/releases/download/1.6/PSS.1.6.rar)**

> [!IMPORTANT]  
> Ensure is run with **Administrator privileges** to avoid execution issues.

## Supported Programs

7zip.install, brave, cursoride, discord, file-converter,  
git, googlechrome, imageglass, nilesoft-shell, nvidia-app,  
obs-studio, occt, qbittorrent, revo-uninstaller, spotify,  
steam, telegram, vcredist-all, vlc, winrar, wiztree.

## 💥 Features

### 1. Main Menu

- **[A] Install All Programs**
- **[G] Select Specific Programs [GUI]**
- **[U] Uninstall Programs [GUI]**
- **[C] Install Custom Package**
- **[T] Disable Windows Telemetry**
- **[X] Activate Spotify**
- **[W] Activate Windows**
- **[N] Update Windows**
- **[F] Create Unattend.xml File [GUI]**
- **[S] System Cleanup [GUI]**
- **[P] Import & Install from File**
- **[I] Show System Information**

> [!TIP]  
> To install specific programs, enter their numbers (e.g., `1 5 17`, or `1,5,17`) from the supported programs list.  
> Type `perdanga` for a hidden cheese! 🧀

### 2. Unattend.xml Creator [F]

Customize your Windows installation with these options:

- **System Settings**: Set computer name, admin user, and password.
- **Regional Preferences**: Configure UI language, system locale, user locale, time zone, and keyboard layouts.
- **OOBE Bypass**: Skip EULA, local/online account setup, and wireless configuration.
- **System Tweaks**: Enable file extensions, disable SmartScreen, enable system restore, or disable app suggestions.
- **Bloatware Removal**: Remove unwanted pre-installed apps during setup.
- **Output**: Saves the `autounattend.xml` file to your Desktop for use with a Windows installation USB.

![Unattend.xml Creator](https://github.com/perdanger/Perdanga-Software-Solutions/blob/main/UnattendxmlFile.png?raw=true)

### 3. System Cleanup [S]

Enhanced system cleanup with dynamic application cache detection:

- **Windows Temporary Files**: Clean standard Windows temp directories.
- **Application Caches**: Automatically detect and clean caches from popular applications including:
  - NVIDIA Cache
  - DirectX Shader Cache
  - Steam Cache
  - Discord Cache
  - EA App Cache
  - Spotify Cache
  - Visual Studio Code Cache
  - And many more...
- **Browser Caches**: Detect and clean caches from installed web browsers.
- **System Caches**: Windows Update Cache, Prefetch files, and Recycle Bin.

![System Cleanup](https://github.com/perdanger/Perdanga-Software-Solutions/blob/main/SystemCleanup.png?raw=true)

### 4. System Information [I]

Enhanced system information display with a modern graphical interface:

- **Operating System**: Name, version, build, architecture, and product ID.
- **Processor**: Name, core count, and virtualization status.
- **System Hardware**: Manufacturer, model, motherboard, BIOS version, Secure Boot, and TPM status.
- **Memory (RAM)**: Total installed memory and detailed module information.
- **Video Card(s)**: GPU name, VRAM, and driver version.
- **Disk Drives**: Physical disks with partition information and free space.
- **Network Adapters**: Detailed network configuration including IP, MAC, gateway, and DNS.

![System Information](https://github.com/perdanger/Perdanga-Software-Solutions/blob/main/SystemInfo.png?raw=true)

### 5. Import & Install from File [P]

Imports a JSON file containing a list of program names to install:

- Select a `.json` file from your Desktop.
- Validates and installs the listed programs.
- Example JSON format: `["7zip.install", "brave", "vlc"]`.

## 🛠️ Troubleshooting

- **Script Fails to Run**:
  - Ensure PowerShell is run with **Administrator privileges**.
  - Verify PowerShell version (5.1 or higher): `Get-Host | Select-Object Version` or `$PSVersionTable.PSVersion`.
  - Check your execution policy: `Get-ExecutionPolicy`. If restricted, set it to `RemoteSigned` with `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned`.

- **GUI Unavailable**:
  - If `[G]`, `[U]`, `[F]`, `[S]`, or `[I]` options fail, ensure `System.Windows.Forms` is available. Alternatively, use CLI-based options.

- **Package Not Found**:
  - For `[C]`, confirm the package ID exists in the Chocolatey repository: `choco search <id>`.

- **Logs**:
  - Review `install_log_YYYYMMDD_HHMMSS.txt` in the script directory or `%TEMP%` folder for detailed error information.

## 📜 Version History

### Version 1.6 (31.08.2025)
- Added dynamic application cache cleaning functionality
- Enhanced System Information GUI with a new graphical layout (Gemini-themed)
- Added Secure Boot and TPM status information to the System Information
- Updated disk information display
- Enhanced browser cache detection in the cleanup function

### Version 1.5 (28.07.2025)
- Initial release with core functionality
- Program installation and uninstallation
- System cleanup features
- Unattend.xml creator
- Windows activation and Spotify activation

## 📜 Credits

This project incorporates code from the following third-party sources:

- **SpotX** (Spotify activation feature)
  - Repository: SpotX-Official
  - Copyright © 2025 SpotX-Official
  - License: MIT License

- **Microsoft Activation Scripts** (Windows activation feature)
  - Repository: massgravel/Microsoft-Activation-Scripts
  - Copyright © 2025 massgravel
  - License: GPL-3.0

<p align="center"><b>⚡PERDANGA FOREVER⚡</b></p>
