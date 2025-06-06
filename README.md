# Perdanga Software Solutions

![image alt](https://github.com/perdanger/Perdanga-Software-Solutions/blob/main/1.3.png?raw=true)

## Overview

Perdanga Software Solutions is a powerful and easy-to-use PowerShell script designed to automate the installation, uninstallation, and management of essential Windows software using. It provides CLI and GUI interfaces for managing Software and Windows.

The scripts consist of:

- `RunPerdangaSoftwareSolutions.bat`: A batch script that verifies administrative privileges and launches the PowerShell script.
- `PerdangaSoftwareSolutions.ps1`: The core PowerShell script that manages software installation, uninstallation, Windows updates, Windows activation, Spotify activation, and telemetry control.
- `config.ps1`: Contains user-configurable settings, such as the program list.
- `functions.ps1`: Defines helper functions used by the main script.

## Features

- **Automated Software Installation**: Install a curated list of essential Windows programs via Chocolatey.
- **Custom Package Installation**: Install any Chocolatey package by specifying its exact package ID.
- **Program Uninstallation**: Uninstall Chocolatey-installed programs via a graphical interface.
- **Dual Interface**: Choose between a command-line interface or a graphical user interface for program selection and uninstallation (GUI availability depends on system configuration).
- **Windows Updates**: Check and install Windows updates using the PSWindowsUpdate module.
- **Windows Activation**: Optional activation using an external script from `https://get.activated.win`.
- **Spotify Activation**: Enhance Spotify with additional features using an external script from SpotX-Official GitHub.
- **Telemetry Control**: Disable Windows telemetry services and registry settings for enhanced privacy (new in v1.3).
- **Detailed Logging**: Logs all actions to a timestamped file for easy troubleshooting.
- **Robust Error Handling**: Includes checks for PowerShell version, administrative privileges, and Chocolatey installation.

## Supported Programs

The script automates the installation of the following essential software via Chocolatey:

- 7zip.install
- brave
- discord
- file-converter
- git
- googlechrome
- gpu-z
- hwmonitor
- imageglass
- nvidia-app
- obs-studio
- occt
- qbittorrent
- revo-uninstaller
- spotify
- steam
- telegram
- vcredist-all
- vlc
- winrar
- wiztree

## Setup Instructions

1. **Run as Administrator**:
   - Right-click `RunPerdangaSoftwareSolutions.bat` and select **Run as Administrator** to ensure proper permissions.
2. **Verify PowerShell**:
   - The script checks for PowerShell. If missing, download it from Microsoft's official site.
3. **Install Chocolatey**:
   - If Chocolatey is not installed, the script will prompt for automatic installation.

## Usage

1. **Main Menu**:
   - The script presents an intuitive menu with the following options:
     - **\[A\] Install All Programs**
     - **\[G\] Select Specific Programs via GUI**
     - **\[U\] Uninstall Programs via GUI**
     - **\[C\] Install Custom Package**
     - **\[T\] Disable Windows Telemetry**
     - **\[X\] Activate Spotify**
     - **\[W\] Activate Windows**
     - **\[N\] Update Windows**
     - **\[E\] Exit Script**
   - Alternatively, enter program numbers (e.g., '`1`' '`1 5 17`', or '`1,5,17`') to install specific programs from the predefined list.

## Troubleshooting

- **Script Fails to Run**: Ensure `RunPerdangaSoftwareSolutions.bat` is run as Administrator.
- **PowerShell Version Error**: Upgrade to PowerShell 5 or higher.
- **GUI Unavailable**: If options `G` or `U` fail, your system may lack `System.Windows.Forms`. Use CLI options instead.
- **Package Not Found**: For option `C`, ensure the entered package ID is valid and exists in the Chocolatey repository.
- **Check Logs**: Review `install_log_YYYYMMDD_HHMMSS.txt` for detailed error messages.

## License

This project incorporates code from the following third-party sources:

- **SpotX**: Used for Spotify activation feature.
  - Repository: https://github.com/SpotX-Official/SpotX
  - Copyright (c) 2025 SpotX-Official
  - License: MIT License
- **Microsoft Activation Scripts**: Used for Windows activation feature.
  - Repository: https://github.com/massgravel/Microsoft-Activation-Scripts
  - Copyright (c) 2025 massgravel
  - License: GPL-3.0
