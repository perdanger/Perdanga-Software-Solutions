# Perdanga Software Solutions

## Overview

![image alt](https://github.com/perdanger/Perdanga-Software-Solutions/blob/main/PerdangaSoftwareSolutions.png?raw=true)Perdanga Software Solutions is a powerful and easy-to-use PowerShell script designed to automate the installation, uninstallation, and management of essential Windows software using Chocolatey. It also provides convenient options for managing Windows updates and activation, all through an intuitive command-line interface and an optional graphical user interface.

The scripts consist of:

- `RunPerdangaSoftwareSolutions.bat`: A batch script that verifies administrative privileges and launches the PowerShell script.
- `PerdangaSoftwareSolutions.ps1`: The core PowerShell script that manages software installation, uninstallation, Windows updates, and activation.

## Features

- **Automated Software Installation**: Install a curated list of essential Windows programs via Chocolatey.
- **Custom Package Installation**: Install any Chocolatey package by specifying its exact package ID.
- **Program Uninstallation**: Uninstall Chocolatey-installed programs via a graphical interface.
- **Dual Interface**: Choose between a command-line interface or a graphical user interface for program selection and uninstallation (GUI availability depends on system configuration).
- **Windows Updates**: Check and install Windows updates using the PSWindowsUpdate module.
- **Windows Activation**: Optional activation using an external script from `https://get.activated.win`.
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

1. **Download the Scripts**:

   - Place `RunPerdangaSoftwareSolutions.bat` and `PerdangaSoftwareSolutions.ps1` in the same directory.

2. **Run as Administrator**:

   - Right-click `RunPerdangaSoftwareSolutions.bat` and select **Run as Administrator** to ensure proper permissions.

3. **Verify PowerShell**:

   - The script checks for PowerShell. If missing, download it from Microsoft's official site.

4. **Install Chocolatey**:

   - If Chocolatey is not installed, the script will prompt for automatic installation.

## Usage

1. **Main Menu**:

   - The script presents an intuitive menu with the following options:
     - **\[A\] Install All Programs**
     - **\[G\] Select Specific Programs via GUI**
     - **\[U\] Uninstall Programs via GUI**
     - **\[C\] Install Custom Package**
     - **\[W\] Activate Windows**
     - **\[N\] Update Windows**
     - **\[E\] Exit Script**
   - Alternatively, enter program numbers (e.g., '`1`' '`1 5 17'`, or '`1,5,17'`) to install specific programs from the predefined list.

## Troubleshooting

- **Script Fails to Run**: Ensure `RunPerdangaSoftwareSolutions.bat` is run as Administrator.
- **PowerShell Version Error**: Upgrade to PowerShell 5 or higher.
- **GUI Unavailable**: If options `G` or `U` fail, your system may lack `System.Windows.Forms`. Use CLI options instead.
- **Package Not Found**: For option `C`, ensure the entered package ID is valid and exists in the Chocolatey repository.
- **Check Logs**: Review `install_log_YYYYMMDD_HHMMSS.txt` for detailed error messages.