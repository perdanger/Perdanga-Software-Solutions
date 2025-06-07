# Perdanga Software Solutions

![image alt](https://github.com/perdanger/Perdanga-Software-Solutions/blob/main/1.3.png?raw=true)
# Overview

Perdanga Software Solutions is a powerful and easy-to-use PowerShell script designed to automate the installation, uninstallation, and management of essential Windows software.

1. **PowerShell Launch**:
     Run the following command in PowerShell as an Administrator:

     ```
     irm https://raw.githubusercontent.com/perdanger/Perdanga-Software-Solutions/main/PerdangaLoader.ps1 | iex
     ```
   - This method downloads and executes the PerdangaLoader.ps1 script directly from the repository to initiate the setup.
2. **Download archive**:
     Run the bat file as an administrator:

    ```
    https://github.com/perdanger/Perdanga-Software-Solutions/archive/refs/tags/1.3.zip
    ```
   

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

- **Script Fails to Run**: Ensure `RunPerdangaSoftwareSolutions.bat` is run as Administrator or the PowerShell command is executed with administrative privileges.
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
