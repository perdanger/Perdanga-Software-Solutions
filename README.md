# Perdanga Software Solutions 💡

![GitHub release (latest by date)](https://img.shields.io/github/v/release/perdanger/Perdanga-Software-Solutions?color=blue)
![License](https://img.shields.io/github/license/perdanger/Perdanga-Software-Solutions?color=green)
![Chocolatey](https://img.shields.io/badge/Powered%20by-Chocolatey-brown)

![Perdanga Software Solutions GUI](https://github.com/perdanger/Perdanga-Software-Solutions/blob/main/1.4.png?raw=true)

## Overview

Perdanga Software Solutions is a robust and intuitive PowerShell script designed to simplify the installation, uninstallation, and management of essential Windows software.

1. **Run the following command in PowerShell as an Administrator:**

   ```
   irm https://raw.githubusercontent.com/perdanger/Perdanga-Software-Solutions/main/PerdangaLoader.ps1 | iex
   ```

2. **Alternatively, download the archive and run the batch file as an Administrator:**

   ```
   https://github.com/perdanger/Perdanga-Software-Solutions/archive/refs/tags/1.4.zip
   ```

## Supported Programs

The script automates the installation of key software via Chocolatey:  
- 7zip.install
- brave
- cursoride
- discord
- file-converter
- git
- googlechrome
- imageglass
- nilesoft-shell
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

## Features ⚙️

1. **Main Menu**:

   - The script provides a user-friendly CLI menu with the following options:
     - **\[A\] Install All Programs**: Install all supported programs effortlessly.
     - **\[G\] Select Specific Programs**: Choose specific programs to install via a graphical interface.
     - **\[U\] Uninstall Programs**: Remove Chocolatey-installed programs using a GUI.
     - **\[C\] Install Custom Package**: Install a specific Chocolatey package by entering its ID.
     - **\[T\] Disable Windows Telemetry**: Disable telemetry services and modify registry settings.
     - **\[X\] Activate Spotify**: Apply SpotX enhancements to Spotify for an enhanced experience.
     - **\[W\] Activate Windows**: Run an external activation script from `https://get.activated.win`.
     - **\[N\] Update Windows**: Check for and install the latest Windows updates.
     - **\[F\] Create Unattend.xml File**: Generate an `autounattend.xml` file for automated Windows setup.
     - **\[E\] Exit Script**: Safely exit the script and disable Chocolatey auto-confirmation.
   - Alternatively, enter program names (e.g., `1`, `1 5 17`, or `1,5,17`) to install specific programs from the list.
   - Enter the secret word `perdanga` for a hidden cheese! 🧀

2. **Unattend.xml Creator**:
     - Configure computer name, admin user, and password settings.
     - Set regional preferences (UI language, system locale, user locale, time zone, keyboard layouts).
     - Enable OOBE bypass options (e.g., skip EULA, local/online account setup, wireless configuration).
     - Apply system tweaks (e.g., show file extensions, disable SmartScreen, enable system restore, or disable app suggestions).
     - Remove bloatware during Windows setup for a cleaner installation.
     - Save the generated `autounattend.xml` to your Desktop for use with a Windows installation USB.

## Troubleshooting 🛠️

- **Script Fails to Run**:
  - Ensure the script or `RunPerdangaSoftwareSolutions.bat` is executed with Administrator privileges.
  - Verify PowerShell 5.1 or higher is installed by running `$PSVersionTable.PSVersion`.
- **GUI Unavailable**:
  - If options `[G]`, `[U]`, or `[F]` fail, your system may lack `System.Windows.Forms`. Use CLI alternatives instead.
- **Package Not Found**:
  - For option `[C]`, verify the package ID exists in the Chocolatey repository by running `choco search <id>`.
- **Log Files**:
  - Check `install_log_YYYYMMDD_HHMMSS.txt` in the script directory or TEMP folder for detailed error logs.

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
