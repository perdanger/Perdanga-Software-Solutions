# Perdanga Software Solutions 💡

![Perdanga Software Solutions GUI](https://github.com/perdanger/Perdanga-Software-Solutions/blob/main/1.4.png?raw=true)

---

## 📋 Overview

Perdanga Software Solutions is a powerful and user-friendly PowerShell script designed to streamline the **installation**, **uninstallation**, and **management** of essential Windows software using Chocolatey.

### Quick Start

1. **Run in PowerShell (Administrator)**:
   ```powershell
   irm https://raw.githubusercontent.com/perdanger/Perdanga-Software-Solutions/main/PerdangaLoader.ps1 | iex
   ```

2. **Or Download and Run**:
   - Download the archive: [Version 1.4](https://github.com/perdanger/Perdanga-Software-Solutions/archive/refs/tags/1.4.zip)
   - Run `RunPerdangaSoftwareSolutions.bat` as an Administrator.

---

## Supported Programs

The script automates the installation of key software via Chocolatey:  
7zip.install, brave, cursoride, discord, file-converter,  
git, googlechrome, imageglass, nilesoft-shell, nvidia-app,  
obs-studio, occt, qbittorrent, revo-uninstaller, spotify,  
steam, telegram, vcredist-all, vlc, winrar, wiztree.

---

## ⚙️ Features

### 1. Main Menu

The script offers an intuitive CLI menu with the following options:

- **[A] Install All Programs**: Install all supported programs in one go.
- **[G] Select Specific Programs**: Choose programs to install using a graphical interface.
- **[U] Uninstall Programs**: Remove Chocolatey-installed programs via a GUI.
- **[C] Install Custom Package**: Install a specific Chocolatey package by entering its ID.
- **[T] Disable Windows Telemetry**: Disable telemetry services and tweak registry settings.
- **[X] Activate Spotify**: Apply SpotX enhancements for an improved Spotify experience.
- **[W] Activate Windows**: Run an external activation script from [get.activated.win](https://get.activated.win).
- **[N] Update Windows**: Check for and install the latest Windows updates.
- **[F] Create Unattend.xml File**: Generate an `autounattend.xml` file for automated Windows setup.
- **[E] Exit Script**: Safely exit and disable Chocolatey auto-confirmation.

Enter program numbers (e.g., `1`, `1 5 17`, or `1,5,17`) to install specific programs from the list.  
Type `perdanga` for a hidden surprise! 🧀

### 2. Unattend.xml Creator

Customize your Windows setup with the following options:
- **System Settings**: Configure computer name, admin user, and password.
- **Regional Preferences**: Set UI language, system locale, user locale, time zone, and keyboard layouts.
- **OOBE Bypass**: Skip EULA, local/online account setup, and wireless configuration.
- **System Tweaks**: Show file extensions, disable SmartScreen, enable system restore, or disable app suggestions.
- **Bloatware Removal**: Remove unwanted apps during Windows setup.
- **Output**: Saves the `autounattend.xml` file to your Desktop for use with a Windows installation USB.

---

## 🛠️ Troubleshooting

- **Script Fails to Run**:
  - Ensure you run the script or `RunPerdangaSoftwareSolutions.bat` with **Administrator privileges**.
  - Verify PowerShell 5.1 or higher: `$PSVersionTable.PSVersion`.

- **GUI Unavailable**:
  - If `[G]`, `[U]`, or `[F]` options fail, your system may lack `System.Windows.Forms`. Use CLI alternatives.

- **Package Not Found**:
  - For `[C]`, verify the package ID in the Chocolatey repository: `choco search <id>`.

- **Logs**:
  - Check `install_log_YYYYMMDD_HHMMSS.txt` in the script directory or TEMP folder for errors.

---

## 📜 License

This project incorporates code from the following third-party sources:

- **SpotX** (Spotify activation feature)
  - Repository: [SpotX-Official](https://github.com/SpotX-Official/SpotX)
  - Copyright © 2025 SpotX-Official
  - License: MIT License

- **Microsoft Activation Scripts** (Windows activation feature)
  - Repository: [massgravel/Microsoft-Activation-Scripts](https://github.com/massgravel/Microsoft-Activation-Scripts)
  - Copyright © 2025 massgravel
  - License: GPL-3.0

---

*Perdanga Forever*
