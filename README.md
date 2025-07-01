# Perdanga Software Solutions

![GitHub release (latest by date)](https://img.shields.io/github/v/release/perdanger/Perdanga-Software-Solutions?color=blue) ![License](https://img.shields.io/github/license/perdanger/Perdanga-Software-Solutions?color=green) ![Chocolatey](https://img.shields.io/badge/Powered%20by-Chocolatey-brown)

> Perdanga Software Solutions is a user-friendly PowerShell script designed to simplify the **installation**, **uninstallation**, and **management** of essential Windows software.

## Quick Start

1. **Run in PowerShell**:

   ```powershell
   irm https://raw.githubusercontent.com/perdanger/Perdanga-Software-Solutions/main/PerdangaLoader.ps1 | iex
   ```

2. **Alternatively, download the archive and run the batch file:**

   ```
  <a href="https://github.com/Parcoil/Sparkle/releases/latest">Download Latest</a>
   ```

> [!IMPORTANT]  
> Ensure PowerShell is run with **Administrator privileges** to avoid execution issues.

## Supported Programs

7zip.install, brave, cursoride, discord, file-converter,  
git, googlechrome, imageglass, nilesoft-shell, nvidia-app,  
obs-studio, occt, qbittorrent, revo-uninstaller, spotify,  
steam, telegram, vcredist-all, vlc, winrar, wiztree.

## 💥 Features

### 1. Main Menu

- **[A] Install All Programs**  

- **[G] Select Specific Programs**  

- **[U] Uninstall Programs**  

- **[C] Install Custom Package**  

- **[T] Disable Windows Telemetry**  

- **[X] Activate Spotify**  

- **[W] Activate Windows**  

- **[N] Update Windows**  

- **[F] Create Unattend.xml File**

> [!TIP]  
> To install specific programs, enter their numbers (e.g., `1`, `1 5 17`, or `1,5,17`) from the supported programs list.  
> Type `perdanga` for a hidden surprise! 🧀

### 2. Unattend.xml Creator [F]

Customize your Windows installation with these options:

- **System Settings**: Set computer name, admin user, and password.
- **Regional Preferences**: Configure UI language, system locale, user locale, time zone, and keyboard layouts.
- **OOBE Bypass**: Skip EULA, local/online account setup, and wireless configuration.
- **System Tweaks**: Enable file extensions, disable SmartScreen, enable system restore, or disable app suggestions.
- **Bloatware Removal**: Remove unwanted pre-installed apps during setup.
- **Output**: Saves the `autounattend.xml` file to your Desktop for use with a Windows installation USB.

## 🛠️ Troubleshooting

- **Script Fails to Run**:

  - Ensure PowerShell is run with **Administrator privileges**.
  - Verify PowerShell version (5.1 or higher): `Get-Host | Select-Object Version` or `$PSVersionTable.PSVersion`.
  - Check your execution policy: `Get-ExecutionPolicy`. If restricted, set it to `RemoteSigned` with `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned`.

- **GUI Unavailable**:

  - If `[G]`, `[U]`, or `[F]` options fail, ensure `System.Windows.Forms` is available. Alternatively, use CLI-based options.

- **Package Not Found**:

  - For `[C]`, confirm the package ID exists in the Chocolatey repository: `choco search <id>`.

- **Logs**:

  - Review `install_log_YYYYMMDD_HHMMSS.txt` in the script directory or `%TEMP%` folder for detailed error information.

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

# Perdanga Forever
