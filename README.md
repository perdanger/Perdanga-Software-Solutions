

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
- **[X] Activate Spotify**
- **[W] Activate Windows**
- **[N] Update Windows**
- **[F] Create Unattend.xml File [GUI]**
- **[S] System Cleanup [GUI]**
- **[L] Import & Install from File**
- **[I] Show System Information [GUI]**
- **[P] Perdanga System Manager [GUI]**
- **[R] System Repair**

> [!TIP]  
> To install specific programs, enter their numbers (e.g., `1 5 17`, or `1,5,17`) from the supported programs list.  
> Type `perdanga` for a hidden cheese! 🧀

### 2. Perdanga System Manager [P]

Efficient Windows Management for Software, Tweaks, Power, Updates.

- **System Tweaks**:
  - Apply or undo tweaks in categories like Privacy, System, Network, Appearance, Explorer, Input, and Advanced Debloat.

![System Tweaks](https://github.com/perdanger/Perdanga-Software-Solutions/blob/main/System%20Tweaks.png?raw=true)

- **Install Apps**:
  - Batch install, update, or uninstall software using Winget or Chocolatey.

![Install Apss](https://github.com/perdanger/Perdanga-Software-Solutions/blob/main/IntallApps.png?raw=true)

- **Maintenance:**
  - Set DNS (Google, Cloudflare, or reset to DHCP).
  - Reset Windows Update components.
  - System Scan (CHKDSK, SFC, DISM).
  - Deep Cleanup (Disk Cleanup + Component Store).
  - Network Reset (Flush DNS, reset IP/Winsock).
  - Reinstall Package Manager (Winget/Chocolatey).
  - Remove OneDrive.
  - Run external tools: O&O ShutUp10, Sysinternals Autoruns, Adobe CC Cleaner.
  - Enable SSH Server.
- **Windows Update Config**:
  - Default (Enabled).
  - Security Only (Defer features for 1 year).
  - Disabled.
- **Quick Tools**: 
  - Control Panel, Network Connections, Power Options.
  - Sound Settings, System Properties, User Accounts.
  - Registry Editor, Services, God Mode.
  - Task Manager, Restart Explorer.

### 3. Unattend.xml Creator [F]

Customize your Windows installation with these options:

- **System Settings**: Set computer name, admin user, and password.
- **Regional Preferences**: Configure UI language, system locale, user locale, time zone, and keyboard layouts.
- **OOBE Bypass**: Skip EULA, local/online account setup, and wireless configuration.
- **System Tweaks**: Enable file extensions, disable SmartScreen, enable system restore, or disable app suggestions.
- **Bloatware Removal**: Remove unwanted pre-installed apps during setup.
- **Output**: Saves the `autounattend.xml` file to your Desktop for use with a Windows installation USB.

> [!NOTE]  
> Learn more about answer files (Unattend.xml) [Microsoft Documentation](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/update-windows-settings-and-scripts-create-your-own-answer-file-sxs?view=windows-11).

![Unattend.xml Creator](https://github.com/perdanger/Perdanga-Software-Solutions/blob/main/UnattendxmlFile.png?raw=true)

### 4. System Cleanup [S]

Enhanced system cleanup with dynamic application cache detection:

- **Windows Temporary Files**
- **Application Caches**
- **Browser Caches**
- **System Caches**

![System Cleanup](https://github.com/perdanger/Perdanga-Software-Solutions/blob/main/SystemCleanup.png?raw=true)

### 5. System Information [I]

- **Operating System**
- **Processor**
- **System Hardware**
- **Memory (RAM)**
- **Video Card(s)**
- **Disk Drives**
- **Network Adapters**

![System Information](https://github.com/perdanger/Perdanga-Software-Solutions/blob/main/SystemInfo.png?raw=true)

### 5. Import & Install from File [L]

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

![Logo](https://github.com/perdanger/Perdanga-Software-Solutions/blob/main/PerdangaForeverLogo.png?raw=true)
<p align="center"><b>⚡PERDANGA FOREVER⚡</b></p>
