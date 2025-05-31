# Perdanga Software Solutions

## Overview
https://github.com/perdanger/Perdanga-Software-Solutions/blob/cfab64e4442c46059d256c1d11cd9170b903573d/PerdangaSoftwareSolutions.png

Perdanga Software Solutions is a powerful and easy-to-use PowerShell script designed to automate the installation of essential Windows software using Chocolatey. It also provides convenient options for managing Windows updates and activation, all through an intuitive command-line interface and an optional graphical user interface.

The scripts consist of:

- `RunPerdangaSoftwareSolutions.bat`: A batch script that verifies administrative privileges and launches the PowerShell script.
- `PerdangaSoftwareSolutions.ps1`: The core PowerShell script that manages software installation, Windows updates, and activation.

## Features

- **Automated Software Installation**: Install a curated list of essential Windows programs via Chocolatey.
- **Dual Interface**: Choose between a command-line interface or a graphical user interface for program selection (GUI availability depends on system configuration).
- **Windows Updates**: Check and install Windows updates using the PSWindowsUpdate module.
- **Windows Activation**: Optional activation using an external script 

  ```powershell
  https://get.activated.win
  ```
- **Detailed Logging**: Logs all actions to a timestamped file for easy troubleshooting.
- **Robust Error Handling**: Includes checks for PowerShell version, administrative privileges, and Chocolatey installation.

## Supported Programs

The script automates the installation of the following essential software via Chocolatey:

- 7zip.install
- brave
- file-converter
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


- **\[A\] Install All Programs**: Install all listed programs.
- **\[G\] Select Specific Programs via GUI**
- **\[W\] Activate Windows**: Run an external Windows activation script.
- **\[N\] Update Windows**: Check and install Windows updates.
- **\[E\] Exit Script**: Exit the program.


- Alternatively, enter program numbers (e.g., `1`, `1 5 17`, or `1,5,17`) to install specific programs.

1. **Program Selection**:

   - **Single Program**: Enter a number (e.g., `1`) to install one program.
   - **Multiple Programs**: Enter numbers separated by spaces or commas (e.g., `1 5 17` or `1,5,17`).
   - **GUI Selection**: Choose option `G` to select programs via checkboxes.

2. **Logs**:

   - All actions are logged to `install_log_YYYYMMDD_HHMMSS.txt` in the script directory.

## Troubleshooting

- **Script Fails to Run**: Ensure `RunPerdangaSoftwareSolutions.bat` is run as Administrator.
- **PowerShell Version Error**: Upgrade to PowerShell 5 or higher.
- **Chocolatey Installation Fails**: Verify your internet connection and access to `https://chocolatey.org/install.ps1`.
- **GUI Unavailable**: If option `G` fails, your system may lack `System.Windows.Forms`. Use CLI options instead.
- **Check Logs**: Review `install_log_YYYYMMDD_HHMMSS.txt` for detailed error messages.
