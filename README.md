Perdanga Software Solutions

Perdanga Software Solutions is a PowerShell script designed to simplify the installation, uninstallation, and management of essential Windows software, with enhanced system cleanup and information features.

Quick Start

Run in PowerShell:

irm [https://bit.ly/PerdangaSoftwareSolutions](https://bit.ly/PerdangaSoftwareSolutions) | iex


Direct Download the Latest Version (v1.7)

$$\!IMPORTANT$$

Ensure the script is run with Administrator privileges to avoid execution issues.

What's New in v1.7

Enhanced Feature: GUI System Manager! The [G] option now launches the powerful graphical user interface (Invoke-PerdangaSystemManager v5.6) for comprehensive management of power plans, system tweaks, maintenance, software, and Windows Update settings.

New Supported Programs: Added qbittorrent, microsoft-edge, and crystaldiskinfo to the main installation menu.

Menu Update: The Windows Activation option is now labeled [A] Activate Windows (MAS) for better clarity.

Supported Programs

7zip.install, brave, cursoride, discord, file-explorer-tabs, firefox, git, googlechrome, notepadplusplus, obs-studio, open-shell, powertoys, qbit, steam, vscode, vlc, qbittorrent, microsoft-edge, crystaldiskinfo

Main Menu Options

Option

Description



$$A$$



Activate Windows (MAS) - NEW Label



$$C$$



Install program from Chocolatey ID



$$G$$



Launch GUI System Manager (BETA) - ENHANCED



$$R$$



Remove program (Select from list of all installed programs)



$$U$$



Uninstall (Select from official Windows installed applications list)



$$F$$



Flush DNS & Reset Net/Firewall settings



$$S$$



SpotX (Spotify Ad-Blocker/Features)



$$I$$



System Information



$$L$$



Launch Logs



$$Q$$



Quit

📦 Batch Install (CLI)

Use the -BatchInstall parameter to install multiple programs without using the main menu.

# Example: Installs 7zip and Brave
.\[scriptname].ps1 -BatchInstall '7zip.install', 'brave'


📜 Programs from JSON

The script can read a list of desired packages from a JSON file named packages.json placed in the same directory as the script file from your Desktop.

Validates and installs the listed programs.

Example JSON format: ["7zip.install", "brave", "vlc"].

🛠️ Troubleshooting

Script Fails to Run:

Ensure PowerShell is run with Administrator privileges.

Verify PowerShell version (5.1 or higher): Get-Host | Select-Object Version or $PSVersionTable.PSVersion.

Check your execution policy: Get-ExecutionPolicy. If restricted, set it to RemoteSigned with Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned.

GUI Unavailable:

If the [G] option fails, ensure System.Windows.Forms is available. Alternatively, use CLI-based options.

Package Not Found:

For [C], confirm the package ID exists in the Chocolatey repository: choco search <id>.

Logs:

Review install_log_YYYYMMDD_HHMMSS.txt in the script directory or %TEMP% folder for detailed error information.

📜 Credits

This project incorporates code from the following third-party sources:

SpotX (Spotify activation feature)

Repository: SpotX-Official

Copyright © 2025 SpotX-Official

License: MIT License

Microsoft Activation Scripts (Windows activation feature)

Repository: massgravel/Microsoft-Activation-Scripts

Copyright © 2025 massgravel

License: GPL-3.0
