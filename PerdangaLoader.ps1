<#
.SYNOPSIS
    Author: Roman Zhdanov
    Version: 1.5
    Last Modified: 28.07.2025
.DESCRIPTION
    Perdanga Software Solutions is a PowerShell script designed to simplify the installation, 
    uninstallation, and management of essential Windows software.

DISCLAIMER: This script contains features that download and execute third-party scripts. 
(https://github.com/SpotX-Official/SpotX)
(https://github.com/massgravel/Microsoft-Activation-Scripts)
#>

# ================================================================================
#                                 PART 1: INITIALIZATION
# ================================================================================

# Set log file name with a timestamp and use the script's directory.
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
# When run from `irm | iex`, $PSScriptRoot is not available. Default to a temp path.
$scriptDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($scriptDir)) {
    $scriptDir = $env:TEMP
}
$script:logFile = Join-Path -Path $scriptDir -ChildPath "install_log_$timestamp.txt"

# ================================================================================
#                                 PART 2: CORE FUNCTIONS
# ================================================================================

# ENHANCED FUNCTION: Writes messages to both the console and the log file.
# Automatically prefixes messages with "ERROR:" or "WARNING:" based on HostColor
# for standardized output and cleaner calling code.
function Write-LogAndHost {
    param (
        [string]$Message,
        [string]$LogPrefix = "", # Optional context, e.g., function name.
        [string]$HostColor = "White",
        [switch]$NoLog,
        [switch]$NoHost,
        [switch]$NoNewline
    )

    # Standardize host output for errors and warnings.
    $hostOutput = $Message
    $logOutput = $Message 
    
    if ($HostColor -eq 'Red' -and -not ($Message -match '^(ERROR|FATAL)')) {
        $hostOutput = "ERROR: $Message"
        $logOutput = "ERROR: $Message"
    }
    elseif (($HostColor -eq 'Yellow' -or $HostColor -eq 'DarkYellow') -and -not ($Message -match '^WARNING')) {
        $hostOutput = "WARNING: $Message"
        $logOutput = "WARNING: $Message"
    }
    
    # Construct the full log message with a timestamp.
    $fullLogMessage = "[$((Get-Date))] "
    if (-not [string]::IsNullOrEmpty($LogPrefix)) {
        $fullLogMessage += "[$LogPrefix] "
    }
    $fullLogMessage += $logOutput

    if (-not $NoLog) {
        try {
            $fullLogMessage | Out-File -FilePath $script:logFile -Append -Encoding UTF8 -ErrorAction Stop
        }
        catch {
            # Critical failure, bypass our own function to avoid loops.
            Write-Host "FATAL: Could not write to log file at $($script:logFile). Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    if (-not $NoHost) {
        if ($NoNewline) {
            Write-Host $hostOutput -ForegroundColor $HostColor -NoNewline
        } else {
            Write-Host $hostOutput -ForegroundColor $HostColor
        }
    }
}

# Function to install Chocolatey if it's not present.
function Install-Chocolatey {
    try {
        Write-LogAndHost "Chocolatey is not installed. Would you like to install it? (Type y/n then press Enter)" -HostColor Yellow -LogPrefix "Install-Chocolatey"
        $confirmInput = Read-Host
        if ($confirmInput.Trim().ToLower() -eq 'y') {
            Write-LogAndHost "User chose to install Chocolatey." -NoHost
            try {
                Write-LogAndHost "Installing Chocolatey..." -NoLog
                Set-ExecutionPolicy Bypass -Scope Process -Force
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                # The official Chocolatey installation command.
                $installOutput = Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1')) 2>&1
                $installOutput | Out-File -FilePath $script:logFile -Append -Encoding UTF8
                if ($LASTEXITCODE -eq 0) {
                    Write-LogAndHost "Chocolatey installed successfully." -HostColor Green
                    # Refresh environment variables to ensure 'choco' is available in the current session.
                    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
                    return $true
                } else {
                    Write-LogAndHost "Failed to install Chocolatey. Exit code: $LASTEXITCODE. Details: $($installOutput | Out-String)" -HostColor Red -LogPrefix "Install-Chocolatey"
                    return $false
                }
            } catch {
                Write-LogAndHost "Exception occurred while installing Chocolatey - $($_.Exception.Message)" -HostColor Red -LogPrefix "Install-Chocolatey"
                return $false
            }
        } else {
            Write-LogAndHost "Chocolatey installation cancelled by user." -HostColor Yellow
            return $false
        }
    } catch {
        Write-LogAndHost "Could not read user input for Chocolatey installation. $($_.Exception.Message)" -HostColor Red -LogPrefix "Install-Chocolatey"
        return $false
    }
}

# Function to perform Windows activation using an external script.
function Invoke-WindowsActivation {
    $script:activationAttempted = $true
    Write-LogAndHost "Windows activation uses an external script from 'https://get.activated.win'. Ensure you trust the source before proceeding." -HostColor Yellow -LogPrefix "Invoke-WindowsActivation"
    try {
        Write-LogAndHost "Continue with Windows activation? (Type y/n then press Enter)" -HostColor Yellow -LogPrefix "Invoke-WindowsActivation"
        $confirmActivation = Read-Host
        if ($confirmActivation.Trim().ToLower() -ne 'y') {
            Write-LogAndHost "Windows activation cancelled by user." -HostColor Yellow
            return
        }
    } catch {
        Write-LogAndHost "Could not read user input for Windows activation. $($_.Exception.Message)" -HostColor Red -LogPrefix "Invoke-WindowsActivation"
        return
    }
    Write-LogAndHost "Attempting Windows activation..." -NoHost
    try {
        # Ensure TLS 1.2 for the activation script download.
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        $activationScriptContent = Invoke-RestMethod -Uri "https://get.activated.win" -UseBasicParsing
        # Execute the script and show its output directly to the user while also logging it.
        Invoke-Expression -Command $activationScriptContent 2>&1 | Tee-Object -FilePath $script:logFile -Append | ForEach-Object { Write-Host $_ -ForegroundColor Gray }
        
        if ($LASTEXITCODE -eq 0) { 
            Write-LogAndHost "Windows activation script executed. Check console output above for status."
        }
        else {
            Write-LogAndHost "Windows activation script execution might have failed. Exit code: $LASTEXITCODE" -HostColor Yellow -LogPrefix "Invoke-WindowsActivation"
        }
    }
    catch {
        Write-LogAndHost "Exception during Windows activation - $($_.Exception.Message)" -HostColor Red -LogPrefix "Invoke-WindowsActivation"
    }
    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}

# Function to apply SpotX modifications to Spotify.
function Invoke-SpotXActivation {
    Write-LogAndHost "Attempting Spotify Activation..." -HostColor Cyan
    Write-LogAndHost "INFO: This process modifies your Spotify client. It is recommended to close Spotify before proceeding." -HostColor Yellow
    Write-LogAndHost "This script downloads and executes code from the internet (SpotX-Official GitHub). Ensure you trust the source." -HostColor Yellow -LogPrefix "Invoke-SpotXActivation"

    try {
        Write-LogAndHost "Continue with Spotify Activation? (Type y/n then press Enter)" -HostColor Yellow -LogPrefix "Invoke-SpotXActivation"
        $confirmSpotX = Read-Host
        if ($confirmSpotX.Trim().ToLower() -ne 'y') {
            Write-LogAndHost "Spotify Activation cancelled by user." -HostColor Yellow
            return
        }
    } catch {
        Write-LogAndHost "Could not read user input for Spotify Activation confirmation. $($_.Exception.Message)" -HostColor Red -LogPrefix "Invoke-SpotXActivation"
        return
    }

    $spotxParams = "-new_theme"
    $spotxUrlPrimary = 'https://raw.githubusercontent.com/SpotX-Official/spotx-official.github.io/main/run.ps1'
    $spotxUrlFallback = 'https://spotx-official.github.io/run.ps1'

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $scriptContentToExecute = $null
    $effectiveParams = $spotxParams

    try {
        Write-LogAndHost "Downloading SpotX script from primary URL: $spotxUrlPrimary" -NoHost
        $scriptContentToExecute = (Invoke-WebRequest -UseBasicParsing -Uri $spotxUrlPrimary -ErrorAction Stop).Content
        Write-LogAndHost "Successfully downloaded from primary URL." -NoHost
    } catch {
        Write-LogAndHost "Primary SpotX URL failed. Error: $($_.Exception.Message)" -HostColor DarkYellow -LogPrefix "Invoke-SpotXActivation"
        Write-LogAndHost "Attempting fallback URL: $spotxUrlFallback" -NoHost
        try {
            $scriptContentToExecute = (Invoke-WebRequest -UseBasicParsing -Uri $spotxUrlFallback -ErrorAction Stop).Content
            Write-LogAndHost "Successfully downloaded from fallback URL." -NoHost
        } catch {
            Write-LogAndHost "Fallback SpotX URL also failed. Error: $($_.Exception.Message)" -HostColor Red -LogPrefix "Invoke-SpotXActivation"
            Write-LogAndHost "Spotify Activation cannot proceed." -HostColor Red
            $_ | Out-File -FilePath $script:logFile -Append -Encoding UTF8
            Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
            $null = Read-Host
            return
        }
    }

    if ($scriptContentToExecute) {
        Write-LogAndHost "Executing SpotX script with parameters: '$effectiveParams'"
        $fullScriptToRun = "$scriptContentToExecute $effectiveParams"
        try {
            Invoke-Expression -Command $fullScriptToRun 2>&1 | Tee-Object -FilePath $script:logFile -Append | ForEach-Object { Write-Host $_ }
            Write-LogAndHost "SpotX script execution attempt finished. Check console output from SpotX above for status." -HostColor Green
        } catch {
            Write-LogAndHost "Exception occurred during SpotX script execution. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Invoke-SpotXActivation"
            $_ | Out-File -FilePath $script:logFile -Append -Encoding UTF8
        }
    } else {
        Write-LogAndHost "Failed to obtain SpotX script content." -HostColor Red -LogPrefix "Invoke-SpotXActivation"
    }

    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}

# Function to check for and install Windows updates using the PSWindowsUpdate module.
function Invoke-WindowsUpdate {
    Write-LogAndHost "Checking for Windows updates..."
    Set-ExecutionPolicy Bypass -Scope Process -Force
    try {
        # Check if the required module is available, and install it if not.
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-LogAndHost "PSWindowsUpdate module not found. Installing..."
            Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction Stop -AcceptLicense
            Write-LogAndHost "PSWindowsUpdate module installed successfully for current user."
        }
        Import-Module PSWindowsUpdate -ErrorAction Stop
        Write-LogAndHost "Checking for available updates..."
        $updates = Get-WindowsUpdate -ErrorAction Stop
        if ($updates.Count -gt 0) {
            Write-LogAndHost "Found $($updates.Count) updates. Installing..."
            Install-WindowsUpdate -AcceptAll -ErrorAction Stop
            Write-LogAndHost "Windows updates installed successfully. A manual reboot is strongly recommended to finalize the installation." -HostColor Green
        } else {
            Write-LogAndHost "No updates available."
        }
    } catch {
        Write-LogAndHost "Failed to update Windows. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Invoke-WindowsUpdate"
    }
    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}

# Function to disable common Windows Telemetry services and registry keys.
function Invoke-DisableTelemetry {
    Write-LogAndHost "Checking Windows Telemetry status..." -HostColor Cyan
    
    $telemetryService = Get-Service -Name "DiagTrack" -ErrorAction SilentlyContinue
    $telemetryRegValue = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -ErrorAction SilentlyContinue

    # Check if telemetry already appears to be disabled to avoid unnecessary work.
    if ($telemetryService -and $telemetryService.StartType -eq 'Disabled' -and $telemetryRegValue -and $telemetryRegValue.AllowTelemetry -eq 0) {
        Write-LogAndHost "Windows Telemetry appears to be already disabled." -HostColor Green
        Write-LogAndHost "No changes were made."
        Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
        $null = Read-Host
        return
    }
    
    try {
        Write-LogAndHost "Telemetry is currently enabled. Continue with disabling? (Type y/n then press Enter)" -HostColor Yellow -LogPrefix "Invoke-DisableTelemetry"
        $confirmTelemetry = Read-Host
        if ($confirmTelemetry.Trim().ToLower() -ne 'y') {
            Write-LogAndHost "Telemetry disabling cancelled by user." -HostColor Yellow
            return
        }
    } catch {
        Write-LogAndHost "Could not read user input for Telemetry confirmation. $($_.Exception.Message)" -HostColor Red -LogPrefix "Invoke-DisableTelemetry"
        return
    }

    Write-LogAndHost "Applying telemetry settings..." -NoHost
    
    try {
        # List of services related to telemetry.
        $servicesToDisable = @("DiagTrack", "dmwappushservice")
        foreach ($serviceName in $servicesToDisable) {
            $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
            if ($service) {
                try {
                    Write-LogAndHost "Stopping service: $serviceName..." -NoLog
                    Stop-Service -Name $serviceName -Force -ErrorAction Stop
                    Write-LogAndHost "Disabling service: $serviceName..." -NoLog
                    Set-Service -Name $serviceName -StartupType Disabled -ErrorAction Stop
                    Write-LogAndHost "$serviceName service stopped and disabled."
                } catch {
                    Write-LogAndHost "Could not stop or disable service '$serviceName'. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Invoke-DisableTelemetry"
                }
            } else {
                Write-LogAndHost "Service '$serviceName' not found, skipping." -HostColor DarkGray -NoLog
            }
        }

        Write-LogAndHost "Configuring registry keys..." -NoLog
        
        # A collection of registry keys to disable telemetry and related features.
        $regKeys = @{
            "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" = @{ Name = "AllowTelemetry"; Value = 0; Type = "DWord" };
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection" = @{ Name = "AllowTelemetry"; Value = 0; Type = "DWord" };
            "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" = @{ Name = "DisableWindowsConsumerFeatures"; Value = 1; Type = "DWord" };
            "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo" = @{ Name = "DisabledByGroupPolicy"; Value = 1; Type = "DWord" };
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Privacy" = @{ Name = "TailoredExperiencesWithDiagnosticDataEnabled"; Value = 0; Type = "DWord" };
        }

        foreach ($path in $regKeys.Keys) {
            $keyInfo = $regKeys[$path]
            try {
                # Create the registry path if it doesn't exist.
                if (-not (Test-Path $path)) {
                    Write-LogAndHost "Creating registry path: $path" -NoLog
                    New-Item -Path $path -Force -ErrorAction Stop | Out-Null
                }
                Set-ItemProperty -Path $path -Name $keyInfo.Name -Value $keyInfo.Value -Type $keyInfo.Type -Force -ErrorAction Stop
                Write-LogAndHost "Successfully set registry value '$($keyInfo.Name)' at '$path'." -NoHost
            } catch {
                Write-LogAndHost "Failed to set registry key at '$path'. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Invoke-DisableTelemetry"
            }
        }
        
        Write-LogAndHost "Telemetry has been successfully disabled." -HostColor Green
    } catch {
        Write-LogAndHost "An unexpected error occurred while disabling telemetry. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Invoke-DisableTelemetry"
    }

    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}

# ENHANCED FUNCTION (v1.5): Displays key system information, including Video Card.
function Show-SystemInfo {
    Write-LogAndHost "Gathering system information..." -HostColor Cyan -LogPrefix "Show-SystemInfo"
    $infoOutput = New-Object System.Collections.Generic.List[string]
    $line = "-" * 60
    $infoOutput.Add($line)
    $infoOutput.Add(" System Information")
    $infoOutput.Add($line)

    try {
        # OS Info
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
        $infoOutput.Add(" OS Name: $($osInfo.Caption)")
        $infoOutput.Add(" OS Version: $($osInfo.Version)")
        
        # CPU Info
        $cpuInfo = Get-CimInstance -ClassName Win32_Processor
        $infoOutput.Add(" Processor: $($cpuInfo.Name.Trim())")
        
        # RAM Info
        $ramInfo = Get-CimInstance -ClassName Win32_ComputerSystem
        $ramGB = [math]::Round($ramInfo.TotalPhysicalMemory / 1GB, 2)
        $infoOutput.Add(" Installed RAM: $($ramGB) GB")
        
        $infoOutput.Add($line)
        $infoOutput.Add(" Network Information")
        $infoOutput.Add($line)
        
        # Network Info
        $netAdapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled }
        if ($netAdapters) {
            foreach ($adapter in $netAdapters) {
                $infoOutput.Add(" Description: $($adapter.Description)")
                $infoOutput.Add("   IP Address: $($adapter.IPAddress -join ', ')")
                $infoOutput.Add("   MAC Address: $($adapter.MACAddress)")
                $infoOutput.Add("")
            }
        } else {
            $infoOutput.Add(" No active network adapters with an IP address found.")
        }

        $infoOutput.Add($line)
        $infoOutput.Add(" Disk Information")
        $infoOutput.Add($line)

        # Disk Info
        $disks = Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
        if ($disks) {
             foreach ($disk in $disks) {
                $sizeGB = [math]::Round($disk.Size / 1GB, 2)
                $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
                $percentFree = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 2)
                $infoOutput.Add(" Drive: $($disk.DeviceID)")
                $infoOutput.Add("   Size: $($sizeGB) GB")
                $infoOutput.Add("   Free Space: $($freeGB) GB ($($percentFree)%)")
                $infoOutput.Add("")
            }
        } else {
             $infoOutput.Add(" No fixed logical disks found.")
        }
        
        $infoOutput.Add($line)
        $infoOutput.Add(" Video Card Information")
        $infoOutput.Add($line)

        # Video Card Info
        $videoControllers = Get-CimInstance -ClassName Win32_VideoController
        if ($videoControllers) {
            foreach ($video in $videoControllers) {
                $infoOutput.Add(" Name: $($video.Name)")
                if ($video.AdapterRAM) {
                    $adapterRamMB = [math]::Round($video.AdapterRAM / 1MB)
                    $infoOutput.Add("   Adapter RAM: $($adapterRamMB) MB")
                }
                $infoOutput.Add("   Driver Version: $($video.DriverVersion)")
                $infoOutput.Add("")
            }
        } else {
            $infoOutput.Add(" No video controllers found.")
        }
    }
    catch {
        Write-LogAndHost "Failed to gather some system information. Error: $($_.Exception.Message)" -HostColor Red -LogPrefix "Show-SystemInfo"
        $infoOutput.Add("ERROR: Could not retrieve all information.")
    }
    
    # Display the collected information
    Clear-Host
    foreach($entry in $infoOutput) {
        Write-Host $entry
    }
    
    Write-LogAndHost "System information displayed." -NoHost
    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}

# NEW FUNCTION (v1.5): Imports a list of programs from a file and installs them.
function Import-ProgramSelection {
    if (-not $script:guiAvailable) {
        Write-LogAndHost "GUI is not available, cannot launch the Program Import tool." -HostColor Red -LogPrefix "Import-ProgramSelection"
        Start-Sleep -Seconds 2
        return
    }

    Write-LogAndHost "Launching Program Import..." -HostColor Cyan -LogPrefix "Import-ProgramSelection"

    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select Program List to Import"
    $openFileDialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

    if ($openFileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-LogAndHost "File import cancelled by user." -HostColor Yellow
        return
    }
    
    $filePath = $openFileDialog.FileName
    Write-LogAndHost "User selected file to import: '$filePath'." -NoHost
    
    if (-not (Test-Path $filePath)) {
        Write-LogAndHost "The selected file does not exist: '$filePath'." -HostColor Red
        Start-Sleep -Seconds 2
        return
    }

    $programsToInstall = @()
    try {
        # Read the file and convert it from JSON into a PowerShell object.
        $programsToInstall = Get-Content -Path $filePath -Raw | ConvertFrom-Json -ErrorAction Stop
        # Basic validation to ensure the JSON is an array.
        if ($null -eq $programsToInstall -or $programsToInstall.GetType().Name -ne 'Object[]') {
            throw "File does not contain a valid JSON array of program names."
        }
    } catch {
        Write-LogAndHost "Failed to read or parse the program list file. Make sure it's a valid JSON file containing an array of strings. Error: $($_.Exception.Message)" -HostColor Red -LogPrefix "Import-ProgramSelection"
        Start-Sleep -Seconds 3
        return
    }

    if ($programsToInstall.Count -eq 0) {
        Write-LogAndHost "The imported file contains no programs to install." -HostColor Yellow
        Start-Sleep -Seconds 2
        return
    }
    
    Clear-Host
    Write-LogAndHost "The following programs will be installed from the file:" -HostColor Cyan
    $programsToInstall | ForEach-Object { Write-Host " - $_" }
    Write-Host ""
    
    try {
        Write-LogAndHost "Continue with installation? (Type y/n then press Enter)" -HostColor Yellow
        $confirmInstall = Read-Host
    } catch {
        Write-LogAndHost "Could not read user input. Aborting. $($_.Exception.Message)" -HostColor Red -LogPrefix "Import-ProgramSelection"
        Start-Sleep -Seconds 2
        return
    }

    if ($confirmInstall.Trim().ToLower() -eq 'y') {
        Clear-Host
        Write-LogAndHost "Starting installation from imported file..." -NoHost
        if (Install-Programs -ProgramsToInstall $programsToInstall) {
            Write-LogAndHost "Installation from file completed." -HostColor Green
        } else {
            Write-LogAndHost "Some programs from the imported list may have failed to install. Check the log for details." -HostColor Yellow
        }
    } else {
        Write-LogAndHost "Installation from file cancelled by user." -HostColor Yellow
    }

    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}

# FUNCTION: Clean temporary system files with a GUI.
function Invoke-TempFileCleanup {
    if (-not $script:guiAvailable) {
        Write-LogAndHost "GUI is not available, cannot launch the System Cleanup tool." -HostColor Red -LogPrefix "Invoke-TempFileCleanup"
        Start-Sleep -Seconds 2
        return
    }
    Write-LogAndHost "Launching System Cleanup GUI..." -HostColor Cyan

    # --- GUI Setup ---
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Perdanga System Cleanup"
    $form.Size = New-Object System.Drawing.Size(600, 550)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)

    # Common Font and Color scheme for a modern look.
    $commonFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $labelColor = [System.Drawing.Color]::White
    $controlBackColor = [System.Drawing.Color]::FromArgb(60, 60, 63)
    $controlForeColor = [System.Drawing.Color]::White
    $groupboxForeColor = [System.Drawing.Color]::Gainsboro

    # Helper functions for creating styled controls to reduce code repetition.
    function New-StyledCheckBox($Text, $Location, $Checked) {
        $checkbox = New-Object System.Windows.Forms.CheckBox; $checkbox.Text = $Text; $checkbox.Location = $Location; $checkbox.Font = $commonFont; $checkbox.ForeColor = $labelColor; $checkbox.AutoSize = $true; $checkbox.Checked = $Checked; return $checkbox
    }
    function New-StyledGroupBox($Text, $Location, $Size) {
        $groupbox = New-Object System.Windows.Forms.GroupBox; $groupbox.Text = $Text; $groupbox.Location = $Location; $groupbox.Size = $Size; $groupbox.Font = $commonFont; $groupbox.ForeColor = $groupboxForeColor; return $groupbox
    }

    # --- GroupBox for Cleanup Options ---
    $groupOptions = New-StyledGroupBox "Select items to clean" "15,15" "560,200"
    $form.Controls.Add($groupOptions) | Out-Null
    
    $yPos = 30
    $checkWinTemp = New-StyledCheckBox "Windows Temporary Files" "20,$yPos" $true; $groupOptions.Controls.Add($checkWinTemp) | Out-Null; $yPos += 30
    $checkNvidia = New-StyledCheckBox "NVIDIA Cache Files" "20,$yPos" $true; $groupOptions.Controls.Add($checkNvidia) | Out-Null; $yPos += 30
    $checkWinUpdate = New-StyledCheckBox "Windows Update Cache" "20,$yPos" $true; $groupOptions.Controls.Add($checkWinUpdate) | Out-Null; $yPos += 30
    $checkPrefetch = New-StyledCheckBox "Windows Prefetch Files" "20,$yPos" $true; $groupOptions.Controls.Add($checkPrefetch) | Out-Null; $yPos += 30
    $checkRecycleBin = New-StyledCheckBox "Empty Recycle Bin" "20,$yPos" $true; $groupOptions.Controls.Add($checkRecycleBin) | Out-Null; $yPos += 30
    
    # --- GroupBox for Log Output ---
    $groupLog = New-StyledGroupBox "Log" "15,225" "560,200"
    $form.Controls.Add($groupLog) | Out-Null
    
    $logBox = New-Object System.Windows.Forms.RichTextBox
    $logBox.Location = "15,25"
    $logBox.Size = "530,160"
    $logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $logBox.BackColor = $controlBackColor
    $logBox.ForeColor = $controlForeColor
    $logBox.ReadOnly = $true
    $logBox.BorderStyle = "FixedSingle"
    $logBox.ScrollBars = "Vertical"
    $groupLog.Controls.Add($logBox) | Out-Null

    # --- Buttons ---
    $buttonAnalyze = New-Object System.Windows.Forms.Button; $buttonAnalyze.Text = "Analyze"; $buttonAnalyze.Size = "120,30"; $buttonAnalyze.Location = "100,450"; $buttonAnalyze.Font = $commonFont; $buttonAnalyze.ForeColor = [System.Drawing.Color]::White; $buttonAnalyze.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180); $buttonAnalyze.FlatStyle = "Flat"; $buttonAnalyze.FlatAppearance.BorderSize = 0;
    $buttonClean = New-Object System.Windows.Forms.Button; $buttonClean.Text = "Clean"; $buttonClean.Size = "120,30"; $buttonClean.Location = "235,450"; $buttonClean.Font = $commonFont; $buttonClean.ForeColor = [System.Drawing.Color]::White; $buttonClean.BackColor = [System.Drawing.Color]::FromArgb(200, 70, 70); $buttonClean.FlatStyle = "Flat"; $buttonClean.FlatAppearance.BorderSize = 0;
    $buttonClose = New-Object System.Windows.Forms.Button; $buttonClose.Text = "Exit"; $buttonClose.Size = "120,30"; $buttonClose.Location = "370,450"; $buttonClose.Font = $commonFont; $buttonClose.ForeColor = [System.Drawing.Color]::White; $buttonClose.BackColor = [System.Drawing.Color]::FromArgb(90, 90, 90); $buttonClose.FlatStyle = "Flat"; $buttonClose.FlatAppearance.BorderSize = 0;
    
    $form.Controls.Add($buttonAnalyze) | Out-Null
    $form.Controls.Add($buttonClean) | Out-Null
    $form.Controls.Add($buttonClose) | Out-Null

    $buttonClose.add_Click({$form.Close()}) | Out-Null

    # --- Logic ---
    $totalSize = 0
    $pathsToClean = @{}

    # Helper scriptblock to add text to the log box with color.
    $logWriter = {
        param($Message, $Color = 'White')
        $logBox.SelectionStart = $logBox.TextLength
        $logBox.SelectionLength = 0
        $logBox.SelectionColor = $Color
        $logBox.AppendText("$(Get-Date -Format 'HH:mm:ss') - $Message`n")
        $logBox.ScrollToCaret()
    }
    
    # Analyze button logic: calculates the size of deletable files.
    $buttonAnalyze.add_Click({
        $logBox.Clear()
        & $logWriter "Starting analysis..." 'Cyan'
        $totalSize = 0
        $pathsToClean.Clear()

        $calculateSize = {
            param($path)
            $size = 0
            try {
                if (Test-Path $path -ErrorAction Stop) {
                    $items = Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue
                    $size = ($items | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                }
            } catch {
                & $logWriter "Could not access path: $path. Error: $($_.Exception.Message)" 'Yellow'
            }
            return $size
        }

        if ($checkWinTemp.Checked) {
            $paths = @("$env:TEMP", "$env:windir\Temp")
            $pathsToClean['Windows Temp'] = $paths
            foreach($p in $paths) { $totalSize += & $calculateSize $p }
        }
        if ($checkNvidia.Checked) {
            $paths = @("$env:LOCALAPPDATA\NVIDIA\GLCache", "$env:ProgramData\NVIDIA Corporation\Downloader")
            $pathsToClean['NVIDIA Cache'] = $paths
            foreach($p in $paths) { $totalSize += & $calculateSize $p }
        }
        if ($checkWinUpdate.Checked) {
            $paths = @("$env:windir\SoftwareDistribution\Download")
            $pathsToClean['Windows Update Cache'] = $paths
            foreach($p in $paths) { $totalSize += & $calculateSize $p }
        }
        if ($checkPrefetch.Checked) {
            $paths = @("$env:windir\Prefetch")
            $pathsToClean['Prefetch'] = $paths
            $totalSize += & $calculateSize $paths[0]
        }
        if ($checkRecycleBin.Checked) {
            try {
                # Use the Shell.Application COM object to query the Recycle Bin size.
                $shell = New-Object -ComObject Shell.Application
                $recycleBin = $shell.NameSpace(0xA)
                $items = $recycleBin.Items()
                $size = ($items | ForEach-Object { $_.Size } | Measure-Object -Sum).Sum
                if ($size -gt 0) {
                    $pathsToClean['Recycle Bin'] = $recycleBin # Store the object itself for later.
                    $totalSize += $size
                }
            } catch {
                & $logWriter "Could not access Recycle Bin. Error: $($_.Exception.Message)" 'Yellow'
            }
        }
        
        $sizeInMB = [math]::Round($totalSize / 1MB, 2)
        & $logWriter "Analysis complete. Found $sizeInMB MB of files to clean." 'Green'
        & $logWriter "Press 'Clean' to remove these files." 'Cyan'
    })

    # Clean button logic: performs the actual deletion after confirmation.
    $buttonClean.add_Click({
        if ($pathsToClean.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please run an analysis first by clicking the 'Analyze' button.", "Analysis Required", "OK", "Information") | Out-Null
            return
        }
        
        $confirmResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to permanently delete these files? This cannot be undone.", "Confirm Deletion", "YesNo", "Warning")
        if ($confirmResult -ne 'Yes') {
            & $logWriter "Cleanup cancelled by user." 'Yellow'
            return
        }

        & $logWriter "Starting cleanup..." 'Cyan'
        $buttonClean.Enabled = $false
        $buttonAnalyze.Enabled = $false
        $form.Update()

        $totalDeleted = 0
        foreach($category in $pathsToClean.Keys) {
            & $logWriter "Cleaning $category..." 'White'
            $paths = $pathsToClean[$category]
            
            if ($category -eq 'Recycle Bin') {
                try {
                    $recycleBinObject = $pathsToClean[$category]
                    $sizeToDelete = ($recycleBinObject.Items() | ForEach-Object { $_.Size } | Measure-Object -Sum).Sum
                    $totalDeleted += $sizeToDelete
                    # Use the built-in cmdlet to empty the recycle bin.
                    Clear-RecycleBin -Force -ErrorAction Stop
                    & $logWriter "Recycle Bin emptied successfully." 'Green'
                } catch {
                    & $logWriter "Failed to empty Recycle Bin. Error: $($_.Exception.Message)" 'Red'
                }
            } else {
                foreach($path in $paths) {
                    try {
                        $items = Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue
                        $size = ($items | Measure-Object -Property Length -Sum).Sum
                        $totalDeleted += $size
                        Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
                        & $logWriter "Cleaned path: $path" 'Green'
                    } catch {
                        & $logWriter "Failed to clean path: $path. Error: $($_.Exception.Message)" 'Red'
                    }
                }
            }
        }
        
        $deletedInMB = [math]::Round($totalDeleted / 1MB, 2)
        & $logWriter "Cleanup complete. Freed approximately $deletedInMB MB of space." 'Green'
        $pathsToClean.Clear()
        $buttonClean.Enabled = $true
        $buttonAnalyze.Enabled = $true
    })

    # Show the form and dispose of it when closed.
    try {
        $null = $form.ShowDialog()
    }
    catch {
        Write-LogAndHost "An unexpected error occurred with the System Cleanup GUI. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Invoke-TempFileCleanup"
    }
    finally {
        $form.Dispose()
    }
    
    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}

# ENHANCED FUNCTION: Create a detailed autounattend.xml file via GUI with regional settings and tooltips.
function Create-UnattendXml {
    if (-not $script:guiAvailable) {
        Write-LogAndHost "GUI is not available, cannot launch the Unattend XML Creator." -HostColor Red -LogPrefix "Create-UnattendXml"
        Start-Sleep -Seconds 2
        return
    }
    Write-LogAndHost "Launching Unattend XML Creator GUI..." -HostColor Cyan
    
    # --- State management for keyboard layout selection ---
    $checkedKeyboardLayoutNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    
    # --- State management for Time Zone selection ---
    $timeZoneMap = @{}

    # --- GUI Setup ---
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Perdanga Unattend.xml Creator"
    $form.Size = New-Object System.Drawing.Size(800, 750)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)

    # --- ToolTip Setup (for descriptions) ---
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.AutoPopDelay = 10000 # Keep tooltip visible for 10 seconds.
    $toolTip.InitialDelay = 500   # Show after 0.5 seconds.
    $toolTip.ReshowDelay = 500

    # --- Button Panel ---
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Height = 50
    $buttonPanel.Dock = "Bottom"
    $buttonPanel.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 63)
    $form.Controls.Add($buttonPanel) | Out-Null

    # --- Tab Control ---
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = "Fill"
    $tabControl.Padding = New-Object System.Drawing.Point(10, 5)
    $form.Controls.Add($tabControl) | Out-Null
    $tabControl.BringToFront()

    # Common Font and Color scheme.
    $commonFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $labelColor = [System.Drawing.Color]::White
    $controlBackColor = [System.Drawing.Color]::FromArgb(60, 60, 63)
    $controlForeColor = [System.Drawing.Color]::White
    $groupboxForeColor = [System.Drawing.Color]::Gainsboro

    # Helper functions for creating styled controls.
    function New-StyledLabel($Text, $Location) {
        $label = New-Object System.Windows.Forms.Label; $label.Text = $Text; $label.Location = $Location; $label.Font = $commonFont; $label.ForeColor = $labelColor; $label.AutoSize = $true; return $label
    }
    function New-StyledTextBox($Location, $Size) {
        $textbox = New-Object System.Windows.Forms.TextBox; $textbox.Location = $Location; $textbox.Size = $Size; $textbox.Font = $commonFont; $textbox.BackColor = $controlBackColor; $textbox.ForeColor = $controlForeColor; $textbox.BorderStyle = "FixedSingle"; return $textbox
    }
    function New-StyledComboBox($Location, $Size) {
        $combobox = New-Object System.Windows.Forms.ComboBox; $combobox.Location = $Location; $combobox.Size = $Size; $combobox.Font = $commonFont; $combobox.BackColor = $controlBackColor; $combobox.ForeColor = $controlForeColor; $combobox.FlatStyle = "Flat"; return $combobox
    }
    function New-StyledCheckBox($Text, $Location, $Checked) {
        $checkbox = New-Object System.Windows.Forms.CheckBox; $checkbox.Text = $Text; $checkbox.Location = $Location; $checkbox.Font = $commonFont; $checkbox.ForeColor = $labelColor; $checkbox.AutoSize = $true; $checkbox.Checked = $Checked; return $checkbox
    }
    function New-StyledGroupBox($Text, $Location, $Size) {
        $groupbox = New-Object System.Windows.Forms.GroupBox; $groupbox.Text = $Text; $groupbox.Location = $Location; $groupbox.Size = $Size; $groupbox.Font = $commonFont; $groupbox.ForeColor = $groupboxForeColor; return $groupbox
    }
    
    # --- Tab 1: General Settings ---
    $tabGeneral = New-Object System.Windows.Forms.TabPage; $tabGeneral.Text = "General"; $tabGeneral.BackColor = $form.BackColor; $tabControl.Controls.Add($tabGeneral) | Out-Null
    $yPos = 30
    $tabGeneral.Controls.Add((New-StyledLabel -Text "Computer Name:" -Location "20,$yPos")) | Out-Null; $textComputerName = New-StyledTextBox -Location "180,$yPos" -Size "280,20"; $textComputerName.Text = "DESKTOP-PC"; $tabGeneral.Controls.Add($textComputerName) | Out-Null; $yPos += 40
    $tabGeneral.Controls.Add((New-StyledLabel -Text "Admin User Name:" -Location "20,$yPos")) | Out-Null; $textUserName = New-StyledTextBox -Location "180,$yPos" -Size "280,20"; $textUserName.Text = "Admin"; $tabGeneral.Controls.Add($textUserName) | Out-Null; $yPos += 40
    
    $tabGeneral.Controls.Add((New-StyledLabel -Text "Password:" -Location "20,$yPos")) | Out-Null
    $textPassword = New-StyledTextBox -Location "180,$yPos" -Size "280,20"
    $textPassword.UseSystemPasswordChar = $true
    $textPassword.MaxLength = 127
    $tabGeneral.Controls.Add($textPassword) | Out-Null
    $labelPasswordCounter = New-StyledLabel -Text "0/127" -Location "470,$yPos"
    $tabGeneral.Controls.Add($labelPasswordCounter) | Out-Null
    $textPassword.Add_TextChanged({
        $length = $textPassword.Text.Length
        $labelPasswordCounter.Text = "$length/127"
        if ($length -eq 127) {
            $labelPasswordCounter.ForeColor = [System.Drawing.Color]::Crimson
        } else {
            $labelPasswordCounter.ForeColor = $labelColor
        }
    })
    $yPos += 40

    # --- Tab 2: Regional Settings ---
    $tabRegional = New-Object System.Windows.Forms.TabPage; $tabRegional.Text = "Regional"; $tabRegional.BackColor = $form.BackColor; $tabRegional.Padding = New-Object System.Windows.Forms.Padding(10)
    $tabControl.Controls.Add($tabRegional) | Out-Null

    $groupLocale = New-StyledGroupBox "Language & Locale" "15,15" "750,150"
    $tabRegional.Controls.Add($groupLocale) | Out-Null
    $yPos = 30
    $commonLocales = @("ar-SA", "cs-CZ", "da-DK", "de-DE", "el-GR", "en-GB", "en-US", "es-ES", "es-MX", "fi-FI", "fr-CA", "fr-FR", "he-IL", "hu-HU", "it-IT", "ja-JP", "ko-KR", "nb-NO", "nl-NL", "pl-PL", "pt-BR", "pt-PT", "ro-RO", "ru-RU", "sk-SK", "sv-SE", "th-TH", "tr-TR", "zh-CN", "zh-TW")
    
    $groupLocale.Controls.Add((New-StyledLabel -Text "UI Language:" -Location "15,$yPos")) | Out-Null; $comboUiLanguage = New-StyledComboBox -Location "150,$yPos" -Size "250,20"; $comboUiLanguage.Items.AddRange($commonLocales) | Out-Null; $comboUiLanguage.Text = (Get-UICulture).Name; $groupLocale.Controls.Add($comboUiLanguage) | Out-Null
    $groupLocale.Controls.Add((New-StyledLabel -Text "(e.g., en-US, de-DE)" -Location "410,$yPos")) | Out-Null
    $yPos += 40
    $groupLocale.Controls.Add((New-StyledLabel -Text "System Locale:" -Location "15,$yPos")) | Out-Null; $comboSystemLocale = New-StyledComboBox -Location "150,$yPos" -Size "250,20"; $comboSystemLocale.Items.AddRange($commonLocales) | Out-Null; $comboSystemLocale.Text = (Get-Culture).Name; $groupLocale.Controls.Add($comboSystemLocale) | Out-Null
    $groupLocale.Controls.Add((New-StyledLabel -Text "(e.g., en-US, ja-JP)" -Location "410,$yPos")) | Out-Null
    $yPos += 40
    $groupLocale.Controls.Add((New-StyledLabel -Text "User Locale:" -Location "15,$yPos")) | Out-Null; $comboUserLocale = New-StyledComboBox -Location "150,$yPos" -Size "250,20"; $comboUserLocale.Items.AddRange($commonLocales) | Out-Null; $comboUserLocale.Text = (Get-Culture).Name; $groupLocale.Controls.Add($comboUserLocale) | Out-Null
    $groupLocale.Controls.Add((New-StyledLabel -Text "(e.g., en-US, tr-TR)" -Location "410,$yPos")) | Out-Null

    $groupTimeZone = New-StyledGroupBox "Time Zone" "15,180" "750,220"
    $tabRegional.Controls.Add($groupTimeZone) | Out-Null
    $yPos = 30
    $groupTimeZone.Controls.Add((New-StyledLabel -Text "Search:" -Location "15,$yPos")) | Out-Null; $textTimeZoneSearch = New-StyledTextBox -Location "85,$yPos" -Size "645,20"; $groupTimeZone.Controls.Add($textTimeZoneSearch) | Out-Null; $yPos += 35
    $listTimeZone = New-Object System.Windows.Forms.ListBox; $listTimeZone.Location = "15,$yPos"; $listTimeZone.Size = "715,100"; $listTimeZone.Font = $commonFont; $listTimeZone.BackColor = $controlBackColor; $listTimeZone.ForeColor = $controlForeColor
    
    # A static list of modern Windows Time Zone IDs for reliability.
    $windows11TimeZoneIds = @(
        "Dateline Standard Time", "UTC-11", "Aleutian Standard Time", "Hawaiian Standard Time", 
        "Marquesas Standard Time", "Alaskan Standard Time", "UTC-09", "Pacific Standard Time (Mexico)", 
        "UTC-08", "Pacific Standard Time", "US Mountain Standard Time", "Mountain Standard Time (Mexico)", 
        "Mountain Standard Time", "Yukon Standard Time", "Central America Standard Time", "Central Standard Time", 
        "Easter Island Standard Time", "Central Standard Time (Mexico)", "Canada Central Standard Time", 
        "SA Pacific Standard Time", "Eastern Standard Time (Mexico)", "Eastern Standard Time", "Haiti Standard Time", 
        "Cuba Standard Time", "US Eastern Standard Time", "Turks And Caicos Standard Time", "Paraguay Standard Time", 
        "Atlantic Standard Time", "Venezuela Standard Time", "Central Brazilian Standard Time", "SA Western Standard Time", 
        "Pacific SA Standard Time", "Newfoundland Standard Time", "Tocantins Standard Time", "E. South America Standard Time", 
        "SA Eastern Standard Time", "Argentina Standard Time", "Greenland Standard Time", "Montevideo Standard Time", 
        "Magallanes Standard Time", "Saint Pierre Standard Time", "Bahia Standard Time", "UTC-02", 
        "Azores Standard Time", "Cape Verde Standard Time", "UTC", "GMT Standard Time", "Greenwich Standard Time", 
        "W. Europe Standard Time", "Central Europe Standard Time", "Romance Standard Time", 
        "Central European Standard Time", "W. Central Africa Standard Time", "Jordan Standard Time", 
        "GTB Standard Time", "Middle East Standard Time", "Egypt Standard Time", "E. Europe Standard Time", 
        "Syria Standard Time", "West Bank Standard Time", "South Africa Standard Time", "FLE Standard Time", 
        "Israel Standard Time", "Kaliningrad Standard Time", "Sudan Standard Time", "Libya Standard Time", 
        "Namibia Standard Time", "Arabic Standard Time", "Turkey Standard Time", "Arab Standard Time", 
        "Belarus Standard Time", "Russian Standard Time", "E. Africa Standard Time", "Iran Standard Time", 
        "Arabian Standard Time", "Astrakhan Standard Time", "Azerbaijan Standard Time", "Russia Time Zone 3", 
        "Mauritius Standard Time", "Saratov Standard Time", "Georgian Standard Time", "Volgograd Standard Time", 
        "Caucasus Standard Time", "Afghanistan Standard Time", "West Asia Standard Time", "Ekaterinburg Standard Time", 
        "Pakistan Standard Time", "Qyzylorda Standard Time", "India Standard Time", "Sri Lanka Standard Time", 
        "Nepal Standard Time", "Central Asia Standard Time", "Bangladesh Standard Time", "Omsk Standard Time", 
        "Myanmar Standard Time", "SE Asia Standard Time", "Altai Standard Time", "W. Mongolia Standard Time", 
        "North Asia Standard Time", "N. Central Asia Standard Time", "Tomsk Standard Time", "China Standard Time", 
        "North Asia East Standard Time", "Singapore Standard Time", "W. Australia Standard Time", "Taipei Standard Time", 
        "Ulaanbaatar Standard Time", "Aus Central W. Standard Time", "Transbaikal Standard Time", "Tokyo Standard Time", 
        "North Korea Standard Time", "Korea Standard Time", "Yakutsk Standard Time", "Cen. Australia Standard Time", 
        "AUS Central Standard Time", "E. Australia Standard Time", "AUS Eastern Standard Time", "West Pacific Standard Time", 
        "Tasmania Standard Time", "Vladivostok Standard Time", "Lord Howe Standard Time", "Bougainville Standard Time", 
        "Russia Time Zone 10", "Magadan Standard Time", "Norfolk Standard Time", "Sakhalin Standard Time", 
        "Central Pacific Standard Time", "Russia Time Zone 11", "New Zealand Standard Time", "UTC+12", 
        "Fiji Standard Time", "Chatham Islands Standard Time", "UTC+13", "Tonga Standard Time", 
        "Samoa Standard Time", "Line Islands Standard Time"
    )

    $allTimeZonesInfo = try { 
        $windows11TimeZoneIds | ForEach-Object { [System.TimeZoneInfo]::FindSystemTimeZoneById($_) }
    } catch { 
        Write-LogAndHost "Could not find all static time zones. The list may be incomplete. Falling back to system's available time zones." -HostColor Yellow -LogPrefix "Create-UnattendXml"
        [System.TimeZoneInfo]::GetSystemTimeZones() 
    }
    
    $formattedTimeZones = foreach ($tz in $allTimeZonesInfo) {
        $offset = $tz.BaseUtcOffset
        $offsetSign = if ($offset.Ticks -ge 0) { "+" } else { "-" }
        $offsetString = "{0:hh\:mm}" -f $offset
        $displayString = "(UTC{0}{1}) {2}" -f $offsetSign, $offsetString, $tz.Id
        $timeZoneMap[$displayString] = $tz.Id
        $displayString
    }
    $sortedFormattedTimeZones = $formattedTimezones | Sort-Object
    
    if ($null -ne $sortedFormattedTimeZones) { $listTimeZone.Items.AddRange($sortedFormattedTimeZones) | Out-Null }
    
    # Pre-select the user's current time zone.
    try {
        $currentTimeZoneId = (Get-TimeZone).Id
        $currentFormattedTz = $timeZoneMap.GetEnumerator() | Where-Object { $_.Value -eq $currentTimeZoneId } | Select-Object -First 1 -ExpandProperty Key
        if ($currentFormattedTz) { $listTimeZone.SelectedItem = $currentFormattedTz }
    } catch {}

    $groupTimeZone.Controls.Add($listTimeZone) | Out-Null; $yPos += $listTimeZone.Height + 10

    $groupTimeZone.Controls.Add((New-StyledLabel -Text "Current Selection:" -Location "15,$yPos")) | Out-Null
    $labelSelectedTimeZone = New-StyledLabel -Text "None" -Location "150,$yPos"; $labelSelectedTimeZone.ForeColor = [System.Drawing.Color]::LightSteelBlue; $labelSelectedTimeZone.AutoSize = $false; $labelSelectedTimeZone.Size = '580,20'
    $groupTimeZone.Controls.Add($labelSelectedTimeZone) | Out-Null
    
    $listTimeZone.Add_SelectedIndexChanged({
        if ($listTimeZone.SelectedItem) { $labelSelectedTimeZone.Text = $listTimeZone.SelectedItem } else { $labelSelectedTimeZone.Text = "None" }
    }) | Out-Null
    $textTimeZoneSearch.Add_TextChanged({
        $selected = $listTimeZone.SelectedItem
        $listTimeZone.BeginUpdate()
        $listTimeZone.Items.Clear()
        $searchText = $textTimeZoneSearch.Text
        $filteredTimeZones = $sortedFormattedTimeZones | Where-Object { $_ -match [regex]::Escape($searchText) }
        if ($null -ne $filteredTimeZones) { $listTimeZone.Items.AddRange($filteredTimeZones) | Out-Null }
        if ($selected -and $listTimeZone.Items.Contains($selected)) { $listTimeZone.SelectedItem = $selected } elseif ($listTimeZone.Items.Count -gt 0) { $listTimeZone.SelectedIndex = 0 }
        $listTimeZone.EndUpdate()
    }) | Out-Null
    if ($listTimeZone.SelectedItem) { $labelSelectedTimeZone.Text = $listTimeZone.SelectedItem }

    $groupKeyboard = New-StyledGroupBox "Keyboard Layouts (select up to 5)" "15,415" "750,245"
    $tabRegional.Controls.Add($groupKeyboard) | Out-Null
    $yPos = 30
    $groupKeyboard.Controls.Add((New-StyledLabel -Text "Search:" -Location "15,$yPos")) | Out-Null; $textKeyboardSearch = New-StyledTextBox -Location "85,$yPos" -Size "645,20"; $groupKeyboard.Controls.Add($textKeyboardSearch) | Out-Null; $yPos += 35
    $listKeyboardLayouts = New-Object System.Windows.Forms.CheckedListBox; $listKeyboardLayouts.Location = "15,$yPos"; $listKeyboardLayouts.Size = "715,110"; $listKeyboardLayouts.Font = $commonFont; $listKeyboardLayouts.BackColor = $controlBackColor; $listKeyboardLayouts.ForeColor = $controlForeColor; $listKeyboardLayouts.CheckOnClick = $true
    # A map of friendly names to the required unattend.xml format for keyboard layouts.
    $keyboardLayoutData = @{
        "Arabic (101)"="0401:00000401"; "Bulgarian"="0402:00000402"; "Chinese (Traditional) - US Keyboard"="0404:00000404"; "Czech"="0405:00000405"; "Danish"="0406:00000406"; "German"="0407:00000407";
        "Greek"="0408:00000408"; "English (United States)"="0409:00000409"; "Spanish"="040a:0000040a"; "Finnish"="040b:0000040b"; "French"="040c:0000040c"; "Hebrew"="040d:0000040d";
        "Hungarian"="040e:0000040e"; "Icelandic"="040f:0000040f"; "Italian"="0410:00000410"; "Japanese"="0411:00000411"; "Korean"="0412:00000412"; "Dutch"="0413:00000413";
        "Norwegian"="0414:00000414"; "Polish (Programmers)"="0415:00000415"; "Portuguese (Brazilian ABNT)"="0416:00000416"; "Romanian (Standard)"="0418:00000418"; "Russian"="0419:00000419"; "Croatian"="041a:0000041a";
        "Slovak"="041b:0000041b"; "Albanian"="041c:0000041c"; "Swedish"="041d:0000041d"; "Thai"="041e:0000041e"; "Turkish Q"="041f:0000041f"; "Urdu"="0420:00000420";
        "Ukrainian"="0422:00000422"; "Belarusian"="0423:00000423"; "Slovenian"="0424:00000424"; "Estonian"="0425:00000425"; "Latvian (Standard)"="0426:00000426"; "Lithuanian"="0427:00000427";
        "Persian"="0429:00000429"; "Vietnamese"="042a:0000042a"; "Armenian Eastern"="042b:0000042b"; "Azeri Latin"="042c:0000042c"; "Macedonian"="042f:0000042f"; "Georgian"="0437:00000437";
        "Kazakh"="043f:0000043f"; "English (United Kingdom)"="0809:00000809"; "Swiss German"="0807:00000807"; "Swiss French"="100c:0000100c"; "Serbian (Latin)"="081a:0000081a"
    }
    $sortedKeyboardLayouts = $keyboardLayoutData.GetEnumerator() | Sort-Object Name
    $listKeyboardLayouts.Items.AddRange($sortedKeyboardLayouts.Name) | Out-Null
    $groupKeyboard.Controls.Add($listKeyboardLayouts) | Out-Null; $yPos += $listKeyboardLayouts.Height + 10

    $groupKeyboard.Controls.Add((New-StyledLabel -Text "Current Selection:" -Location "15,$yPos")) | Out-Null
    $labelSelectedKeyboards = New-StyledLabel -Text "None" -Location "150,$yPos"; $labelSelectedKeyboards.ForeColor = [System.Drawing.Color]::LightSteelBlue; $labelSelectedKeyboards.AutoSize = $false; $labelSelectedKeyboards.Size = '580,50'
    $groupKeyboard.Controls.Add($labelSelectedKeyboards) | Out-Null

    $updateKeyboardLabel = {
        $checkedItemsText = ($checkedKeyboardLayoutNames | Sort-Object) -join ', '
        if ([string]::IsNullOrWhiteSpace($checkedItemsText)) { $labelSelectedKeyboards.Text = "None" } else { $labelSelectedKeyboards.Text = $checkedItemsText }
    }

    # Enforce a maximum of 5 keyboard layouts.
    $listKeyboardLayouts.Add_ItemCheck({
        param($sender, $e)
        $itemName = $sender.Items[$e.Index]
        if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
            if ($checkedKeyboardLayoutNames.Count -ge 5) {
                $e.NewValue = [System.Windows.Forms.CheckState]::Unchecked
                [System.Windows.Forms.MessageBox]::Show("You can select a maximum of 5 keyboard layouts.", "Selection Limit Reached", "OK", "Information") | Out-Null
            } else {
                [void]$checkedKeyboardLayoutNames.Add($itemName)
            }
        } else {
            [void]$checkedKeyboardLayoutNames.Remove($itemName)
        }
    }) | Out-Null
    $listKeyboardLayouts.Add_MouseUp({ & $updateKeyboardLabel }) | Out-Null

    $textKeyboardSearch.Add_TextChanged({
        $listKeyboardLayouts.BeginUpdate()
        $listKeyboardLayouts.Items.Clear()
        $searchText = $textKeyboardSearch.Text
        $filteredLayouts = $sortedKeyboardLayouts | Where-Object { $_.Name -match [regex]::Escape($searchText) }
        if ($null -ne $filteredLayouts) { $listKeyboardLayouts.Items.AddRange($filteredLayouts.Name) | Out-Null }
        
        # Re-check the items that were previously selected.
        for($i = 0; $i -lt $listKeyboardLayouts.Items.Count; $i++) {
            if ($checkedKeyboardLayoutNames.Contains($listKeyboardLayouts.Items[$i])) {
                $listKeyboardLayouts.SetItemChecked($i, $true)
            }
        }
        $listKeyboardLayouts.EndUpdate()
    }) | Out-Null
    
    # Pre-select the user's current keyboard layout.
    try { 
        $currentLayoutId = (Get-WinUserLanguageList)[0].InputMethodTips[0]
        $defaultKeyboardName = ($keyboardLayoutData.GetEnumerator() | Where-Object { $_.Value -eq $currentLayoutId }).Name
        if ($defaultKeyboardName) {
            [void]$checkedKeyboardLayoutNames.Add($defaultKeyboardName)
            $itemIndex = $listKeyboardLayouts.Items.IndexOf($defaultKeyboardName)
            if ($itemIndex -ne -1) { $listKeyboardLayouts.SetItemChecked($itemIndex, $true) }
        }
    } catch {}; & $updateKeyboardLabel

    # --- Tab 3: Automation & Tweaks ---
    $tabAutomation = New-Object System.Windows.Forms.TabPage; $tabAutomation.Text = "Automation & Tweaks"; $tabAutomation.BackColor = $form.BackColor; $tabAutomation.Padding = New-Object System.Windows.Forms.Padding(10)
    $tabControl.Controls.Add($tabAutomation) | Out-Null

    $groupOobe = New-StyledGroupBox "OOBE Skip Options" "15,15" "750,220"
    $tabAutomation.Controls.Add($groupOobe) | Out-Null
    $yPos = 30
    $checkHideEula = New-StyledCheckBox -Text "Hide EULA Page" -Location "20,$yPos" -Checked $true; $groupOobe.Controls.Add($checkHideEula) | Out-Null
    $toolTip.SetToolTip($checkHideEula, "Automatically accepts the End User License Agreement (EULA) during setup.")
    $yPos += 40
    $checkHideLocalAccount = New-StyledCheckBox -Text "Hide Local Account Screen" -Location "20,$yPos" -Checked $true; $groupOobe.Controls.Add($checkHideLocalAccount) | Out-Null
    $toolTip.SetToolTip($checkHideLocalAccount, "Bypasses the screen that prompts to create a local user account.")
    $yPos += 40
    $checkHideOnlineAccount = New-StyledCheckBox -Text "Hide Online Account Screens" -Location "20,$yPos" -Checked $true; $groupOobe.Controls.Add($checkHideOnlineAccount) | Out-Null
    $toolTip.SetToolTip($checkHideOnlineAccount, "Bypasses the screens that prompt to sign in with or create a Microsoft Account.")
    $yPos += 40
    $checkHideWireless = New-StyledCheckBox -Text "Hide Wireless Setup" -Location "20,$yPos" -Checked $true; $groupOobe.Controls.Add($checkHideWireless) | Out-Null
    $toolTip.SetToolTip($checkHideWireless, "Skips the network and Wi-Fi connection screen during the Out-of-Box Experience (OOBE).")

    $groupCustom = New-StyledGroupBox "First Logon System Tweaks" "15,250" "750,220"
    $tabAutomation.Controls.Add($groupCustom) | Out-Null
    $yPos = 30
    $checkShowFileExt = New-StyledCheckBox -Text "Show Known File Extensions" -Location "20,$yPos" -Checked $true; $groupCustom.Controls.Add($checkShowFileExt) | Out-Null
    $toolTip.SetToolTip($checkShowFileExt, "Configures File Explorer to show file extensions like '.exe', '.txt', '.dll' by default.")
    $yPos += 40
    $checkDisableSmartScreen = New-StyledCheckBox -Text "Disable SmartScreen" -Location "20,$yPos" -Checked $true; $groupCustom.Controls.Add($checkDisableSmartScreen) | Out-Null
    $toolTip.SetToolTip($checkDisableSmartScreen, "Turns off the Microsoft Defender SmartScreen filter, which checks for malicious files and websites.")
    $yPos += 40
    $checkDisableSysRestore = New-StyledCheckBox -Text "Disable System Restore" -Location "20,$yPos" -Checked $true; $groupCustom.Controls.Add($checkDisableSysRestore) | Out-Null
    $toolTip.SetToolTip($checkDisableSysRestore, "Disables the automatic creation of restore points. This can save disk space but limits recovery options.")
    $yPos += 40
    $checkDisableSuggestions = New-StyledCheckBox -Text "Disable App Suggestions" -Location "20,$yPos" -Checked $true; $groupCustom.Controls.Add($checkDisableSuggestions) | Out-Null
    $toolTip.SetToolTip($checkDisableSuggestions, "Prevents Windows from displaying app and content suggestions in the Start Menu and on the lock screen.")

    $automationInfoLabel = New-Object System.Windows.Forms.Label
    $automationInfoLabel.Text = "Hover over an option for a detailed description."
    $automationInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $automationInfoLabel.ForeColor = [System.Drawing.Color]::Gray
    $automationInfoLabel.AutoSize = $true
    $automationInfoLabel.Location = New-Object System.Drawing.Point(20, 485)
    $tabAutomation.Controls.Add($automationInfoLabel) | Out-Null

    # --- Tab 4: Bloatware Removal ---
    $tabBloatware = New-Object System.Windows.Forms.TabPage; $tabBloatware.Text = "Bloatware"; $tabBloatware.BackColor = $form.BackColor; $tabControl.Controls.Add($tabBloatware) | Out-Null
    $bloatTopPanel = New-Object System.Windows.Forms.Panel; $bloatTopPanel.Dock = "Top"; $bloatTopPanel.Height = 40; $bloatTopPanel.BackColor = $form.BackColor; $tabBloatware.Controls.Add($bloatTopPanel) | Out-Null
    $bloatTablePanel = New-Object System.Windows.Forms.TableLayoutPanel; $bloatTablePanel.Dock = "Fill"; $bloatTablePanel.AutoScroll = $true; $bloatTablePanel.BackColor = $form.BackColor; $tabBloatware.Controls.Add($bloatTablePanel) | Out-Null; $bloatTablePanel.BringToFront()
    $bloatBottomPanel = New-Object System.Windows.Forms.Panel; $bloatBottomPanel.Dock = "Bottom"; $bloatBottomPanel.Height = 40; $bloatBottomPanel.BackColor = $form.BackColor; $tabBloatware.Controls.Add($bloatBottomPanel) | Out-Null
    $bloatwareCheckboxes = @()
    # A comprehensive list of removable apps and features.
    $bloatwareList = @(
        '3D Viewer', 'Bing Search', 'Calculator', 'Camera', 'Clipchamp', 'Clock', 'Copilot', 'Cortana', 'Dev Home',
        'Family', 'Feedback Hub', 'Get Help', 'Handwriting (all languages)', 'Internet Explorer', 'Mail and Calendar',
        'Maps', 'Math Input Panel', 'Media Features', 'Mixed Reality', 'Movies & TV', 'News', 'Notepad (modern)',
        'Office 365', 'OneDrive', 'OneNote', 'OneSync', 'OpenSSH Client', 'Outlook for Windows', 'Paint', 'Paint 3D',
        'People', 'Photos', 'Power Automate', 'PowerShell 2.0', 'PowerShell ISE', 'Quick Assist', 'Recall',
        'Remote Desktop Client', 'Skype', 'Snipping Tool', 'Solitaire Collection', 'Speech (all languages)',
        'Steps Recorder', 'Sticky Notes', 'Teams', 'Tips', 'To Do', 'Voice Recorder', 'Wallet', 'Weather',
        'Windows Fax and Scan', 'Windows Hello', 'Windows Media Player (classic)', 'Windows Media Player (modern)',
        'Windows Terminal', 'WordPad', 'Xbox Apps', 'Your Phone / Phone Link'
    ) | Sort-Object
    # Use a TableLayoutPanel for a clean, multi-column layout.
    $bloatTablePanel.ColumnCount = 3; $rowsNeeded = [math]::Ceiling($bloatwareList.Count / $bloatTablePanel.ColumnCount); $bloatTablePanel.RowCount = $rowsNeeded
    for ($i = 0; $i -lt $bloatTablePanel.ColumnCount; $i++) { $bloatTablePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) | Out-Null }
    for ($i = 0; $i -lt $bloatTablePanel.RowCount; $i++) { $bloatTablePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null }
    $col = 0; $row = 0
    foreach ($appName in $bloatwareList) {
        $checkbox = New-StyledCheckBox -Text $appName -Location "0,0" -Checked $false; $checkbox.Margin = [System.Windows.Forms.Padding]::new(10, 5, 10, 5)
        $bloatTablePanel.Controls.Add($checkbox, $col, $row) | Out-Null; $bloatwareCheckboxes += $checkbox; $col++; if ($col -ge $bloatTablePanel.ColumnCount) { $col = 0; $row++ }
    }
    $btnSelectAll = New-Object System.Windows.Forms.Button; $btnSelectAll.Text = "Select All"; $btnSelectAll.Size = "120,30"; $btnSelectAll.Location = "10,5"; $btnSelectAll.Font = $commonFont; $btnSelectAll.ForeColor = [System.Drawing.Color]::White; $btnSelectAll.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180); $btnSelectAll.FlatStyle = "Flat"; $btnSelectAll.FlatAppearance.BorderSize = 0
    $btnSelectAll.add_Click({ foreach($cb in $bloatwareCheckboxes) {$cb.Checked = $true} }) | Out-Null; $bloatTopPanel.Controls.Add($btnSelectAll) | Out-Null
    $btnDeselectAll = New-Object System.Windows.Forms.Button; $btnDeselectAll.Text = "Deselect All"; $btnDeselectAll.Size = "120,30"; $btnDeselectAll.Location = "140,5"; $btnDeselectAll.Font = $commonFont; $btnDeselectAll.ForeColor = [System.Drawing.Color]::White; $btnDeselectAll.BackColor = [System.Drawing.Color]::FromArgb(90, 90, 90); $btnDeselectAll.FlatStyle = "Flat"; $btnDeselectAll.FlatAppearance.BorderSize = 0
    $btnDeselectAll.add_Click({ foreach($cb in $bloatwareCheckboxes) {$cb.Checked = $false} }) | Out-Null; $bloatTopPanel.Controls.Add($btnDeselectAll) | Out-Null
    $infoLabel = New-Object System.Windows.Forms.Label; $infoLabel.Text = "Bloatware removal works best with original Win 10 and 11 ISOs. Functionality on custom images is not guaranteed."; $infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8); $infoLabel.ForeColor = [System.Drawing.Color]::Gray; $infoLabel.Dock = "Fill"; $infoLabel.TextAlign = "MiddleCenter"; $bloatBottomPanel.Controls.Add($infoLabel) | Out-Null

    # --- Create and Cancel Buttons ---
    $buttonCreate = New-Object System.Windows.Forms.Button; $buttonCreate.Text = "Create"; $buttonCreate.Size = "120,30"; $buttonCreate.Location = "265,10"; $buttonCreate.Font = $commonFont; $buttonCreate.ForeColor = [System.Drawing.Color]::White; $buttonCreate.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180); $buttonCreate.FlatStyle = "Flat"; $buttonCreate.FlatAppearance.BorderSize = 0;
    $buttonCreate.add_Click({$form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close()}) | Out-Null
    $buttonPanel.Controls.Add($buttonCreate) | Out-Null
    $buttonCancel = New-Object System.Windows.Forms.Button; $buttonCancel.Text = "Cancel"; $buttonCancel.Size = "120,30"; $buttonCancel.Location = "395,10"; $buttonCancel.Font = $commonFont; $buttonCancel.ForeColor = [System.Drawing.Color]::White; $buttonCancel.BackColor = [System.Drawing.Color]::FromArgb(90, 90, 90); $buttonCancel.FlatStyle = "Flat"; $buttonCancel.FlatAppearance.BorderSize = 0;
    $buttonCancel.add_Click({$form.Close()}) | Out-Null; $buttonPanel.Controls.Add($buttonCancel) | Out-Null

    # Show the form and check the result.
    try {
        $result = $form.ShowDialog()
    }
    catch {
        Write-LogAndHost "An unexpected error occurred with the Unattend XML Creator GUI. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Create-UnattendXml"
        $result = [System.Windows.Forms.DialogResult]::Cancel
    }
    finally {
        $form.Dispose()
    }
    
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-LogAndHost "XML creation cancelled by user." -HostColor Yellow; Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host; return
    }

    # Collect data from the form controls.
    $selectedKeyboardLayouts = $checkedKeyboardLayoutNames | ForEach-Object { $keyboardLayoutData[$_] }
    $selectedTimeZoneId = if ($listTimeZone.SelectedItem) { $timeZoneMap[$listTimeZone.SelectedItem] } else { $null }

    $formData = @{
        ComputerName = $textComputerName.Text; UserName = $textUserName.Text; Password = $textPassword.Text
        UiLanguage = $comboUiLanguage.Text; SystemLocale = $comboSystemLocale.Text; UserLocale = $comboUserLocale.Text
        TimeZone = $selectedTimeZoneId
        KeyboardLayouts = $selectedKeyboardLayouts -join ';'
        HideEula = $checkHideEula.Checked; HideLocalAccount = $checkHideLocalAccount.Checked; HideOnlineAccount = $checkHideOnlineAccount.Checked; HideWireless = $checkHideWireless.Checked
        ShowFileExt = $checkShowFileExt.Checked; DisableSmartScreen = $checkDisableSmartScreen.Checked; DisableSysRestore = $checkDisableSysRestore.Checked; DisableSuggestions = $checkDisableSuggestions.Checked
        BloatwareToRemove = ($bloatwareCheckboxes | Where-Object { $_.Checked } | ForEach-Object { $_.Text })
    }

    # --- Validation ---
    if ([string]::IsNullOrWhiteSpace($formData.ComputerName) -or [string]::IsNullOrWhiteSpace($formData.UserName) -or `
        [string]::IsNullOrWhiteSpace($formData.UiLanguage) -or [string]::IsNullOrWhiteSpace($formData.SystemLocale) -or `
        [string]::IsNullOrWhiteSpace($formData.UserLocale) -or [string]::IsNullOrWhiteSpace($formData.TimeZone) -or `
        $selectedKeyboardLayouts.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please fill in all general and regional settings, including at least one keyboard layout.", "Validation Failed", "OK", "Error") | Out-Null
        Write-LogAndHost "XML creation aborted due to missing required fields." -HostColor Red -LogPrefix "Create-UnattendXml"
        Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host; return
    }

    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $filePath = Join-Path -Path $desktopPath -ChildPath "autounattend.xml"
    Write-LogAndHost "Creating XML structure based on GUI selections..." -NoHost
        
    # --- XML Generation ---
    $xml = New-Object System.Xml.XmlDocument
    $xml.AppendChild($xml.CreateXmlDeclaration("1.0", "utf-8", $null)) | Out-Null
    $root = $xml.CreateElement("unattend"); $root.SetAttribute("xmlns", "urn:schemas-microsoft-com:unattend"); $xml.AppendChild($root) | Out-Null
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable); $ns.AddNamespace("d6p1", "http://schemas.microsoft.com/WMIConfig/2002/State")
    
    # Helper function to create XML components correctly.
    function New-Component($ParentNode, $Name, $Pass, $Token="31bf3856ad364e35", $Arch="amd64") {
        $settings = $ParentNode.SelectSingleNode("//unattend/settings[@pass='$Pass']")
        if (-not $settings) { $settings = $ParentNode.OwnerDocument.CreateElement("settings"); $settings.SetAttribute("pass", $Pass); $ParentNode.AppendChild($settings) | Out-Null }
        $component = $settings.SelectSingleNode("component[@name='$Name']")
        if (-not $component) {
            $component = $ParentNode.OwnerDocument.CreateElement("component"); $component.SetAttribute("name", $Name); $component.SetAttribute("processorArchitecture", $Arch); $component.SetAttribute("publicKeyToken", $Token); $component.SetAttribute("language", "neutral"); $component.SetAttribute("versionScope", "nonSxS")
            $settings.AppendChild($component) | Out-Null
        }
        return $component
    }
    
    # --- Build XML from formData ---
    # Pass 4: specialize
    $compIntlSpec = New-Component -ParentNode $root -Name "Microsoft-Windows-International-Core" -Pass "specialize"
    $compIntlSpec.AppendChild($xml.CreateElement("InputLocale")).InnerText = $formData.KeyboardLayouts
    $compIntlSpec.AppendChild($xml.CreateElement("SystemLocale")).InnerText = $formData.SystemLocale
    $compIntlSpec.AppendChild($xml.CreateElement("UILanguage")).InnerText = $formData.UiLanguage
    $compIntlSpec.AppendChild($xml.CreateElement("UserLocale")).InnerText = $formData.UserLocale

    $compShellSpec = New-Component -ParentNode $root -Name "Microsoft-Windows-Shell-Setup" -Pass "specialize"
    $compShellSpec.AppendChild($xml.CreateElement("ComputerName")).InnerText = $formData.ComputerName
    $compShellSpec.AppendChild($xml.CreateElement("TimeZone")).InnerText = $formData.TimeZone
    
    # Pass 7: oobeSystem
    $compIntlOobe = New-Component -ParentNode $root -Name "Microsoft-Windows-International-Core" -Pass "oobeSystem"
    $compIntlOobe.AppendChild($xml.CreateElement("InputLocale")).InnerText = $formData.KeyboardLayouts
    $compIntlOobe.AppendChild($xml.CreateElement("SystemLocale")).InnerText = $formData.SystemLocale
    $compIntlOobe.AppendChild($xml.CreateElement("UILanguage")).InnerText = $formData.UiLanguage
    $compIntlOobe.AppendChild($xml.CreateElement("UserLocale")).InnerText = $formData.UserLocale

    $compShellOobe = New-Component -ParentNode $root -Name "Microsoft-Windows-Shell-Setup" -Pass "oobeSystem"
    $oobeNode = $compShellOobe.AppendChild($xml.CreateElement("OOBE"))
    if ($formData.HideEula) { $oobeNode.AppendChild($xml.CreateElement("HideEULAPage")).InnerText = "true" }
    if ($formData.HideLocalAccount) { $oobeNode.AppendChild($xml.CreateElement("HideLocalAccountScreen")).InnerText = "true" }
    if ($formData.HideOnlineAccount) { $oobeNode.AppendChild($xml.CreateElement("HideOnlineAccountScreens")).InnerText = "true" }
    if ($formData.HideWireless) { $oobeNode.AppendChild($xml.CreateElement("HideWirelessSetupInOOBE")).InnerText = "true" }
    $oobeNode.AppendChild($xml.CreateElement("ProtectYourPC")).InnerText = "1"

    $userAccounts = $compShellOobe.AppendChild($xml.CreateElement("UserAccounts"))
    $localAccount = $userAccounts.AppendChild($xml.CreateElement("LocalAccounts")).AppendChild($xml.CreateElement("LocalAccount"))
    $localAccount.SetAttribute("action", $ns.LookupNamespace("d6p1"), "add"); $localAccount.AppendChild($xml.CreateElement("Name")).InnerText = $formData.UserName
    $localAccount.AppendChild($xml.CreateElement("Group")).InnerText = "Administrators"; $localAccount.AppendChild($xml.CreateElement("DisplayName")).InnerText = $formData.UserName
    $passwordNode = $localAccount.AppendChild($xml.CreateElement("Password")); $passwordNode.AppendChild($xml.CreateElement("Value")).InnerText = $formData.Password; $passwordNode.AppendChild($xml.CreateElement("PlainText")).InnerText = "true"
    
    # FirstLogonCommands are executed after the user logs on for the first time.
    $firstLogonCommands = $compShellOobe.AppendChild($xml.CreateElement("FirstLogonCommands"))
    $commandIndex = 1
    
    if ($formData.ShowFileExt) { $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand")); $syncCmd.SetAttribute("Order", $commandIndex++); $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = 'cmd /c reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v HideFileExt /t REG_DWORD /d 0 /f' }
    if ($formData.DisableSmartScreen) { $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand")); $syncCmd.SetAttribute("Order", $commandIndex++); $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = 'cmd /c reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" /v SmartScreenEnabled /t REG_SZ /d "Off" /f' }
    if ($formData.DisableSysRestore) { $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand")); $syncCmd.SetAttribute("Order", $commandIndex++); $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = 'powershell.exe -Command "Disable-ComputerRestore -Drive C:\"' }
    if ($formData.DisableSuggestions) { $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand")); $syncCmd.SetAttribute("Order", $commandIndex++); $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = 'cmd /c reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v SubscribedContent-338389Enabled /t REG_DWORD /d 0 /f' }

    # A map of friendly bloatware names to their removal commands.
    $bloatwareCommands = @{
        '3D Viewer' = 'Get-AppxPackage *Microsoft.Microsoft3DViewer* | Remove-AppxPackage -AllUsers'; 'Bing Search' = 'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Search" /v BingSearchEnabled /t REG_DWORD /d 0 /f'; 'Calculator' = 'Get-AppxPackage *Microsoft.WindowsCalculator* | Remove-AppxPackage -AllUsers'; 'Camera' = 'Get-AppxPackage *Microsoft.WindowsCamera* | Remove-AppxPackage -AllUsers'; 'Clipchamp' = 'Get-AppxPackage *Microsoft.Clipchamp* | Remove-AppxPackage -AllUsers'; 'Clock' = 'Get-AppxPackage *Microsoft.WindowsAlarms* | Remove-AppxPackage -AllUsers'; 'Copilot' = 'reg add "HKCU\Software\Policies\Microsoft\Windows\WindowsCopilot" /v TurnOffWindowsCopilot /t REG_DWORD /d 1 /f'; 'Cortana' = 'Get-AppxPackage *Microsoft.549981C3F5F10* | Remove-AppxPackage -AllUsers'; 'Dev Home' = 'Get-AppxPackage *Microsoft.DevHome* | Remove-AppxPackage -AllUsers'; 'Family' = 'Get-AppxPackage *Microsoft.Windows.Family* | Remove-AppxPackage -AllUsers'; 'Feedback Hub' = 'Get-AppxPackage *Microsoft.WindowsFeedbackHub* | Remove-AppxPackage -AllUsers'; 'Get Help' = 'Get-AppxPackage *Microsoft.GetHelp* | Remove-AppxPackage -AllUsers'; 'Handwriting (all languages)' = 'Get-WindowsCapability -Online | Where-Object { $_.Name -like "Language.Handwriting*" } | ForEach-Object { Remove-WindowsCapability -Online -Name $_.Name -NoRestart }'; 'Internet Explorer' = 'Disable-WindowsOptionalFeature -Online -FeatureName "Internet-Explorer-Optional-amd64" -NoRestart'; 'Mail and Calendar' = 'Get-AppxPackage *microsoft.windowscommunicationsapps* | Remove-AppxPackage -AllUsers'; 'Maps' = 'Get-AppxPackage *Microsoft.WindowsMaps* | Remove-AppxPackage -AllUsers'; 'Math Input Panel' = 'Remove-WindowsCapability -Online -Name "MathRecognizer~~~~0.0.1.0" -NoRestart'; 'Media Features' = 'Disable-WindowsOptionalFeature -Online -FeatureName "MediaPlayback" -NoRestart'; 'Mixed Reality' = 'Get-AppxPackage *Microsoft.MixedReality.Portal* | Remove-AppxPackage -AllUsers'; 'Movies & TV' = 'Get-AppxPackage *Microsoft.ZuneVideo* | Remove-AppxPackage -AllUsers'; 'News' = 'Get-AppxPackage *Microsoft.BingNews* | Remove-AppxPackage -AllUsers'; 'Notepad (modern)' = 'Get-AppxPackage *Microsoft.WindowsNotepad* | Remove-AppxPackage -AllUsers'; 'Office 365' = 'Get-AppxPackage *Microsoft.MicrosoftOfficeHub* | Remove-AppxPackage -AllUsers'; 'OneDrive' = '$process = Start-Process "$env:SystemRoot\SysWOW64\OneDriveSetup.exe" -ArgumentList "/uninstall" -PassThru -Wait; if ($process.ExitCode -ne 0) { Start-Process "$env:SystemRoot\System32\OneDriveSetup.exe" -ArgumentList "/uninstall" -PassThru -Wait }'; 'OneNote' = 'Get-AppxPackage *Microsoft.Office.OneNote* | Remove-AppxPackage -AllUsers'; 'OneSync' = '# Handled by Mail and Calendar'; 'OpenSSH Client' = 'Remove-WindowsCapability -Online -Name "OpenSSH.Client~~~~0.0.1.0" -NoRestart'; 'Outlook for Windows' = 'Get-AppxPackage *Microsoft.OutlookForWindows* | Remove-AppxPackage -AllUsers'; 'Paint' = 'Get-AppxPackage *Microsoft.Paint* | Remove-AppxPackage -AllUsers'; 'Paint 3D' = 'Get-AppxPackage *Microsoft.MSPaint* | Remove-AppxPackage -AllUsers'; 'People' = 'Get-AppxPackage *Microsoft.People* | Remove-AppxPackage -AllUsers'; 'Photos' = 'Get-AppxPackage *Microsoft.Windows.Photos* | Remove-AppxPackage -AllUsers'; 'Power Automate' = 'Get-AppxPackage *Microsoft.PowerAutomateDesktop* | Remove-AppxPackage -AllUsers'; 'PowerShell 2.0' = 'Disable-WindowsOptionalFeature -Online -FeatureName "MicrosoftWindowsPowerShellV2" -NoRestart'; 'PowerShell ISE' = 'Remove-WindowsCapability -Online -Name "PowerShell-ISE-v2~~~~0.0.1.0" -NoRestart'; 'Quick Assist' = 'Get-AppxPackage *Microsoft.QuickAssist* | Remove-AppxPackage -AllUsers'; 'Recall' = 'reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" /v DisableAllScreenshotCapture /t REG_DWORD /d 1 /f'; 'Remote Desktop Client' = '# Core component, removal not recommended.'; 'Skype' = 'Get-AppxPackage *Microsoft.SkypeApp* | Remove-AppxPackage -AllUsers'; 'Snipping Tool' = 'Get-AppxPackage *Microsoft.ScreenSketch* | Remove-AppxPackage -AllUsers'; 'Solitaire Collection' = 'Get-AppxPackage *Microsoft.MicrosoftSolitaireCollection* | Remove-AppxPackage -AllUsers'; 'Speech (all languages)' = 'Get-WindowsCapability -Online | Where-Object { $_.Name -like "Language.Speech*" } | ForEach-Object { Remove-WindowsCapability -Online -Name $_.Name -NoRestart }'; 'Steps Recorder' = 'Disable-WindowsOptionalFeature -Online -FeatureName "StepsRecorder" -NoRestart'; 'Sticky Notes' = 'Get-AppxPackage *Microsoft.MicrosoftStickyNotes* | Remove-AppxPackage -AllUsers'; 'Teams' = 'Get-AppxPackage *MicrosoftTeams* | Remove-AppxPackage -AllUsers'; 'Tips' = 'Get-AppxPackage *Microsoft.Getstarted* | Remove-AppxPackage -AllUsers'; 'To Do' = 'Get-AppxPackage *Microsoft.Todos* | Remove-AppxPackage -AllUsers'; 'Voice Recorder' = 'Get-AppxPackage *Microsoft.WindowsSoundRecorder* | Remove-AppxPackage -AllUsers'; 'Wallet' = 'Get-AppxPackage *Microsoft.Wallet* | Remove-AppxPackage -AllUsers'; 'Weather' = 'Get-AppxPackage *Microsoft.BingWeather* | Remove-AppxPackage -AllUsers'; 'Windows Fax and Scan' = 'Disable-WindowsOptionalFeature -Online -FeatureName "Windows-Fax-And-Scan" -NoRestart'; 'Windows Hello' = 'reg add "HKLM\SOFTWARE\Policies\Microsoft\Biometrics" /v Enabled /t REG_DWORD /d 0 /f; reg add "HKLM\SOFTWARE\Policies\Microsoft\Biometrics\CredentialProviders" /v Enabled /t REG_DWORD /d 0 /f'; 'Windows Media Player (classic)' = 'Disable-WindowsOptionalFeature -Online -FeatureName "WindowsMediaPlayer" -NoRestart'; 'Windows Media Player (modern)' = 'Get-AppxPackage *Microsoft.ZuneMusic* | Remove-AppxPackage -AllUsers'; 'Windows Terminal' = 'Get-AppxPackage *Microsoft.WindowsTerminal* | Remove-AppxPackage -AllUsers'; 'WordPad' = 'Remove-WindowsCapability -Online -Name "WordPad~~~~0.0.1.0" -NoRestart'; 'Xbox Apps' = 'Get-AppxPackage *Microsoft.Xbox* | Remove-AppxPackage -AllUsers; Get-AppxPackage *Microsoft.GamingApp* | Remove-AppxPackage -AllUsers'; 'Your Phone / Phone Link' = 'Get-AppxPackage *Microsoft.YourPhone* | Remove-AppxPackage -AllUsers'
    }

    foreach ($bloat in $formData.BloatwareToRemove) {
        if ($bloatwareCommands.ContainsKey($bloat)) {
            $command = $bloatwareCommands[$bloat]; if ($command.StartsWith("#")) { continue } # Skip commented-out commands.
            # Use PowerShell's -EncodedCommand for complex commands to avoid quoting issues.
            if ($command -match 'Get-AppxPackage|Remove-AppxPackage|Get-WindowsCapability|Remove-WindowsCapability|Disable-WindowsOptionalFeature|Start-Process') {
                $bytes = [System.Text.Encoding]::Unicode.GetBytes($command); $encodedCommand = [Convert]::ToBase64String($bytes)
                $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand")); $syncCmd.SetAttribute("Order", $commandIndex++); $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = "powershell.exe -EncodedCommand $encodedCommand"
            } elseif ($command -match 'reg add|reg delete') {
                 $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand")); $syncCmd.SetAttribute("Order", $commandIndex++); $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = "cmd /c $command"
            }
        }
    }

    Write-LogAndHost "Saving file to: $filePath"
    try {
        $xml.Save($filePath)
        Write-LogAndHost "autounattend.xml created successfully on your Desktop." -HostColor Green
        Write-LogAndHost "Copy this file to the root of your Windows installation USB drive." -HostColor Yellow
    }
    catch {
        Write-LogAndHost "Failed to save the XML file to '$filePath'. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Create-UnattendXml"
    }
    
    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}

# Function to install programs using Chocolatey.
function Install-Programs {
    param (
        [string[]]$ProgramsToInstall,
        [string]$Source = "https://community.chocolatey.org/api/v2/"
    )

    if ($ProgramsToInstall.Count -eq 0) {
        Write-LogAndHost "No programs to install." -HostColor Yellow
        return $true
    }

    $allSuccess = $true
    Write-LogAndHost "Starting installation of $($ProgramsToInstall.Count) program(s)..." -HostColor Cyan
    Write-Host ""

    foreach ($program in $ProgramsToInstall) {
        Write-LogAndHost "Installing $program..." -LogPrefix "Install-Programs"
        try {
            $installOutput = & choco install $program -y --source=$Source --no-progress 2>&1
            $installOutput | Out-File -FilePath $script:logFile -Append -Encoding UTF8

            if ($LASTEXITCODE -eq 0) {
                # Check output to see if the program was already there.
                if ($installOutput -match "is already installed|already installed|Nothing to do") {
                    Write-LogAndHost "$program is already installed or up to date." -HostColor Green
                } else {
                    Write-LogAndHost "$program installed successfully." -HostColor White
                }
            } else {
                $allSuccess = $false
                Write-LogAndHost "Failed to install $program. Exit code: $LASTEXITCODE. Details: $($installOutput | Out-String)" -HostColor Red -LogPrefix "Install-Programs"
            }
        } catch {
            $allSuccess = $false
            Write-LogAndHost "Exception occurred while installing $program. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Install-Programs"
        }
        Write-Host ""
    }
    return $allSuccess
}

# Function to get a list of locally installed Chocolatey packages.
function Get-InstalledChocolateyPackages {
    $chocoLibPath = Join-Path -Path $env:ChocolateyInstall -ChildPath "lib"
    $installedPackages = @()
    if (Test-Path $chocoLibPath) {
        try {
            # Packages are stored as directories in the 'lib' folder.
            $installedPackages = Get-ChildItem -Path $chocoLibPath -Directory | Select-Object -ExpandProperty Name
            Write-LogAndHost "Found installed packages: $($installedPackages -join ', ')" -NoHost
        } catch {
            Write-LogAndHost "Could not retrieve installed packages from $chocoLibPath. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Get-InstalledChocolateyPackages"
        }
    } else {
        Write-LogAndHost "Chocolatey lib directory not found at $chocoLibPath. Cannot list installed packages." -HostColor Yellow -LogPrefix "Get-InstalledChocolateyPackages"
    }
    return $installedPackages
}

# Function to uninstall programs using Chocolatey.
function Uninstall-Programs {
    param (
        [string[]]$ProgramsToUninstall
    )

    if ($ProgramsToUninstall.Count -eq 0) {
        Write-LogAndHost "No programs selected for uninstallation." -HostColor Yellow
        return $true
    }

    $allSuccess = $true
    Write-LogAndHost "Starting uninstallation of $($ProgramsToUninstall.Count) program(s)..." -HostColor Cyan
    Write-Host ""

    foreach ($program in $ProgramsToUninstall) {
        Write-LogAndHost "Uninstalling $program..." -LogPrefix "Uninstall-Programs"
        try {
            $uninstallOutput = & choco uninstall $program -y --no-progress 2>&1
            $uninstallOutput | Out-File -FilePath $script:logFile -Append -Encoding UTF8

            if ($LASTEXITCODE -eq 0) {
                Write-LogAndHost "$program uninstalled successfully." -HostColor White
            } else {
                $allSuccess = $false
                Write-LogAndHost "Failed to uninstall $program. Exit code: $LASTEXITCODE. Details: $($uninstallOutput | Out-String)" -HostColor Red -LogPrefix "Uninstall-Programs"
            }
        } catch {
            $allSuccess = $false
            Write-LogAndHost "Exception occurred while uninstalling $program. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Uninstall-Programs"
        }
        Write-Host ""
    }
    return $allSuccess
}

# Function to test if a Chocolatey package exists in the repository.
function Test-ChocolateyPackage {
    param (
        [string]$PackageName
    )
    Write-LogAndHost "Searching for package '$PackageName' in Chocolatey repository..." -NoLog
    try {
        # Use --exact and --limit-output for a fast, clean search.
        $searchOutput = & choco search $PackageName --exact --limit-output --source="https://community.chocolatey.org/api/v2/" --no-progress 2>&1
        $searchOutput | Out-File -FilePath $script:logFile -Append -Encoding UTF8

        if ($LASTEXITCODE -ne 0) {
             Write-LogAndHost "Error during 'choco search' for '$PackageName'. Exit code: $LASTEXITCODE. Output: $($searchOutput | Out-String)" -HostColor Red -LogPrefix "Test-ChocolateyPackage"
             return $false
        }
        
        # Check if the output indicates that one package was found.
        if ($searchOutput -match "$([regex]::Escape($PackageName))\|.*" -or $searchOutput -match "1 packages found.") {
             Write-LogAndHost "Package '$PackageName' found in repository." -HostColor Green
             return $true
        } else {
             Write-LogAndHost "Package '$PackageName' not found as an exact match in Chocolatey repository. Search output: $($searchOutput | Out-String)" -HostColor Yellow -LogPrefix "Test-ChocolateyPackage"
             return $false
        }

    } catch {
        Write-LogAndHost "Exception occurred while searching for package '$PackageName'. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Test-ChocolateyPackage"
        return $false
    }
}

# --- EASTER EGG FUNCTION ---
function Show-PerdangaArt {
    Clear-Host
    $perdangaArt = @"
.................................................................................................................................
.................................................................................................................................
.................................................................................................................................
.................................................................................................................................
.................................................................................................................................
.................................................................................................................................
.................................................................................................................................
.................................................................................................................................
.................................................................................................................................
.........@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@.........
.........@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@.........
.........@@.....:::::-+#%@@@@#+++++============+*+++++=+++++++++==++====+++==++++=====++++++==++**++++++*##*=-::......@@.........
.........@@....:..::-=+*#%#*++++====-----=------======-------==-----=---========------==========-=====-+####+=-:......@@.........
.........@@......:::--+####+=======-----------::::------::::-----::--------==------::::--:--------=-==-=**#*+==-......@@.........
.........@@....:..::--=+%@%+=====--=-------::::.:--::.:::::::----:--:-:------::.::::::---:-----=========+***+=--......@@.........
.........@@........:::--+@@#++=====-----:---=------------::-:::::::-::--::::::::::::::-----==--========+*+*#+-:.......@@.........
.........@@...........--+%@@#++======+===-:-+###*++=======+++++==--=---------:-------------------======++++**+-.......@@.........
.........@@.......::::::+@@@%++==+++++===-*%@%%%%##*++++++***##*****+*#####***++++++++=--:::----========+++*+=-.......@@.........
.........@@...........:-%@@@%#+++**+=-:-*@@@%#*++======++++====++++*******########%%@@@%+-:-----=====++++++*++-.......@@.........
.........@@...........-*@@@@%#+++=---:=%@@@#++++==-::---------------------=====+*#%@@@@@@%=-:-====++++###*###+-.......@@.........
.........@@...........-#@@%#+=======+#@@@@#*+=+#+=--::::::::--::---------------=++++++#@@@%=--=++++*+=+#%%%%#+-.......@@.........
.........@@...........-#@@@@@#*++++*@@@@@#+=--+#+=-----:::......:::::::....:--=========+%@@@*-.-===+*+==+#%@%+-.......@@.........
.........@@...........:+%@@@@@%###%@@@@@**+=+++*=---::::......:............::---=========%@@@@+======+==#@@@#=-.......@@.........
.........@@............-+%%@@%%%%%@@@@@+=#+++====--:::::::........:....:::------=========+#@@@@*+++#%@%@@@@%+-........@@.........
.........@@.............-*%@#*@@@@@@@*--*#+++====--------::........:::----------======+=-=+%@@@%###@@@@@@@%*=:........@@.........
.........@@.............:=@@##%%*+%@#-#@@@@@%*==----:::::::::::::::-:::-----:::--====+++===*%@@%%@@%%%@@%#+=-.........@@.........
.........@@..........:...-#@#--+*+@@-@@@@@@@@@@%*==---:::..:::::::::...:::::--=++#%@@@@@+---*@@++###+*%#=---.::::::...@@.........
.........@@.......:..:....-#@#*++=@%*@@@+++#%@@@@%##*+=-::..::::...::---==+*##%%@@@@@@@@@@+--@@+=++-:=#%*=-------:::::@@.........
.........@@...............:=@@*--#@+%@@@@@@%*=-:=+###*+==-:.....:.:-==+#%@@@@%%##*+===*@@@%=-#@%=*#-:=+##+------------@@.........
.........@@................:+@@@=%@:#*###%##*++++=--------:...::::--===++=-:--=+*##%%@@%@@@*=-@@:+#=-=+##+-:::::...::.@@.........
.........@@..............:..:+@@%@*+#+*%@@%*+*##+=---=--=----:..:::----===+==----=====+*+#%%%+@@++#*-+#@%=------------@@.........
.........@@.:::...............=@@#-#%#****#%%######*===-==+==---===-------=*#%%%##+*####*+*#%%-@@=*#*@@@+:..::::::::::@@.........
.........@@...................:@*=-#%++@@@@#@@++====:-==*%#*+--=+*+===++=--+====++--##-=**++*%=@@@@@@@%+-::...........@@.........
.........@@..:................@:===#%%%*+%@%######+=--=#@@##+=-=++%%+==--==*#%%%%%#@@@@%#***##*+%@@@@*-...............@@.........
.........@@:.:.:::::::::::::::@==.-+*+=+@@@#++*++===+#@@@@##+--==-+@@%#*+==+#%%#####*-=%@@@@@%+++%@=::................@@.........
.........@@.........::::::::::@*++:=++###%#=--====+*%@@@@=##+-----:+@@@%*===---====++=++**###+-#+%@=..................@@.........
.........@@.................::@%=*-++#*##+=======+++@@@%-*%#=-.-===-=*##*=======---=++====++--:#*%@=..................@@.........
.........@@..................-@###-++*++*+===--:::-=@@%-#%%#+===++*#+=++=-:::-----=+++===+++=-:#*%@=..................@@.........
.........@@..................=@%%*-=++++++==---===+@@@=%@@@*===+*#+=#%%%*-::::---==++==--=++==-#+%@=..................@@.........
.........@@..................+@##+-==++++++====--+@@%=%@@%*=---=*%@%+-=@@+.:::---====-==-=+====#=%@+..................@@.........
.........@@.................:+@*%+====**+++===--:%@@@@@%#*=-:::-=+@@@@@#%#.::::---======-===-=*#=#@*:.................@@.........
.........@@.................:#@@%+===+##*++==--::%@@@@@%++=-::-:-+%%%@@@@#:::::----==========*%#=@@+:.................@@.........
.........@@.................:#@@#*+==+##++===-..:-=+=:+#%#+==+++**=-==-==-..::-----=======-=*##++@@-..................@@.........
.........@@.................:+@@-++==+##+==---====-----+#*##*###+===--::::----.----========*###+@@@:..................@@.........
.........@@...............::.-@@+===-+**+=-+##*+=-:--==--==++++=--:---::.:--==+=---======+*##*=%@@+:..................@@.........
.........@@..................:#@@=+++=+++=+##+==------=+=============-:.:::---+*+---====++++=:+@@#....................@@.........
.........@@............::.....=@@-+++=-=+==-=+##+=-----==++####*++==-:::.:--===-+=--=--==+=-:=@@@-....................@@.........
.........@@...................-@@#--:---=-=*%%%*++++++===::--=-::-++*#*+++++=+*+:------=====-%@@*:....................@@.........
.........@@....................%@@-=:..--===+=---=--:::-------------::::.:---=+++-:::.:-----+@@%:.....................@@.........
.........@@....................=@@#-=-::=++=----==++++===-.::-:.-=+*##*####+=--==--::-==--:+@@@=......................@@.........
.........@@....................:%@@=++==-=++---:.--------===+==----:::.::.::::-==--:=+*++=-@@@#:......................@@.........
.........@@.....................-@@@-*#*+====---::--==--==++++++====-------::.-===--+*+===@@@%-.......................@@.........
.........@@.....................:@@@@++**+=---===---::---======---------------===-=*#+=-#@@@%.........................@@.........
.........@@......................-@@@@@-=+++-::---===========+===---====--------=##*+-:%@@@@-.........................@@.........
.........@@.......................-+@@@@%-==---::-------=-:-----:------:::::::-++*+=--@@@@+...........................@@.........
.........@@........................:=#@@@@=:===-:--::---=------------------===+++==-+@@@@+............................@@.........
.........@@.....................:=++*+*@@@@+--==--------==--::-------===--=++====-.#@@@@+:............................@@.........
.........@@.................:-=+*##*=-=*%@@@#=-==----=====-------========-+==+===#@@@@#**+=-:.........................@@.........
.........@@...........:--=++*####+=---=+*##%%*=---=--==++====-========--=++-:--#@@@@@*=+####+=-:......................@@.........
.........@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@.........
.........@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@.........
.................................................................................................................................
.................................................................................................................................
.................................................................................................................................
.................................................................................................................................
.................................................................................................................................
.................................................................................................................................
.................................................................................................................................
.................................................................................................................................
.................................................................................................................................
"@
    Write-Host $perdangaArt -ForegroundColor Magenta
    Write-Host "Press any key to return..." -ForegroundColor DarkGray
    $null = Read-Host
}

# CORRECTED FUNCTION: Displays the selection menu, with robust centering for all console sizes.
function Show-Menu {
    Clear-Host

    # --- ASCII Art Definition ---
    $asciiArt = @(
       "__/\\\\\\\\\\\\\_______________________________________/\\\____________________________________________________________",
        " _\/\\\/////////\\\____________________________________\/\\\____________________________________________________________",
        "  _\/\\\_______\/\\\____________________________________\/\\\_________________________________/\\\\\\\\__________________",
        "   _\/\\\\\\\\\\\\\/_____/\\\\\\\\___/\\/\\\\\\\_________\/\\\___/\\\\\\\\\_____/\\/\\\\\\____/\\\////\\\__/\\\\\\\\\_____",
        "    _\/\\\/////////_____/\\\/////\\\_\/\\\/////\\\___/\\\\\\\\\__\////////\\\___\/\\\////\\\__\//\\\\\\\\\_\////////\\\____",
        "     _\/\\\_____________/\\\\\\\\\\\__\/\\\___\///___/\\\////\\\____/\\\\\\\\\\__\/\\\__\//\\\__\///////\\\___/\\\\\\\\\\___",
        "      _\/\\\____________\//\\///////___\/\\\_________\/\\\__\/\\\___/\\\/////\\\__\/\\\___\/\\\__/\\_____\\\__/\\\/////\\\___",
        "       _\/\\\_____________\//\\\\\\\\\\_\/\\\_________\//\\\\\\\/\\_\//\\\\\\\\/\\_\/\\\___\/\\\_\//\\\\\\\\__\//\\\\\\\\/\\__",
        "        _\///_______________\//////////__\///___________\///////\//___\////////\//__\///____\///___\////////____\////////\//___",
        "         __/\\\\\\\\\\\\\\\_____________________________________________________________________________________________________",
        "          _\/\\\///////////______________________________________________________________________________________________________",
        "           _\/\\\_________________________________________________________________________________________________________________",
        "            _\/\\\\\\\\\\\_____/\\\\\_____/\\/\\\\\\\______/\\\\\\\\___/\\\____/\\\_____/\\\\\\\\___/\//\\\\\\\____________________",
        "             _\/\\\///////____/\\\///\\\__\/\\\/////\\\___/\\\/////\\\_\//\\\__/\\\____/\\\/////\\\_\/\\\/////\\\___________________",
        "              _\/\\\__________/\\\__\//\\\_\/\\\___\///___/\\\\\\\\\\\___\//\\\/\\\____/\\\\\\\\\\\__\/\\\___\///____________________",
        "               _\/\\\_________\//\\\///\\\__\/\\\_________\//\\///////_____\//\\\\\____\//\\///////___\/\\\___________________________",
        "                _\/\\\__________\///\\\\\/___\/\\\__________\//\\\\\\\\\\____\//\\\-----\//\\\\\\\\\\_ \/\\\___________________________",
        "                 _\///_____________\/////_____\///____________\//////////______\///______\//////////____\///____________________________"
    )

    # --- Menu Content Generation ---
    $menuLines = New-Object System.Collections.Generic.List[string]
    $fixedMenuWidth = 80 
    $pssText = "Perdanga Software Solutions"
    $pssUnderline = "=" * $fixedMenuWidth
    $dashedLine = "-" * $fixedMenuWidth
    $fixedHeaderPadding = [math]::Floor(($fixedMenuWidth - $pssText.Length) / 2)
    if ($fixedHeaderPadding -lt 0) { $fixedHeaderPadding = 0 }
    $centeredPssTextLine = (" " * $fixedHeaderPadding) + $pssText

    $menuLines.Add($pssUnderline)
    $menuLines.Add($centeredPssTextLine)
    $menuLines.Add($dashedLine)
    $menuLines.Add(" Windows & Software Manager [PSS v1.5] ($(Get-Date -Format "dd.MM.yyyy HH:mm"))") 
    $menuLines.Add(" Log saved to: $(Split-Path -Leaf $script:logFile)")
    $menuLines.Add(" $($script:sortedPrograms.Count) programs available for installation")
    $menuLines.Add($pssUnderline)
    $menuLines.Add("")

    $programHeader = "Available Programs for Installation:"
    $centeredProgramHeader = (" " * (($fixedMenuWidth - $programHeader.Length) / 2)) + $programHeader
    $programUnderline = "-" * $fixedMenuWidth
    $menuLines.Add($centeredProgramHeader)
    $menuLines.Add($programUnderline)

    $formattedPrograms = @()
    $sortedDisplayNumbers = $script:numberToProgramMap.Keys | Sort-Object { [int]$_ }

    foreach ($dispNumber in $sortedDisplayNumbers) {
        $programName = $script:numberToProgramMap[$dispNumber]
        $formattedPrograms += "$($dispNumber). $($programName)"
    }

    $programColumns = @{ 0 = @(); 1 = @(); 2 = @() }
    $programsPerColumn = [math]::Ceiling($formattedPrograms.Count / 3.0) 

    for ($i = 0; $i -lt $formattedPrograms.Count; $i++) {
        if ($i -lt $programsPerColumn) { $programColumns[0] += $formattedPrograms[$i] }
        elseif ($i -lt ($programsPerColumn * 2)) { $programColumns[1] += $formattedPrograms[$i] }
        else { $programColumns[2] += $formattedPrograms[$i] }
    }

    $col1MaxLength = ($programColumns[0] | Measure-Object -Property Length -Maximum).Maximum
    $col2MaxLength = ($programColumns[1] | Measure-Object -Property Length -Maximum).Maximum

    if ($col1MaxLength -eq $null) {$col1MaxLength = 0}
    if ($col2MaxLength -eq $null) {$col2MaxLength = 0}

    $maxRows = [math]::Max($programColumns[0].Count, [math]::Max($programColumns[1].Count, $programColumns[2].Count))

    for ($i = 0; $i -lt $maxRows; $i++) {
        $line = "  "
        if ($i -lt $programColumns[0].Count) { $line += $programColumns[0][$i].PadRight($col1MaxLength + 4) }
        else { $line += "".PadRight($col1MaxLength + 4) }
        
        if ($i -lt $programColumns[1].Count) { $line += $programColumns[1][$i].PadRight($col2MaxLength + 4) }
        else { $line += "".PadRight($col2MaxLength + 4) }
        
        if ($i -lt $programColumns[2].Count) { $line += $programColumns[2][$i] }
        $menuLines.Add($line.TrimEnd())
    }
    $menuLines.Add("")

    $optionsHeader = "Select an Option:"
    $optionsUnderline = "-" * $fixedMenuWidth
    $centeredOptionsHeader = (" " * (($fixedMenuWidth - $optionsHeader.Length) / 2)) + $optionsHeader
    $menuLines.Add($centeredOptionsHeader)
    $menuLines.Add($optionsUnderline)
    
    $optionPairs = @(
        @{ Left = "[A] Install All Programs";              Right = "[W] Activate Windows" },
        @{ Left = "[G] Select Specific Programs [GUI]";    Right = "[N] Update Windows" },
        @{ Left = "[U] Uninstall Programs [GUI]";          Right = "[T] Disable Windows Telemetry" },
        @{ Left = "[C] Install Custom Program";            Right = "[S] System Cleanup [GUI]" },
        @{ Left = "[X] Activate Spotify";                  Right = "[F] Create Unattend.xml File [GUI]" },
        @{ Left = "[P] Import & Install from File";        Right = "[I] Show System Information" }
    )

    $column1Width = ($optionPairs.Left | Measure-Object -Property Length -Maximum).Maximum + 5

    foreach ($pair in $optionPairs) {
        $leftColumn = $pair.Left.PadRight($column1Width)
        $fullLine = "{0}{1}" -f $leftColumn, $pair.Right
        $menuLines.Add($fullLine.TrimEnd())
    }

    $menuLines.Add("")

    $exitLine = "[E] Exit Script"
    $exitPadding = [math]::Floor(($fixedMenuWidth - $exitLine.Length) / 2)
    if ($exitPadding -lt 0) { $exitPadding = 0 }
    $centeredExitLine = (" " * $exitPadding) + $exitLine
    $menuLines.Add($centeredExitLine)
    
    $menuLines.Add($optionsUnderline)

    # --- Calculate Padding (REVISED for dynamic console width) ---
    $consoleWidth = 0
    try {
        # This is the most reliable way to get the current window width.
        $consoleWidth = [System.Console]::WindowWidth
    } catch {}

    if ($consoleWidth -le 1) { 
        try {
            # Fallback for environments where the above fails (like some ISE versions).
            $consoleWidth = $Host.UI.RawUI.WindowSize.Width
        } catch {}
    }
    if ($consoleWidth -le 1) {
        # Final fallback to a default width if all else fails.
        $consoleWidth = 120 
    }

    $menuPaddingValue = [math]::Floor(($consoleWidth - $fixedMenuWidth) / 2)
    if ($menuPaddingValue -lt 0) { $menuPaddingValue = 0 }
    $menuPaddingString = " " * $menuPaddingValue

    $artWidth = ($asciiArt | Measure-Object -Property Length -Maximum).Maximum
    $artPaddingValue = [math]::Floor(($consoleWidth - $artWidth) / 2)
    if ($artPaddingValue -lt 0) { $artPaddingValue = 0 }
    $artPaddingString = " " * $artPaddingValue


    # --- Display Logic ---
    if ($script:firstRun) {
        # --- Animated Reveal on First Run ---
        foreach ($line in $asciiArt) {
            Write-Host ($artPaddingString + $line) -ForegroundColor Cyan
            Start-Sleep -Milliseconds 10
        }
        
        foreach ($lineEntry in $menuLines) {
            $trimmedEntry = $lineEntry.Trim()
            $color = if ($trimmedEntry -eq $pssText -or $trimmedEntry -like ($pssUnderline.Trim()) -or $trimmedEntry -like ($dashedLine.Trim()) -or $trimmedEntry -eq $programHeader -or $trimmedEntry -eq $optionsHeader -or $trimmedEntry -like ($programUnderline.Trim()) -or $trimmedEntry -like ($optionsUnderline.Trim())) { "Cyan" } else { "White" }
            Write-Host ($menuPaddingString + $lineEntry) -ForegroundColor $color
            Start-Sleep -Milliseconds 15
        }
        $script:firstRun = $false
    } else {
        # --- Instant Display for Subsequent Runs ---
        foreach ($line in $asciiArt) {
            Write-Host ($artPaddingString + $line) -ForegroundColor Cyan
        }
        
        foreach ($lineEntry in $menuLines) {
            $trimmedEntry = $lineEntry.Trim()
            $color = if ($trimmedEntry -eq $pssText -or $trimmedEntry -like ($pssUnderline.Trim()) -or $trimmedEntry -like ($dashedLine.Trim()) -or $trimmedEntry -eq $programHeader -or $trimmedEntry -eq $optionsHeader -or $trimmedEntry -like ($programUnderline.Trim()) -or $trimmedEntry -like ($optionsUnderline.Trim())) { "Cyan" } else { "White" }
            Write-Host ($menuPaddingString + $lineEntry) -ForegroundColor $color
        }
    }
    
    Write-Host "" 
    $promptTextForOneLine = "Enter option, single number, or list of numbers:"
    $promptInternalPadding = [math]::Floor(($fixedMenuWidth - $promptTextForOneLine.Length) / 2)
    if ($promptInternalPadding -lt 0) { $promptInternalPadding = 0 }
    $centeredPromptText = (" " * $promptInternalPadding) + $promptTextForOneLine
    
    Write-Host ($menuPaddingString + $centeredPromptText) -NoNewline -ForegroundColor Yellow
}

Write-LogAndHost "Core functions loaded successfully." -NoHost


# ================================================================================
#                               PART 3: CONFIGURATION
# ================================================================================

# --- PROGRAM DEFINITIONS ---
# Edit this list to add or remove programs.
# Use the exact Chocolatey package ID.
$programs = @(
    "7zip.install",
    "brave",
    "cursoride",
    "discord",
    "file-converter",
    "git",
    "googlechrome",
    "imageglass",
    "nilesoft-shell",
    "nvidia-app",
    "occt",
    "obs-studio",
    "qbittorrent",
    "revo-uninstaller",
    "spotify",
    "steam",
    "telegram",
    "vcredist-all",
    "vlc",
    "winrar",
    "wiztree"
)

# --- SCRIPT-WIDE VARIABLES ---
$script:sortedPrograms = $programs | Sort-Object
$script:mainMenuLetters = @('a', 'c', 'e', 'f', 'g', 'i', 'n', 'p', 's', 't', 'u', 'w', 'x')
$script:mainMenuRegexPattern = "^(perdanga|" + ($script:mainMenuLetters -join '|') + "|[0-9,\s]+)$"
$script:availableProgramNumbers = 1..($script:sortedPrograms.Count) | ForEach-Object { $_.ToString() }
$script:programToNumberMap = @{}
$script:numberToProgramMap = @{}

# Dynamically create maps for selecting programs by number.
for ($i = 0; $i -lt $script:sortedPrograms.Count; $i++) {
    $assignedNumber = $script:availableProgramNumbers[$i]
    $programName = $script:sortedPrograms[$i]
    $script:programToNumberMap[$programName] = $assignedNumber
    $script:numberToProgramMap[$assignedNumber] = $programName
}

Write-LogAndHost "Configuration loaded. $($script:sortedPrograms.Count) programs defined." -NoHost


# ================================================================================
#                                PART 4: MAIN SCRIPT
# ================================================================================

# --- INITIAL CHECKS & SETUP ---
if ($PSVersionTable.PSVersion -eq $null) {
    Write-Host "ERROR: This script must be run in PowerShell." -ForegroundColor Red
    exit 1
}

# Ensure correct encoding for output.
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Attempt to resize the console window and buffer for a better viewing experience.
try {
    $Host.UI.RawUI.BufferSize = New-Object Management.Automation.Host.Size(150, 3000)
    $Host.UI.RawUI.WindowSize = New-Object Management.Automation.Host.Size(150, 50)
} catch {
    Write-LogAndHost "Could not set console buffer or window size. Error: $($_.Exception.Message)" -ForegroundColor Yellow -LogPrefix "Initial-Setup"
}

# --- SCRIPT-WIDE CHECKS ---
$script:activationAttempted = $false
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-LogAndHost "This script must be run as Administrator." -ForegroundColor Red -LogPrefix "Startup-Check"
    "[$((Get-Date))] [Startup-Check] ERROR: Script not run as Administrator." | Out-File -FilePath $script:logFile -Append -Encoding UTF8
    exit 1
}

# Check if Windows Forms is available for GUI features.
$script:guiAvailable = $true
try { Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop }
catch { 
    $script:guiAvailable = $false
    Write-LogAndHost "GUI features are not available. Error: $($_.Exception.Message)" -HostColor Yellow -LogPrefix "Startup-Check"
}


# --- CHOCOLATEY INITIALIZATION ---
Write-LogAndHost "Checking Chocolatey installation..."
try {
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        if (-not (Install-Chocolatey)) { Write-LogAndHost "Chocolatey is required to proceed. Exiting script." -HostColor Red -LogPrefix "Choco-Init"; exit 1 }
        if (-not (Get-Command choco -ErrorAction SilentlyContinue)) { Write-LogAndHost "Chocolatey command not found after installation. Please install manually." -HostColor Red -LogPrefix "Choco-Init"; exit 1 }
        if (-not $env:ChocolateyInstall) { $env:ChocolateyInstall = "$($env:ProgramData)\chocolatey"; Write-LogAndHost "ChocolateyInstall environment variable set to: $env:ChocolateyInstall" -NoHost }
    } else {
         if (-not $env:ChocolateyInstall) { $env:ChocolateyInstall = "$($env:ProgramData)\chocolatey"; Write-LogAndHost "ChocolateyInstall environment variable set to: $env:ChocolateyInstall" -NoHost }
    }
    $chocoVersion = & choco --version 2>&1
    if ($LASTEXITCODE -ne 0) { Write-LogAndHost "Chocolatey is not functioning correctly. Exit code: $LASTEXITCODE" -HostColor Red -LogPrefix "Choco-Init"; exit 1 }
    Write-LogAndHost "Found Chocolatey version: $($chocoVersion -join ' ')"
}
catch { Write-LogAndHost "Exception occurred while checking Chocolatey. $($_.Exception.Message)" -HostColor Red -LogPrefix "Choco-Init"; exit 1 }

Write-Host ""; Write-LogAndHost "Enabling Chocolatey features for a smoother experience..."
try {
    # Enable global confirmation to prevent prompts during installations.
    & choco feature enable -n allowGlobalConfirmation 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8
    if ($LASTEXITCODE -eq 0) { Write-LogAndHost "Automatic confirmation enabled." } else { Write-LogAndHost "Failed to enable automatic confirmation. $($LASTEXITCODE)" -HostColor Yellow -LogPrefix "Choco-Init" }
} catch { Write-LogAndHost "Exception enabling automatic confirmation. $($_.Exception.Message)" -HostColor Red -LogPrefix "Choco-Init" }
Write-Host ""

# Set a flag to run the menu animation only once per session.
$script:firstRun = $true

# --- MAIN LOOP ---
do {
    Show-Menu
    try { $userInput = Read-Host } catch { Write-LogAndHost "Could not read user input. $($_.Exception.Message)" -HostColor Red -LogPrefix "Main-Loop"; Start-Sleep -Seconds 2; continue }
    $userInput = $userInput.Trim().ToLower()

    # --- EASTER EGG CHECK ---
    if ($userInput -eq 'perdanga') {
        Show-PerdangaArt
        continue
    }

    if ([string]::IsNullOrEmpty($userInput)) {
        Clear-Host; Write-LogAndHost "No input detected. Please enter an option." -HostColor Yellow
        Write-LogAndHost "Press any key to return to the menu..." -HostColor DarkGray -NoLog; $null = Read-Host; continue
    }
    
    # --- PROCESS USER INPUT ---
    # Case 1: A single letter command was entered.
    if ($script:mainMenuLetters -contains $userInput) {
        switch ($userInput) {
            'e' {
                Write-LogAndHost "Exiting script..."
                try { & choco feature disable -n allowGlobalConfirmation 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8 } catch { Write-LogAndHost "Exception disabling auto-confirm - $($_.Exception.Message)" -HostColor Red -LogPrefix "Main-Loop" }
                exit 0
            }
            'a' {
                Clear-Host
                try {
                    Write-LogAndHost "Are you sure you want to install all programs? (y/n)" -HostColor Yellow -LogPrefix "Main-Loop"
                    $confirmInput = Read-Host
                } catch { Write-LogAndHost "Could not read user input." -HostColor Red -LogPrefix "Main-Loop"; Start-Sleep -Seconds 2; continue }
                if ($confirmInput.Trim().ToLower() -eq 'y') {
                    Write-LogAndHost "User chose to install all programs." -NoHost; Clear-Host
                    if (Install-Programs -ProgramsToInstall $script:sortedPrograms) { Write-LogAndHost "All programs installation process completed." } else { Write-LogAndHost "Some programs may not have installed correctly. Check log." -HostColor Yellow }
                    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host
                } else { Write-LogAndHost "Installation of all programs cancelled." }
            }
            'g' {
                if ($script:guiAvailable) {
                    Write-LogAndHost "User chose GUI-based installation." -NoHost
                    $form = New-Object System.Windows.Forms.Form; $form.Text = "Perdanga GUI - Install Programs"; $form.Size = New-Object System.Drawing.Size(400, 450); $form.StartPosition = "CenterScreen"; $form.FormBorderStyle = "FixedDialog"; $form.MaximizeBox = $false; $form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
                    $panel = New-Object System.Windows.Forms.Panel; $panel.Size = New-Object System.Drawing.Size(360, 350); $panel.Location = New-Object System.Drawing.Point(10, 10); $panel.AutoScroll = $true; $panel.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48); $form.Controls.Add($panel)
                    $checkboxes = @(); $yPos = 10
                    for ($i = 0; $i -lt $script:sortedPrograms.Length; $i++) {
                        $progName = $script:sortedPrograms[$i]; $dispNumber = $script:programToNumberMap[$progName]; $displayText = "$($dispNumber). $progName".PadRight(30)
                        $checkbox = New-Object System.Windows.Forms.CheckBox; $checkbox.Text = $displayText; $checkbox.Location = New-Object System.Drawing.Point(10, $yPos); $checkbox.Size = New-Object System.Drawing.Size(330, 24); $checkbox.Font = New-Object System.Drawing.Font("Segoe UI", 10); $checkbox.ForeColor = [System.Drawing.Color]::White; $checkbox.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48); $panel.Controls.Add($checkbox); $checkboxes += $checkbox; $yPos += 28
                    }
                    $okButton = New-Object System.Windows.Forms.Button; $okButton.Text = "Install Selected"; $okButton.Location = New-Object System.Drawing.Point(140, 370); $okButton.Size = New-Object System.Drawing.Size(120, 30); $okButton.Font = New-Object System.Drawing.Font("Segoe UI", 10); $okButton.ForeColor = [System.Drawing.Color]::White; $okButton.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180); $okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $okButton.FlatAppearance.BorderSize = 0; $okButton.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() }); $form.Controls.Add($okButton)
                    Clear-Host
                    $result = $form.ShowDialog()
                    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                        $selectedProgramsFromGui = @(); for ($i = 0; $i -lt $checkboxes.Length; $i++) { if ($checkboxes[$i].Checked) { $selectedProgramsFromGui += $script:sortedPrograms[$i] } }
                        if ($selectedProgramsFromGui.Count -eq 0) { Write-LogAndHost "No programs selected via GUI for installation." }
                        else { Clear-Host; if (Install-Programs -ProgramsToInstall $selectedProgramsFromGui) { Write-LogAndHost "Selected programs installation process completed." } else { Write-LogAndHost "Some GUI selected programs failed to install." -HostColor Yellow } }
                        Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host
                    } else { Write-LogAndHost "GUI installation cancelled by user." }
                } else { Clear-Host; Write-LogAndHost "GUI selection (g) is not available." -HostColor Red -LogPrefix "Main-Loop"; Start-Sleep -Seconds 2; }
            }
            'u' {
                if ($script:guiAvailable) {
                    Write-LogAndHost "User chose GUI-based uninstallation." -NoHost
                    $installedChocoPackages = Get-InstalledChocolateyPackages
                    if ($installedChocoPackages.Count -eq 0) { Clear-Host; Write-LogAndHost "No Chocolatey packages found to uninstall." -HostColor Yellow; Start-Sleep -Seconds 2; continue }
                    $form = New-Object System.Windows.Forms.Form; $form.Text = "Perdanga GUI - Uninstall Programs"; $form.Size = New-Object System.Drawing.Size(400, 450); $form.StartPosition = "CenterScreen"; $form.FormBorderStyle = "FixedDialog"; $form.MaximizeBox = $false; $form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
                    $panel = New-Object System.Windows.Forms.Panel; $panel.Size = New-Object System.Drawing.Size(360, 350); $panel.Location = New-Object System.Drawing.Point(10, 10); $panel.AutoScroll = $true; $panel.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48); $form.Controls.Add($panel)
                    $checkboxes = @(); $yPos = 10
                    foreach ($packageName in ($installedChocoPackages | Sort-Object)) {
                        $checkbox = New-Object System.Windows.Forms.CheckBox; $checkbox.Text = $packageName; $checkbox.Location = New-Object System.Drawing.Point(10, $yPos); $checkbox.Size = New-Object System.Drawing.Size(330, 24); $checkbox.Font = New-Object System.Drawing.Font("Segoe UI", 10); $checkbox.ForeColor = [System.Drawing.Color]::White; $checkbox.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48); $panel.Controls.Add($checkbox); $checkboxes += $checkbox; $yPos += 28
                    }
                    $okButton = New-Object System.Windows.Forms.Button; $okButton.Text = "Uninstall Selected"; $okButton.Location = New-Object System.Drawing.Point(130, 370); $okButton.Size = New-Object System.Drawing.Size(140, 30); $okButton.Font = New-Object System.Drawing.Font("Segoe UI", 10); $okButton.ForeColor = [System.Drawing.Color]::White; $okButton.BackColor = [System.Drawing.Color]::FromArgb(180, 70, 70); $okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $okButton.FlatAppearance.BorderSize = 0; $okButton.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() }); $form.Controls.Add($okButton)
                    Clear-Host
                    $result = $form.ShowDialog()
                    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                        $selectedProgramsToUninstall = @(); foreach ($cb in $checkboxes) { if ($cb.Checked) { $selectedProgramsToUninstall += $cb.Text } }
                        if ($selectedProgramsToUninstall.Count -eq 0) { Write-LogAndHost "No programs selected via GUI for uninstallation." }
                        else { Clear-Host; if (Uninstall-Programs -ProgramsToUninstall $selectedProgramsToUninstall) { Write-LogAndHost "Selected programs uninstallation process completed." } else { Write-LogAndHost "Some GUI selected programs failed to uninstall." -HostColor Yellow } }
                        Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host
                    } else { Write-LogAndHost "GUI uninstallation cancelled by user." }
                } else { Clear-Host; Write-LogAndHost "GUI uninstallation (u) is not available." -HostColor Red -LogPrefix "Main-Loop"; Start-Sleep -Seconds 2; }
            }
            'c' {
                Clear-Host
                $customPackageName = ""
                try {
                    $customPackageName = Read-Host "Enter the exact Chocolatey package ID (e.g., 'notepadplusplus.install', 'git')"
                    $customPackageName = $customPackageName.Trim()
                } catch {
                    Write-LogAndHost "Could not read user input for custom package name. $($_.Exception.Message)" -HostColor Red -LogPrefix "Main-Loop"
                    Start-Sleep -Seconds 2
                    continue
                }

                if ([string]::IsNullOrWhiteSpace($customPackageName)) {
                    Write-LogAndHost "No package name entered. Returning to menu." -HostColor Yellow
                    Start-Sleep -Seconds 2
                    continue
                }
                
                Write-LogAndHost "Checking if package '$customPackageName' exists..."
                if (Test-ChocolateyPackage -PackageName $customPackageName) {
                    Write-LogAndHost "Package '$customPackageName' found. Proceed with installation?" -HostColor Yellow
                    try {
                        $confirmInstallCustom = Read-Host "(Type y/n then press Enter)"
                        if ($confirmInstallCustom.Trim().ToLower() -eq 'y') {
                            Write-LogAndHost "User confirmed installation of custom package '$customPackageName'." -NoHost
                            Clear-Host
                            if (Install-Programs -ProgramsToInstall @($customPackageName)) {
                                Write-LogAndHost "Custom package '$customPackageName' installation process completed."
                            } else {
                                Write-LogAndHost "Failed to install custom package '$customPackageName'. Check log for details." -HostColor Red
                            }
                        } else {
                            Write-LogAndHost "Installation of custom package '$customPackageName' cancelled by user." -HostColor Yellow
                        }
                    } catch {
                         Write-LogAndHost "Could not read user input for custom package installation confirmation. $($_.Exception.Message)" -HostColor Red -LogPrefix "Main-Loop"
                         Start-Sleep -Seconds 2
                    }
                } else {
                    Write-LogAndHost "Custom package '$customPackageName' could not be installed (either not found or validation failed)." -HostColor Red
                }
                Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
                $null = Read-Host
            }
            't' { Clear-Host; Invoke-DisableTelemetry }
            'x' { Clear-Host; Invoke-SpotXActivation }
            'w' { Clear-Host; Invoke-WindowsActivation }
            'n' {
                Clear-Host
                try {
                    Write-LogAndHost "This will install Windows Updates. Proceed? (y/n)" -HostColor Yellow -LogPrefix "Main-Loop"
                    $confirmInput = Read-Host
                } catch { Write-LogAndHost "Could not read user input." -HostColor Red -LogPrefix "Main-Loop"; Start-Sleep -Seconds 2; continue }
                if ($confirmInput.Trim().ToLower() -eq 'y') { Clear-Host; Invoke-WindowsUpdate; Write-LogAndHost "Windows Update process finished." } else { Write-LogAndHost "Update process cancelled." }
            }
            'f' { Clear-Host; Create-UnattendXml }
            's' { Clear-Host; Invoke-TempFileCleanup }
            'i' { Clear-Host; Show-SystemInfo }
            'p' { Clear-Host; Import-ProgramSelection }
        }
    }
    # Case 2: A list of numbers (separated by commas or spaces) was entered.
    elseif ($userInput -match '[, ]+') {
        Clear-Host
        $selectedIndividualInputs = $userInput -split '[, ]+' | ForEach-Object { $_.Trim() } | Where-Object {$_ -ne ""}
        $validProgramNamesToInstall = New-Object System.Collections.Generic.List[string]
        $invalidNumbersInList = New-Object System.Collections.Generic.List[string]
        
        foreach ($inputNumStr in $selectedIndividualInputs) {
            if ($inputNumStr -match '^\d+$' -and $script:numberToProgramMap.ContainsKey($inputNumStr)) {
                $programName = $script:numberToProgramMap[$inputNumStr]
                if (-not $validProgramNamesToInstall.Contains($programName)) { $validProgramNamesToInstall.Add($programName) }
            } elseif ($inputNumStr -ne "") { $invalidNumbersInList.Add($inputNumStr) }
        }
        
        if ($validProgramNamesToInstall.Count -eq 0) {
            Write-LogAndHost "No valid program numbers found in your input: '$userInput'." -HostColor Red
            if ($invalidNumbersInList.Count -gt 0) { Write-LogAndHost "Unrecognized inputs: $($invalidNumbersInList -join ', ')" -HostColor Red -NoLog }
            Start-Sleep -Seconds 2
        } else {
            Write-LogAndHost "Selected for installation: $($validProgramNamesToInstall -join ', ')" -HostColor Cyan -NoLog
            if ($invalidNumbersInList.Count -gt 0) { Write-LogAndHost "Invalid/skipped inputs: $($invalidNumbersInList -join ', ')" -HostColor Yellow -NoLog }
            try {
                Write-LogAndHost "Install these $($validProgramNamesToInstall.Count) program(s)? (Type y/n then press Enter)" -HostColor Yellow
                $confirmMultiInput = Read-Host
            } catch { Write-LogAndHost "Could not read user input. $($_.Exception.Message)" -HostColor Red -LogPrefix "Main-Loop"; Start-Sleep -Seconds 2; continue }
            
            if ($confirmMultiInput.Trim().ToLower() -eq 'y') {
                Clear-Host
                if (Install-Programs -ProgramsToInstall $validProgramNamesToInstall) { Write-LogAndHost "Selected programs installation process completed." }
                else { Write-LogAndHost "Some selected programs failed. Check log for details." -HostColor Yellow }
                Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host
            } else { Write-LogAndHost "Installation of selected programs cancelled." }
        }
    }
    # Case 3: A single, valid program number was entered.
    elseif ($script:numberToProgramMap.ContainsKey($userInput)) {
        $programToInstall = $script:numberToProgramMap[$userInput]
        Clear-Host
        try {
            Write-LogAndHost "Install '$($programToInstall)' (program #$($userInput))? (Type y/n then press Enter)" -HostColor Yellow
            $confirmSingleInput = Read-Host
        } catch { Write-LogAndHost "Could not read user input. $($_.Exception.Message)" -HostColor Red -LogPrefix "Main-Loop"; Start-Sleep -Seconds 2; continue }
        
        if ($confirmSingleInput.Trim().ToLower() -eq 'y') {
            Clear-Host
            if (Install-Programs -ProgramsToInstall @($programToInstall)) { Write-LogAndHost "$($programToInstall) installation process completed." }
            else { Write-LogAndHost "Failed to install $($programToInstall). Check log for details." -HostColor Red }
            Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host
        } else { Write-LogAndHost "Installation of '$($programToInstall)' cancelled." }
    }
    # Case 4: The input was invalid.
    else {
        Clear-Host
        $validOptions = ($script:mainMenuLetters | Sort-Object | ForEach-Object { $_.ToUpper() }) -join ','
        $errorMessage = "Invalid input: '$userInput'. Use options [$validOptions], program numbers, or a secret word."
        Write-LogAndHost $errorMessage -HostColor Red -LogPrefix "Main-Loop"
        Start-Sleep -Seconds 2
    }
} while ($true)
