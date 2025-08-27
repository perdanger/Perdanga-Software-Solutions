<#
.SYNOPSIS
    Author: Roman Zhdanov
    Version: 1.6
    Last Modified: 11.08.2025
.DESCRIPTION
    Perdanga Software Solutions is a PowerShell script designed to simplify the installation, 
    uninstallation, and management of essential Windows software. Includes dynamic application
    cache cleaning.

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
            Write-LogAndHost "Chocolatey installation declined by user. Proceeding without Chocolatey." -HostColor Yellow
            return $true
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

# ENHANCED FUNCTION: Displays key system information in a graphical, multi-panel layout.
function Show-SystemInfo {
    if (-not $script:guiAvailable) {
        Write-LogAndHost "GUI is not available, cannot launch the System Information tool." -HostColor Red -LogPrefix "Show-SystemInfo"
        Start-Sleep -Seconds 2
        return
    }

    # Use a Gemini-themed color for the launch message
    Write-LogAndHost "Launching System Information GUI..." -HostColor Cyan -LogPrefix "Show-SystemInfo"
    
    # --- GUI Setup ---
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Perdanga System Information"
    $form.Size = New-Object System.Drawing.Size(900, 750)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false

    # Gemini Theme Colors
    # A palette inspired by the Gemini logo: Dark, Blue, Yellow
    $geminiDarkBg = [System.Drawing.Color]::FromArgb(20, 20, 25)
    $geminiPanelBg = [System.Drawing.Color]::FromArgb(35, 35, 40)
    $geminiBlue = [System.Drawing.Color]::FromArgb(60, 100, 180)    # Muted Blue
    $geminiYellow = [System.Drawing.Color]::FromArgb(230, 180, 50)  # Golden Yellow
    $geminiAccent = [System.Drawing.Color]::FromArgb(0, 200, 255)   # Light Blue/Cyan accent for headers
    $geminiGrayText = [System.Drawing.Color]::Gainsboro # Soft gray for labels
    $geminiWhiteText = [System.Drawing.Color]::White    # White for values

    $form.BackColor = $geminiDarkBg
    $form.Opacity = 0

    # --- Control Styles ---
    $commonFont = New-Object System.Drawing.Font("Segoe UI", 9)
    $groupboxFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $labelColor = $geminiGrayText
    $valueColor = $geminiWhiteText
    
    # --- Main Layout Panel ---
    $mainTableLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $mainTableLayout.Dock = "Fill"
    $mainTableLayout.ColumnCount = 2
    $mainTableLayout.RowCount = 5 
    $mainTableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    $mainTableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    $mainTableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
    $mainTableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
    $mainTableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
    $mainTableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
    $mainTableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50))) | Out-Null
    $mainTableLayout.BackColor = $geminiDarkBg
    $form.Controls.Add($mainTableLayout)

    # --- Helper function to create styled GroupBoxes with a unique color for each ---
    function New-InfoGroupBox($Text, $HeaderColor) {
        $groupbox = New-Object System.Windows.Forms.GroupBox
        $groupbox.Text = $Text
        $groupbox.Font = $groupboxFont
        $groupbox.ForeColor = $HeaderColor
        $groupbox.Dock = "Fill"
        $groupbox.Padding = New-Object System.Windows.Forms.Padding(10, 20, 10, 10)
        $groupbox.Margin = New-Object System.Windows.Forms.Padding(10)
        $groupbox.BackColor = $geminiPanelBg # Consistent background for all info boxes

        # Create a nested panel with AutoScroll to handle overflow
        $scrollPanel = New-Object System.Windows.Forms.Panel
        $scrollPanel.Dock = "Fill"
        $scrollPanel.AutoScroll = $true
        $scrollPanel.BackColor = $geminiPanelBg
        $groupbox.Controls.Add($scrollPanel)

        # Create a nested TableLayoutPanel for a clean, two-column layout
        $tlp = New-Object System.Windows.Forms.TableLayoutPanel
        $tlp.Dock = "Top"
        $tlp.AutoSize = $true
        $tlp.ColumnCount = 2
        $tlp.BackColor = $geminiPanelBg
        
        # Modified column styles for proper alignment
        $tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
        $tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
        
        $scrollPanel.Controls.Add($tlp)

        return $groupbox
    }

    # --- Create all GroupBoxes with a more restricted, cohesive color palette ---
    $gbOs = New-InfoGroupBox "Operating System" $geminiBlue
    $gbCpu = New-InfoGroupBox "Processor" $geminiYellow
    $gbRam = New-InfoGroupBox "Memory (RAM)" $geminiBlue
    $gbHardware = New-InfoGroupBox "System Hardware" $geminiYellow
    $gbGpu = New-InfoGroupBox "Video Card(s)" $geminiBlue
    $gbNetwork = New-InfoGroupBox "Network Adapters" $geminiYellow
    $gbDisk = New-InfoGroupBox "Disk Drives" $geminiBlue

    # --- Add GroupBoxes to the layout ---
    # Column 0 (Left)
    $mainTableLayout.Controls.Add($gbOs, 0, 0)
    $mainTableLayout.Controls.Add($gbCpu, 0, 1)
    $mainTableLayout.Controls.Add($gbRam, 0, 2)
    $mainTableLayout.Controls.Add($gbHardware, 0, 3)

    # Column 1 (Right)
    $mainTableLayout.Controls.Add($gbGpu, 1, 0)
    $mainTableLayout.Controls.Add($gbNetwork, 1, 1)
    $mainTableLayout.Controls.Add($gbDisk, 1, 2)
    $mainTableLayout.SetRowSpan($gbDisk, 2)
    
    # --- Button Panel ---
    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonPanel.Dock = "Fill"
    $buttonPanel.FlowDirection = "LeftToRight"
    $buttonPanel.Padding = New-Object System.Windows.Forms.Padding(10, 10, 0, 0)
    $buttonPanel.BackColor = $geminiDarkBg
    $mainTableLayout.SetColumnSpan($buttonPanel, 2)
    $mainTableLayout.Controls.Add($buttonPanel, 0, 4)

    # --- Helper function to create styled Buttons with a consistent theme ---
    function New-ActionButton($Text, $BackColor) {
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $Text
        $button.Size = "130,30"
        $button.Font = $commonFont
        $button.ForeColor = $geminiWhiteText
        $button.BackColor = $BackColor
        $button.FlatStyle = "Flat"
        $button.FlatAppearance.BorderSize = 0
        return $button
    }

    $buttonCopy = New-ActionButton "Copy to Clipboard" $geminiYellow
    $buttonRefresh = New-ActionButton "Refresh" $geminiBlue
    
    $buttonPanel.Controls.AddRange(@($buttonCopy, $buttonRefresh))

    # --- Data Population Logic ---
    $script:infoStore = @{} # Store raw data for clipboard
    
    function Update-SystemInfo {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $script:infoStore.Clear()
        
        # Helper to create and add labels to a groupbox's inner TableLayoutPanel
        $populateGroupBox = {
            param($GroupBox, $Data)
            $tlp = $GroupBox.Controls[0].Controls[0] # Get the nested TableLayoutPanel inside the scroll panel
            $tlp.Controls.Clear()
            $tlp.RowCount = 0
            
            foreach ($item in $Data.GetEnumerator()) {
                $label = New-Object System.Windows.Forms.Label; $label.Text = $item.Key; $label.Font = $commonFont; $label.ForeColor = $labelColor; $label.AutoSize = $true
                $value = New-Object System.Windows.Forms.Label; $value.Text = $item.Value; $value.Font = $commonFont; $value.ForeColor = $valueColor; $value.AutoSize = $true
                
                $tlp.Controls.Add($label, 0, $tlp.RowCount)
                $tlp.Controls.Add($value, 1, $tlp.RowCount)
                
                $tlp.RowCount += 1 # Increment the row count
                $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
            }
        }
        
        try {
            # --- OS Information ---
            $osData = [ordered]@{}
            $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
            $osData["Name:"] = $osInfo.Caption; $script:infoStore["OS Name"] = $osInfo.Caption
            $osData["Version:"] = $osInfo.Version; $script:infoStore["OS Version"] = $osInfo.Version
            $osData["Build:"] = $osInfo.BuildNumber; $script:infoStore["OS Build"] = $osInfo.BuildNumber
            $osData["Architecture:"] = $osInfo.OSArchitecture; $script:infoStore["OS Architecture"] = $osInfo.OSArchitecture
            try { $productID = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductId; $osData["Product ID:"] = $productID; $script:infoStore["Product ID"] = $productID } catch {}
            & $populateGroupBox $gbOs $osData
            
            # --- CPU Information ---
            $cpuData = [ordered]@{}
            $cpuInfo = Get-CimInstance -ClassName Win32_Processor
            $cpuData["Name:"] = $cpuInfo.Name.Trim(); $script:infoStore["CPU"] = $cpuInfo.Name.Trim()
            $cpuData["Cores (Logical):"] = "$($cpuInfo.NumberOfCores) ($($cpuInfo.NumberOfLogicalProcessors))"; $script:infoStore["Cores"] = "$($cpuInfo.NumberOfCores) ($($cpuInfo.NumberOfLogicalProcessors))"
            $virtEnabled = if ($cpuInfo.VirtualizationFirmwareEnabled) { "Enabled" } else { "Disabled" }
            $cpuData["Virtualization:"] = "Firmware $virtEnabled"; $script:infoStore["Virtualization"] = "Firmware $virtEnabled"
            & $populateGroupBox $gbCpu $cpuData

            # --- System Hardware ---
            $hwData = [ordered]@{}
            $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
            $hwData["Manufacturer:"] = $computerSystem.Manufacturer; $script:infoStore["Manufacturer"] = $computerSystem.Manufacturer
            $hwData["Model:"] = $computerSystem.Model; $script:infoStore["Model"] = $computerSystem.Model
            $boardInfo = Get-CimInstance -ClassName Win32_BaseBoard
            $hwData["Motherboard:"] = "$($boardInfo.Manufacturer) $($boardInfo.Product)"; $script:infoStore["Motherboard"] = "$($boardInfo.Manufacturer) $($boardInfo.Product)"
            $biosInfo = Get-CimInstance -ClassName Win32_BIOS
            $hwData["BIOS Version:"] = $biosInfo.SMBIOSBIOSVersion; $script:infoStore["BIOS Version"] = $biosInfo.SMBIOSBIOSVersion
            
            # ADDED: Secure Boot Status
            $secureBootStatus = try { if (Confirm-SecureBootUEFI) { "Enabled" } else { "Disabled" } } catch { "Unsupported / Error" }
            $hwData["Secure Boot:"] = $secureBootStatus; $script:infoStore["Secure Boot"] = $secureBootStatus
            
            # ADDED: TPM Status
            $tpmStatus = "Not Found"
            try {
                $tpmInfo = Get-Tpm -ErrorAction SilentlyContinue
                if ($tpmInfo -and $tpmInfo.TpmPresent) {
                    if ($tpmInfo.TpmReady -and $tpmInfo.SpecificationVersion -eq "2.0") {
                        $tpmStatus = "Present & Ready (2.0)"
                    } elseif ($tpmInfo.SpecificationVersion -eq "2.0") {
                        $tpmStatus = "Present, Not Ready (2.0)"
                    } else {
                        $tpmStatus = "Present (Version < 2.0)"
                    }
                }
            } catch {
                $tpmStatus = "Unsupported / Error"
            }
            $hwData["TPM Status:"] = $tpmStatus; $script:infoStore["TPM Status"] = $tpmStatus

            & $populateGroupBox $gbHardware $hwData

            # --- Memory (RAM) ---
            $ramData = [ordered]@{}
            $ramGB = [math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2)
            $ramData["Total Installed:"] = "$($ramGB) GB"; $script:infoStore["Total RAM"] = "$($ramGB) GB"
            $ramData["Modules:"] = ""
            $memoryModules = Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction SilentlyContinue
            if ($memoryModules) {
                $i = 1
                foreach ($module in @($memoryModules)) {
                    $capacityGB = [math]::Round($module.Capacity / 1GB, 2)
                    $key = "- Slot $i ($($module.DeviceLocator)):"; $value = "$($capacityGB) GB, $($module.ConfiguredClockSpeed) MHz, $($module.Manufacturer)"
                    $ramData[$key] = $value; $script:infoStore["RAM Module $i"] = $value; $i++
                }
            }
            & $populateGroupBox $gbRam $ramData
            
            # --- Video Cards ---
            $gpuData = [ordered]@{}
            $videoControllers = Get-CimInstance -ClassName Win32_VideoController
            $i = 1
            foreach ($video in @($videoControllers)) {
                $gpuData["Name:"] = $video.Name; $script:infoStore["GPU $i Name"] = $video.Name
                $adapterRamGB = $null; $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\"
                $matchingKey = (Get-ChildItem -Path $regPath -EA SilentlyContinue) | ? { ($_.GetValue("DriverDesc") -eq $video.Name) -or ($_.GetValue("Description") -eq $video.Name) } | Select -First 1
                if ($matchingKey) { $vramBytes = $matchingKey.GetValue("HardwareInformation.qwMemorySize"); if ($vramBytes -gt 0) { $adapterRamGB = [math]::Round($vramBytes / 1GB, 2) } }
                if (-not $adapterRamGB -and $video.AdapterRAM) { $adapterRamGB = [math]::Round($video.AdapterRAM / 1GB, 2) }
                $gpuData["Adapter RAM:"] = "$($adapterRamGB) GB"; $script:infoStore["GPU $i VRAM"] = "$($adapterRamGB) GB"
                $gpuData["Driver Version:"] = $video.DriverVersion; $script:infoStore["GPU $i Driver"] = $video.DriverVersion
                $gpuData[" "] = ""; $i++
            }
            & $populateGroupBox $gbGpu $gpuData

            # --- Disk Information (UPDATED) ---
            $diskTlp = $gbDisk.Controls[0].Controls[0]
            $diskTlp.Controls.Clear()
            $diskTlp.RowCount = 0
            
            $disks = Get-CimInstance -ClassName Win32_DiskDrive
            foreach ($disk in @($disks)) {
                $sizeGB = [math]::Round($disk.Size / 1GB, 2)
                $diskType = (Get-PhysicalDisk -DeviceNumber $disk.Index -EA SilentlyContinue).MediaType
                
                # Create label for the disk
                $diskLabel = New-Object System.Windows.Forms.Label
                $diskLabel.Text = "$($disk.Model) ($($sizeGB) GB) - $diskType"
                $diskLabel.Font = $groupboxFont
                $diskLabel.ForeColor = $gbDisk.ForeColor # Use the same color as the GroupBox header
                $diskLabel.AutoSize = $true
                $diskTlp.Controls.Add($diskLabel, 0, $diskTlp.RowCount)
                $diskTlp.SetColumnSpan($diskLabel, 2)
                $diskTlp.RowCount += 1
                $diskTlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
                $script:infoStore["Disk $($disk.Index)"] = $diskLabel.Text

                $partitions = Get-CimAssociatedInstance -InputObject $disk -ResultClassName Win32_DiskPartition
                foreach ($partition in @($partitions)) {
                    $logicalDisk = Get-CimAssociatedInstance -InputObject $partition -ResultClassName Win32_LogicalDisk
                    if ($logicalDisk) {
                        $freeGB = [math]::Round($logicalDisk.FreeSpace / 1GB, 2)
                        $percentFree = [math]::Round(($logicalDisk.FreeSpace / $logicalDisk.Size) * 100, 2)
                        
                        # Create labels for the partition
                        $driveLabel = New-Object System.Windows.Forms.Label
                        $driveLabel.Text = "  - Drive $($logicalDisk.DeviceID):"
                        $driveLabel.Font = $commonFont
                        $driveLabel.ForeColor = $labelColor
                        $driveLabel.AutoSize = $true
                        
                        $freeSpaceLabel = New-Object System.Windows.Forms.Label
                        $freeSpaceLabel.Text = "Free: $($freeGB) GB ($($percentFree)%)"
                        $freeSpaceLabel.Font = $commonFont
                        $freeSpaceLabel.ForeColor = $valueColor
                        $freeSpaceLabel.AutoSize = $true
                        
                        $diskTlp.Controls.Add($driveLabel, 0, $diskTlp.RowCount)
                        $diskTlp.Controls.Add($freeSpaceLabel, 1, $diskTlp.RowCount)
                        $diskTlp.RowCount += 1
                        $diskTlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
                        $script:infoStore["Partition $($logicalDisk.DeviceID)"] = "Free: $($freeGB) GB ($($percentFree)%)"
                    }
                }
            }
            
            # --- Network Information ---
            $netData = [ordered]@{}
            $netAdapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled }
            $i = 1
            foreach ($adapter in @($netAdapters)) {
                $netData["Description:"] = $adapter.Description; $script:infoStore["NIC $i Desc"] = $adapter.Description
                $netData["IP Address:"] = ($adapter.IPAddress -join ', '); $script:infoStore["NIC $i IP"] = ($adapter.IPAddress -join ', ')
                $netData["MAC Address:"] = $adapter.MACAddress; $script:infoStore["NIC $i MAC"] = $adapter.MACAddress
                $netData["Default Gateway:"] = ($adapter.DefaultIPGateway -join ', '); $script:infoStore["NIC $i Gateway"] = ($adapter.DefaultIPGateway -join ', ')
                $netData["DNS Servers:"] = ($adapter.DNSServerSearchOrder -join ', '); $script:infoStore["NIC $i DNS"] = ($adapter.DNSServerSearchOrder -join ', ')
                $netData[" "] = ""; $i++
            }
            & $populateGroupBox $gbNetwork $netData

        } catch {
            Write-LogAndHost "Failed to gather system information. Error: $($_.Exception.Message)" -HostColor Red -LogPrefix "Show-SystemInfo"
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    # --- Button Event Handlers ---
    $buttonRefresh.Add_Click({ Update-SystemInfo })
    $buttonCopy.Add_Click({
        $clipboardText = ""
        $sections = @(
            "--- Operating System ---", "OS Name", "OS Version", "OS Build", "OS Architecture", "Product ID",
            "--- Processor ---", "CPU", "Cores", "Virtualization",
            "--- System Hardware ---", "Manufacturer", "Model", "Motherboard", "BIOS Version", "Secure Boot", "TPM Status",
            "--- Memory (RAM) ---", "Total RAM",
            "--- Video Card(s) ---",
            "--- Disk Drives ---",
            "--- Network Adapters ---"
        )
        foreach ($section in $sections) {
            if ($section.StartsWith("---")) { $clipboardText += "`r`n$section`r`n" }
            elseif ($script:infoStore.ContainsKey($section)) { $clipboardText += "$section`: $($script:infoStore[$section])`r`n" }
        }
        # Special handling for multi-entry sections
        ($script:infoStore.Keys | Where-Object { $_ -like "RAM Module *" } | Sort-Object) | ForEach-Object { $clipboardText += "  - $($script:infoStore[$_])`r`n" }
        ($script:infoStore.Keys | Where-Object { $_ -like "GPU * Name" } | Sort-Object) | ForEach-Object { 
            $gpuNum = $_.Split(' ')[1]; $clipboardText += "`r`nGPU: $($script:infoStore[$_])`r`n"; $clipboardText += "  VRAM: $($script:infoStore["GPU $gpuNum VRAM"])`r`n"; $clipboardText += "  Driver: $($script:infoStore["GPU $gpuNum Driver"])`r`n"
        }
        ($script:infoStore.Keys | Where-Object { $_ -like "Disk *" } | Sort-Object) | ForEach-Object { 
            $clipboardText += "`r`nDisk: $($script:infoStore[$_])`r`n"
            $diskIndex = $_.Split(' ')[1]
            ($script:infoStore.Keys | Where-Object { $_ -like "Partition*" } | Sort-Object) | ForEach-Object {
                 $clipboardText += "  - $($script:infoStore[$_])`r`n"
            }
        }
        ($script:infoStore.Keys | Where-Object { $_ -like "NIC * Desc" } | Sort-Object) | ForEach-Object { 
            $nicNum = $_.Split(' ')[1]; $clipboardText += "`r`nNIC: $($script:infoStore[$_])`r`n"; $clipboardText += "  IP: $($script:infoStore["NIC $nicNum IP"])`r`n"; $clipboardText += "  MAC: $($script:infoStore["NIC $nicNum MAC"])`r`n"
        }
        Set-Clipboard -Value $clipboardText.Trim()
        [System.Windows.Forms.MessageBox]::Show("System information copied to clipboard.", "Success", "OK", "Information") | Out-Null
    })

    # --- Initial Load and Show Form ---
    $form.Add_Shown({ $form.Opacity = 1 })
    Update-SystemInfo

    try {
        $null = $form.ShowDialog()
        Write-LogAndHost "System information GUI closed by user." -NoHost
    } catch {
        Write-LogAndHost "An unexpected error occurred with the System Information GUI. Details: $($_.Exception.Message)" -HostColor Red
    } finally {
        $form.Dispose()
    }

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

# ENHANCED FUNCTION: Dynamically finds potential cache folders for installed applications, including a list of well-known ones and installed browsers.
function Find-DynamicAppCaches {
    Write-LogAndHost "Scanning for application caches..." -NoHost -LogPrefix "Find-DynamicAppCaches"
    
    $discoveredCaches = [ordered]@{}
    $processedApps = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)

    # --- Start with a list of well-known application caches ---
    $wellKnownApps = [ordered]@{
        "NVIDIA Cache"             = @("$env:LOCALAPPDATA\NVIDIA\GLCache", "$env:ProgramData\NVIDIA Corporation\Downloader");
        "DirectX Shader Cache"     = @("$env:LOCALAPPDATA\D3DSCache");
        "Steam Cache"              = @("$env:LOCALAPPDATA\Steam\appcache", "$env:LOCALAPPDATA\Steam\htmlcache");
        "Discord Cache"            = @("$env:APPDATA\discord\Cache", "$env:APPDATA\discord\Code Cache", "$env:APPDATA\discord\GPUCache");
        "EA App Cache"             = @("$env:LOCALAPPDATA\Electronic Arts\EA Desktop\cache");
        "Spotify Cache"            = @("$env:LOCALAPPDATA\Spotify\Storage", "$env:LOCALAPPDATA\Spotify\Browser");
        "Visual Studio Code Cache" = @("$env:APPDATA\Code\Cache", "$env:APPDATA\Code\GPUCache", "$env:APPDATA\Code\CachedData");
        "Slack Cache"              = @("$env:APPDATA\Slack\Cache", "$env:APPDATA\Slack\GPUCache", "$env:APPDATA\Slack\Service Worker\CacheStorage");
        "Zoom Cache"               = @("$env:APPDATA\Zoom\data\Cache");
        "Adobe Cache"              = @("$env:LOCALAPPDATA\Adobe\Common\Media Cache Files", "$env:LOCALAPPDATA\Adobe\Common\Media Cache");
        "Telegram Cache"           = @("$env:APPDATA\Telegram Desktop\tdata\user_data\cache");
    }

    Write-LogAndHost "Checking for well-known application caches..." -NoHost
    foreach ($appName in $wellKnownApps.Keys) {
        $existingPaths = $wellKnownApps[$appName] | ForEach-Object { Resolve-Path $_ -ErrorAction SilentlyContinue } | Select-Object -ExpandProperty Path
        if ($existingPaths) {
            Write-LogAndHost "Found existing well-known cache for '$appName'." -NoHost
            $discoveredCaches[$appName] = @{ Paths = $existingPaths; Type = 'Folder' }
            [void]$processedApps.Add($appName)
        }
    }

    # --- Dynamically find installed web browser caches ---
    Write-LogAndHost "Scanning for installed web browser caches..." -NoHost
    $browserCacheConfigs = [ordered]@{
        "Google Chrome" = @{
            InstallPaths = @(
                "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
                "$env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe"
            );
            CacheDirs = @(
                "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache",
                "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Code Cache",
                "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\GPUCache"
            )
        };
        "Microsoft Edge" = @{
            InstallPaths = @(
                "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe"
            );
            CacheDirs = @(
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache",
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache",
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\GPUCache"
            )
        };
        "Brave Browser" = @{
            InstallPaths = @(
                "$env:ProgramFiles\BraveSoftware\Brave-Browser\Application\brave.exe",
                "$env:ProgramFiles(x86)\BraveSoftware\Brave-Browser\Application\brave.exe"
            );
            CacheDirs = @(
                "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data\Default\Cache",
                "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data\Default\Code Cache",
                "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data\Default\GPUCache"
            )
        };
        "Opera Browser" = @{
            InstallPaths = @(
                "$env:ProgramFiles\Opera\launcher.exe",
                "$env:ProgramFiles(x86)\Opera\launcher.exe"
            );
            CacheDirs = @(
                "$env:LOCALAPPDATA\Opera Software\Opera Stable\Cache",
                "$env:LOCALAPPDATA\Opera Software\Opera Stable\Code Cache",
                "$env:LOCALAPPDATA\Opera Software\Opera Stable\GPUCache"
            )
        };
        "Vivaldi Browser" = @{
            InstallPaths = @(
                "$env:ProgramFiles\Vivaldi\Application\vivaldi.exe",
                "$env:ProgramFiles(x86)\Vivaldi\Application\vivaldi.exe"
            );
            CacheDirs = @(
                "$env:LOCALAPPDATA\Vivaldi\User Data\Default\Cache",
                "$env:LOCALAPPDATA\Vivaldi\User Data\Default\Code Cache",
                "$env:LOCALAPPDATA\Vivaldi\User Data\Default\GPUCache"
            )
        };
        "Mozilla Firefox" = @{
            InstallPaths = @(
                "$env:ProgramFiles\Mozilla Firefox\firefox.exe",
                "$env:ProgramFiles(x86)\Mozilla Firefox\firefox.exe"
            );
            CacheDirs = @() # Populated dynamically below for profiles
        }
    }

    foreach ($browserName in $browserCacheConfigs.Keys) {
        $config = $browserCacheConfigs[$browserName]
        $isBrowserInstalled = $false
        foreach ($installPath in $config.InstallPaths) {
            if (Test-Path $installPath) {
                $isBrowserInstalled = $true
                break
            }
        }

        if ($isBrowserInstalled) {
            $foundBrowserPaths = New-Object System.Collections.Generic.List[string]
            if ($browserName -eq "Mozilla Firefox") {
                # Special handling for Firefox profiles
                $ffProfileDirs = Get-ChildItem -Path "$env:APPDATA\Mozilla\Firefox\Profiles" -Directory -ErrorAction SilentlyContinue
                if ($ffProfileDirs) {
                    $ffProfileDirs | ForEach-Object {
                        $cachePath = Join-Path $_.FullName "cache2"
                        if (Test-Path $cachePath) { [void]$foundBrowserPaths.Add($cachePath) }
                    }
                }
            } else {
                # For Chromium-based browsers
                foreach ($cacheDir in $config.CacheDirs) {
                    $resolvedPath = Resolve-Path $cacheDir -ErrorAction SilentlyContinue
                    if ($resolvedPath) { [void]$foundBrowserPaths.Add($resolvedPath.Path) }
                }
            }

            if ($foundBrowserPaths.Count -gt 0) {
                $uniquePaths = $foundBrowserPaths | Select-Object -Unique
                Write-LogAndHost "Found cache for installed browser '$browserName': $($uniquePaths -join ', ')" -NoHost
                $discoveredCaches[$browserName + " Cache"] = @{ Paths = $uniquePaths; Type = 'Folder' }
                [void]$processedApps.Add($browserName)
            }
        }
    }


    # --- Now, scan the registry for other installed applications ---
    Write-LogAndHost "Scanning registry for other installed applications..." -NoHost
    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    $cacheFolderNames = @("Cache", "cache", "Code Cache", "GPUCache", "ShaderCache", "temp", "tmp")
    $searchRoots = @($env:LOCALAPPDATA, $env:APPDATA, $env:ProgramData)

    foreach ($path in $uninstallPaths) {
        $regKeys = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
        if (-not $regKeys) { continue }

        foreach ($key in $regKeys) {
            $appName = $key.GetValue("DisplayName")
            $publisher = $key.GetValue("Publisher")

            if ([string]::IsNullOrWhiteSpace($appName) -or $appName -match "^KB[0-9]{6,}$" -or $processedApps.Contains($appName)) { continue }
            [void]$processedApps.Add($appName)

            $potentialNames = New-Object System.Collections.Generic.List[string]
            $potentialNames.Add($appName)
            if (-not [string]::IsNullOrWhiteSpace($publisher)) { $potentialNames.Add($publisher) }
            
            $foundPaths = New-Object System.Collections.Generic.List[string]

            foreach ($name in ($potentialNames | Select-Object -Unique)) {
                foreach ($root in $searchRoots) {
                    $basePath = Join-Path -Path $root -ChildPath $name
                    if (Test-Path $basePath) {
                        foreach ($cacheName in $cacheFolderNames) {
                            $cachePath = Join-Path -Path $basePath -ChildPath $cacheName
                            if (Test-Path $cachePath) { $foundPaths.Add($cachePath) }
                        }
                    }
                }
            }
            
            if ($foundPaths.Count -gt 0) {
                $uniquePaths = $foundPaths | Select-Object -Unique
                Write-LogAndHost "Found potential cache for '$appName': $($uniquePaths -join ', ')" -NoHost
                $discoveredCaches[$appName] = @{ Paths = $uniquePaths; Type = 'Folder' }
            }
        }
    }
    
    Write-LogAndHost "Finished cache scan. Found $($discoveredCaches.Count) applications with potential caches." -NoHost
    return $discoveredCaches
}

# ENHANCED FUNCTION: Clean temporary system files with a GUI, with consolidated dynamic cache discovery and path tooltips.
function Invoke-TempFileCleanup {
    if (-not $script:guiAvailable) {
        Write-LogAndHost "GUI is not available, cannot launch the System Cleanup tool." -HostColor Red -LogPrefix "Invoke-TempFileCleanup"
        Start-Sleep -Seconds 2
        return
    }
    Write-LogAndHost "Launching System Cleanup GUI..." -HostColor Cyan

    # --- DATA STRUCTURE for all cleanup items ---
    # Added DNS Cache and Microsoft Store Cache for more thorough cleaning.
    $cleanupItems = [ordered]@{
        "System Items" = [ordered]@{
            "Windows Temporary Files" = @{ Paths = @("$env:TEMP", "$env:windir\Temp"); Type = 'Folder' }
            "Windows Update Cache"    = @{ Paths = @("$env:windir\SoftwareDistribution\Download"); Type = 'Folder' }
            "Delivery Optimization"   = @{ Paths = @("$env:windir\SoftwareDistribution\DeliveryOptimization"); Type = 'Folder' }
            "Windows Log Files"       = @{ Paths = @("$env:windir\Logs"); Type = 'Folder' }
            "System Minidump Files"   = @{ Paths = @("$env:windir\Minidump"); Type = 'Folder' }
            "Windows Prefetch Files"  = @{ Paths = @("$env:windir\Prefetch"); Type = 'Folder' }
            "Thumbnail Cache"         = @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"; Filter = "thumbcache_*.db"; Type = 'File' }
            "Windows Icon Cache"      = @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"; Filter = "iconcache_*.db"; Type = 'File' }
            "Windows Font Cache"      = @{ Paths = @("$env:windir\ServiceProfiles\LocalService\AppData\Local\FontCache"); Type = 'Folder' }
            "Microsoft Store Cache"   = @{ Paths = @("$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache"); Type = 'Folder' }
            "DNS Cache"               = @{ Type = 'Special' }
            "Recycle Bin"             = @{ Type = 'Special' }
        }
        "Discovered Application Caches" = [ordered]@{} # Category for all discovered app caches, including browsers.
    }

    # --- GUI Setup ---
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Perdanga System Cleanup"; $form.Size = New-Object System.Drawing.Size(600, 720) # Increased height for new controls
    $form.StartPosition = "CenterScreen"; $form.FormBorderStyle = "FixedDialog"; $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)

    $commonFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $controlBackColor = [System.Drawing.Color]::FromArgb(60, 60, 63)
    $controlForeColor = [System.Drawing.Color]::White

    # --- ToolTip Setup for displaying paths ---
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.AutoPopDelay = 20000; $toolTip.InitialDelay = 700; $toolTip.ReshowDelay = 500
    $toolTip.UseFading = $true; $toolTip.UseAnimation = $true

    # --- TreeView Setup ---
    $treeView = New-Object System.Windows.Forms.TreeView
    $treeView.Location = New-Object System.Drawing.Point(15, 55); $treeView.Size = New-Object System.Drawing.Size(560, 360) # Repositioned
    $treeView.CheckBoxes = $true; $treeView.Font = $commonFont; $treeView.BackColor = $controlBackColor
    $treeView.ForeColor = $controlForeColor; $treeView.BorderStyle = "FixedSingle"; $treeView.FullRowSelect = $true
    $treeView.ShowNodeToolTips = $false
    $form.Controls.Add($treeView) | Out-Null

    # --- Populate TreeView with static items ---
    foreach ($categoryName in $cleanupItems.Keys) {
        $parentNode = $treeView.Nodes.Add($categoryName, $categoryName)
        if ($categoryName -ne "Discovered Application Caches") {
            foreach ($itemName in $cleanupItems[$categoryName].Keys) {
                $itemConfig = $cleanupItems[$categoryName][$itemName]
                if ($itemConfig.Type -eq 'Folder' -and $itemConfig.Paths.Count -eq 0) { continue }
                $childNode = $parentNode.Nodes.Add($itemName, $itemName); $childNode.Tag = $itemConfig
                $childNode.Checked = $false 
            }
            $parentNode.Checked = $false; $parentNode.Expand()
        }
    }

    # --- Log Box Setup ---
    $logBox = New-Object System.Windows.Forms.RichTextBox
    $logBox.Location = New-Object System.Drawing.Point(15, 455); $logBox.Size = New-Object System.Drawing.Size(560, 150) # Repositioned
    $logBox.Font = New-Object System.Drawing.Font("Consolas", 9); $logBox.BackColor = $controlBackColor; $logBox.ForeColor = $controlForeColor
    $logBox.ReadOnly = $true; $logBox.BorderStyle = "FixedSingle"; $logBox.ScrollBars = "Vertical"
    $form.Controls.Add($logBox) | Out-Null
    
    # --- Progress Bar Setup ---
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(15, 425); $progressBar.Size = New-Object System.Drawing.Size(560, 20) # Repositioned
    $progressBar.Visible = $false
    $form.Controls.Add($progressBar) | Out-Null

    # --- Button Setup ---
    $buttonAnalyze = New-Object System.Windows.Forms.Button; $buttonAnalyze.Text = "Analyze"; $buttonAnalyze.Size = "120,30"; $buttonAnalyze.Location = "100,625"; # Repositioned
    $buttonClean = New-Object System.Windows.Forms.Button; $buttonClean.Text = "Clean"; $buttonClean.Size = "120,30"; $buttonClean.Location = "235,625"; # Repositioned
    $buttonClose = New-Object System.Windows.Forms.Button; $buttonClose.Text = "Exit"; $buttonClose.Size = "120,30"; $buttonClose.Location = "370,625"; # Repositioned
    
    # --- Select/Deselect All Buttons ---
    $buttonSelectAll = New-Object System.Windows.Forms.Button; $buttonSelectAll.Text = "Select All"; $buttonSelectAll.Size = "120,30"; $buttonSelectAll.Location = "15,15";
    $buttonDeselectAll = New-Object System.Windows.Forms.Button; $buttonDeselectAll.Text = "Deselect All"; $buttonDeselectAll.Size = "120,30"; $buttonDeselectAll.Location = "145,15";
    
    @( $buttonAnalyze, $buttonClean, $buttonClose, $buttonSelectAll, $buttonDeselectAll ) | ForEach-Object {
        $_.Font = $commonFont; $_.ForeColor = [System.Drawing.Color]::White; $_.FlatStyle = "Flat"; $_.FlatAppearance.BorderSize = 0; $form.Controls.Add($_) | Out-Null
    }
    $buttonAnalyze.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180); $buttonClean.BackColor = [System.Drawing.Color]::FromArgb(200, 70, 70); $buttonClose.BackColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $buttonSelectAll.BackColor = [System.Drawing.Color]::FromArgb(80, 150, 80); $buttonDeselectAll.BackColor = [System.Drawing.Color]::FromArgb(150, 80, 80)
    
    # --- Event Handlers & Logic ---
    $analyzedData = @{}
    $logWriter = { param($Message, $Color = 'White'); $logBox.SelectionStart = $logBox.TextLength; $logBox.SelectionLength = 0; $logBox.SelectionColor = $Color; $logBox.AppendText("$(Get-Date -Format 'HH:mm:ss') - $Message`n"); $logBox.ScrollToCaret() }
    $form.Tag = @{ logWriter = $logWriter }

    # --- DYNAMICALLY FIND AND POPULATE APP CACHES ---
    $form.Add_Shown({
        $logWriterFunc = $form.Tag.logWriter
        & $logWriterFunc "Scanning for application caches (this may take a moment)..." 'LightBlue'; $form.Update()

        $dynamicCaches = Find-DynamicAppCaches
        $discoveredNode = $treeView.Nodes["Discovered Application Caches"]

        if ($dynamicCaches.Count -gt 0) {
            $treeView.BeginUpdate()
            foreach ($appName in ($dynamicCaches.Keys | Sort-Object)) {
                $itemConfig = $dynamicCaches[$appName]
                $cleanupItems["Discovered Application Caches"][$appName] = $itemConfig
                $childNode = $discoveredNode.Nodes.Add($appName, $appName)
                $childNode.Tag = $itemConfig
                $childNode.Checked = $false 
            }
            $discoveredNode.Expand(); $treeView.EndUpdate()
            $discoveredNode.Checked = $false 
            & $logWriterFunc "Scan complete. Found $($dynamicCaches.Count) items. Review and select them for cleaning." 'Green'
        } else {
            & $logWriterFunc "No additional application caches were found." 'Gray'
            $treeView.Nodes.Remove($discoveredNode)
        }
        # FIX: Ensure the view is scrolled to the top after loading.
        if ($treeView.Nodes.Count -gt 0) {
            $treeView.SelectedNode = $treeView.Nodes[0]
            $treeView.Nodes[0].EnsureVisible()
        }
    })

    # --- Tooltip Handler ---
    $lastHoveredNode = $null
    $treeView.add_MouseMove({
        param($sender, $e)
        $node = $treeView.GetNodeAt($e.X, $e.Y)
        if ($node -and $node -ne $lastHoveredNode) {
            $lastHoveredNode = $node
            $itemConfig = $node.Tag
            if ($itemConfig -and $itemConfig.ContainsKey("Paths")) {
                $tooltipText = "Paths to be cleaned:`n" + ($itemConfig.Paths -join "`n")
                $toolTip.SetToolTip($treeView, $tooltipText)
            } else {
                $toolTip.SetToolTip($treeView, "")
            }
        } elseif (-not $node) {
            $lastHoveredNode = $null
            $toolTip.SetToolTip($treeView, "")
        }
    })

    # --- Checkbox logic ---
    $updatingChecks = $false
    $treeView.add_AfterCheck({
        param($sender, $e)
        if ($updatingChecks) { return }
        $updatingChecks = $true
        try {
            if ($e.Node.Parent -eq $null) { # Parent node was clicked
                foreach ($childNode in $e.Node.Nodes) { $childNode.Checked = $e.Node.Checked }
            } else { # Child node was clicked
                $parent = $e.Node.Parent
                $allChecked = $true; $noneChecked = $true
                foreach ($sibling in $parent.Nodes) { if ($sibling.Checked) { $noneChecked = $false } else { $allChecked = $false } }
                if ($allChecked) { $parent.Checked = $true } elseif ($noneChecked) { $parent.Checked = $false } else { $parent.Checked = $false }
            }
        } finally {
            $updatingChecks = $false
        }
    })

    # --- Button Click Handlers ---
    $buttonSelectAll.add_Click({ foreach ($node in $treeView.Nodes) { $node.Checked = $true } }) | Out-Null
    $buttonDeselectAll.add_Click({ foreach ($node in $treeView.Nodes) { $node.Checked = $false } }) | Out-Null

    $buttonAnalyze.add_Click({
        $logBox.Clear(); & $logWriter "Starting analysis..." 'Cyan'; $analyzedData.Clear()
        $progressBar.Value = 0; $progressBar.Visible = $true; $form.Update()
        
        $nodesToAnalyze = @()
        foreach ($parentNode in $treeView.Nodes) {
            foreach ($childNode in $parentNode.Nodes) {
                if ($childNode.Checked) { $nodesToAnalyze += $childNode }
            }
        }
        if ($nodesToAnalyze.Count -eq 0) { & $logWriter "No items selected for analysis." 'Yellow'; $progressBar.Visible = $false; return }

        $progressBar.Maximum = $nodesToAnalyze.Count
        $treeView.BeginUpdate()
        $totalSize = 0
        
        foreach ($childNode in $nodesToAnalyze) {
            $itemConfig = $childNode.Tag; $itemSize = 0
            switch ($itemConfig.Type) {
                'Folder' { foreach ($p in $itemConfig.Paths) { if(Test-Path $p){ $itemSize += (Get-ChildItem $p -Recurse -Force -EA SilentlyContinue | Measure-Object -Property Length -Sum -EA SilentlyContinue).Sum } } }
                'File' { if(Test-Path $itemConfig.Path){ $itemSize = (Get-ChildItem $itemConfig.Path -Filter $itemConfig.Filter -Force -EA SilentlyContinue | Measure-Object -Property Length -Sum -EA SilentlyContinue).Sum } }
                'Special' { 
                    if ($childNode.Name -eq "Recycle Bin") {
                        $shell = New-Object -ComObject Shell.Application; $recycleBin = $shell.NameSpace(0xA); $itemSize = ($recycleBin.Items() | ForEach-Object { $_.Size } | Measure-Object -Sum).Sum 
                    }
                    # Special items like DNS cache have no size to measure.
                }
            }
            $analyzedData[$childNode.Name] = $itemSize; $totalSize += $itemSize
            $baseNodeName = $childNode.Name -replace " \(\d+(\.\d+)? (MB|KB)\)$"
            $nodeSizeText = if ($itemSize -gt 1MB) { " ($([math]::Round($itemSize/1MB, 2)) MB)" } elseif ($itemSize -gt 0) { " ($([math]::Round($itemSize/1KB, 2)) KB)" } else { "" }
            $childNode.Text = $baseNodeName + $nodeSizeText
            $progressBar.PerformStep()
        }
        $treeView.EndUpdate()
        $sizeInMB = [math]::Round($totalSize / 1MB, 2)
        & $logWriter "Analysis complete. Found $sizeInMB MB of files to clean." 'Green'
        $progressBar.Visible = $false
    })

    $buttonClean.add_Click({
        if ($analyzedData.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please run an analysis first.", "Analysis Required", "OK", "Information") | Out-Null; return }
        $confirmResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to permanently delete these files?", "Confirm Deletion", "YesNo", "Warning")
        if ($confirmResult -ne 'Yes') { & $logWriter "Cleanup cancelled by user." 'Yellow'; return }
        
        & $logWriter "Starting cleanup..." 'Cyan'
        $progressBar.Value = 0; $progressBar.Visible = $true
        @( $buttonAnalyze, $buttonClean, $buttonSelectAll, $buttonDeselectAll ) | ForEach-Object { $_.Enabled = $false }; $form.Update()
        
        $nodesToClean = @()
        foreach ($parentNode in $treeView.Nodes) {
            foreach ($childNode in $parentNode.Nodes) {
                if ($childNode.Checked) { $nodesToClean += $childNode }
            }
        }
        $progressBar.Maximum = $nodesToClean.Count
        $totalDeleted = 0

        foreach ($childNode in $nodesToClean) {
            $baseNodeName = $childNode.Name -replace " \(\d+(\.\d+)? (MB|KB)\)$"
            & $logWriter "Cleaning $($baseNodeName)..." 'White'
            $itemConfig = $childNode.Tag
            try {
                switch ($itemConfig.Type) {
                    'Folder' { foreach ($p in $itemConfig.Paths) { if (Test-Path $p) { Remove-Item -Path "$p\*" -Recurse -Force -EA SilentlyContinue } } }
                    'File' { if (Test-Path $itemConfig.Path) { Remove-Item -Path (Join-Path $itemConfig.Path $itemConfig.Filter) -Force -EA SilentlyContinue } }
                    'Special' { 
                        if ($baseNodeName -eq "Recycle Bin") { Clear-RecycleBin -Force -ErrorAction Stop }
                        if ($baseNodeName -eq "DNS Cache") { Start-Process -FilePath "ipconfig" -ArgumentList "/flushdns" -WindowStyle Hidden -Wait }
                    }
                }
                if ($analyzedData.ContainsKey($childNode.Name)) {
                    $totalDeleted += $analyzedData[$childNode.Name]
                }
            } catch { & $logWriter "Failed to clean $($baseNodeName). Error: $($_.Exception.Message)" 'Red' }
            $progressBar.PerformStep()
        }
        
        $deletedInMB = [math]::Round($totalDeleted / 1MB, 2)
        & $logWriter "Cleanup complete. Freed approximately $deletedInMB MB of space." 'Green'
        & $logWriter "Perdanga Forever!" 'Magenta'
        
        $analyzedData.Clear(); $progressBar.Visible = $false
        @( $buttonAnalyze, $buttonClean, $buttonSelectAll, $buttonDeselectAll ) | ForEach-Object { $_.Enabled = $true }
        
        foreach ($node in $treeView.Nodes) { 
            $node.Text = $node.Name -replace " \(\d+(\.\d+)? (MB|KB)\)$"
            foreach($child in $node.Nodes) { $child.Text = $child.Name -replace " \(\d+(\.\d+)? (MB|KB)\)$" } 
        }
    })

    $buttonClose.add_Click({ $form.Close() }) | Out-Null

    try { $null = $form.ShowDialog() } catch { Write-LogAndHost "An unexpected error occurred with the System Cleanup GUI. Details: $($_.Exception.Message)" -HostColor Red } finally { $form.Dispose() }
    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host
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
    $toolTip.AutoPopDelay = 10000; $toolTip.InitialDelay = 500; $toolTip.ReshowDelay = 500

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
        
        for($i = 0; $i -lt $listKeyboardLayouts.Items.Count; $i++) {
            if ($checkedKeyboardLayoutNames.Contains($listKeyboardLayouts.Items[$i])) {
                $listKeyboardLayouts.SetItemChecked($i, $true)
            }
        }
        $listKeyboardLayouts.EndUpdate()
    }) | Out-Null
    
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
        
    $xml = New-Object System.Xml.XmlDocument
    $xml.AppendChild($xml.CreateXmlDeclaration("1.0", "utf-8", $null)) | Out-Null
    $root = $xml.CreateElement("unattend"); $root.SetAttribute("xmlns", "urn:schemas-microsoft-com:unattend"); $xml.AppendChild($root) | Out-Null
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable); $ns.AddNamespace("d6p1", "http://schemas.microsoft.com/WMIConfig/2002/State")
    
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
    
    $firstLogonCommands = $compShellOobe.AppendChild($xml.CreateElement("FirstLogonCommands"))
    $commandIndex = 1
    
    if ($formData.ShowFileExt) { $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand")); $syncCmd.SetAttribute("Order", $commandIndex++); $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = 'cmd /c reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v HideFileExt /t REG_DWORD /d 0 /f' }
    if ($formData.DisableSmartScreen) { $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand")); $syncCmd.SetAttribute("Order", $commandIndex++); $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = 'cmd /c reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" /v SmartScreenEnabled /t REG_SZ /d "Off" /f' }
    if ($formData.DisableSysRestore) { $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand")); $syncCmd.SetAttribute("Order", $commandIndex++); $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = 'powershell.exe -Command "Disable-ComputerRestore -Drive C:\"' }
    if ($formData.DisableSuggestions) { $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand")); $syncCmd.SetAttribute("Order", $commandIndex++); $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = 'cmd /c reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v SubscribedContent-338389Enabled /t REG_DWORD /d 0 /f' }

    $bloatwareCommands = @{
        '3D Viewer' = 'Get-AppxPackage *Microsoft.Microsoft3DViewer* | Remove-AppxPackage -AllUsers'; 'Bing Search' = 'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Search" /v BingSearchEnabled /t REG_DWORD /d 0 /f'; 'Calculator' = 'Get-AppxPackage *Microsoft.WindowsCalculator* | Remove-AppxPackage -AllUsers'; 'Camera' = 'Get-AppxPackage *Microsoft.WindowsCamera* | Remove-AppxPackage -AllUsers'; 'Clipchamp' = 'Get-AppxPackage *Microsoft.Clipchamp* | Remove-AppxPackage -AllUsers'; 'Clock' = 'Get-AppxPackage *Microsoft.WindowsAlarms* | Remove-AppxPackage -AllUsers'; 'Copilot' = 'reg add "HKCU\Software\Policies\Microsoft\Windows\WindowsCopilot" /v TurnOffWindowsCopilot /t REG_DWORD /d 1 /f'; 'Cortana' = 'Get-AppxPackage *Microsoft.549981C3F5F10* | Remove-AppxPackage -AllUsers'; 'Dev Home' = 'Get-AppxPackage *Microsoft.DevHome* | Remove-AppxPackage -AllUsers'; 'Family' = 'Get-AppxPackage *Microsoft.Windows.Family* | Remove-AppxPackage -AllUsers'; 'Feedback Hub' = 'Get-AppxPackage *Microsoft.WindowsFeedbackHub* | Remove-AppxPackage -AllUsers'; 'Get Help' = 'Get-AppxPackage *Microsoft.GetHelp* | Remove-AppxPackage -AllUsers'; 'Handwriting (all languages)' = 'Get-WindowsCapability -Online | Where-Object { $_.Name -like "Language.Handwriting*" } | ForEach-Object { Remove-WindowsCapability -Online -Name $_.Name -NoRestart }'; 'Internet Explorer' = 'Disable-WindowsOptionalFeature -Online -FeatureName "Internet-Explorer-Optional-amd64" -NoRestart'; 'Mail and Calendar' = 'Get-AppxPackage *microsoft.windowscommunicationsapps* | Remove-AppxPackage -AllUsers'; 'Maps' = 'Get-AppxPackage *Microsoft.WindowsMaps* | Remove-AppxPackage -AllUsers'; 'Math Input Panel' = 'Remove-WindowsCapability -Online -Name "MathRecognizer~~~~0.0.1.0" -NoRestart'; 'Media Features' = 'Disable-WindowsOptionalFeature -Online -FeatureName "MediaPlayback" -NoRestart'; 'Mixed Reality' = 'Get-AppxPackage *Microsoft.MixedReality.Portal* | Remove-AppxPackage -AllUsers'; 'Movies & TV' = 'Get-AppxPackage *Microsoft.ZuneVideo* | Remove-AppxPackage -AllUsers'; 'News' = 'Get-AppxPackage *Microsoft.BingNews* | Remove-AppxPackage -AllUsers'; 'Notepad (modern)' = 'Get-AppxPackage *Microsoft.WindowsNotepad* | Remove-AppxPackage -AllUsers'; 'Office 365' = 'Get-AppxPackage *Microsoft.MicrosoftOfficeHub* | Remove-AppxPackage -AllUsers'; 'OneDrive' = '$process = Start-Process "$env:SystemRoot\SysWOW64\OneDriveSetup.exe" -ArgumentList "/uninstall" -PassThru -Wait; if ($process.ExitCode -ne 0) { Start-Process "$env:SystemRoot\System32\OneDriveSetup.exe" -ArgumentList "/uninstall" -PassThru -Wait }'; 'OneNote' = 'Get-AppxPackage *Microsoft.Office.OneNote* | Remove-AppxPackage -AllUsers'; 'OneSync' = '# Handled by Mail and Calendar'; 'OpenSSH Client' = 'Remove-WindowsCapability -Online -Name "OpenSSH.Client~~~~0.0.1.0" -NoRestart'; 'Outlook for Windows' = 'Get-AppxPackage *Microsoft.OutlookForWindows* | Remove-AppxPackage -AllUsers'; 'Paint' = 'Get-AppxPackage *Microsoft.Paint* | Remove-AppxPackage -AllUsers'; 'Paint 3D' = 'Get-AppxPackage *Microsoft.MSPaint* | Remove-AppxPackage -AllUsers'; 'People' = 'Get-AppxPackage *Microsoft.People* | Remove-AppxPackage -AllUsers'; 'Photos' = 'Get-AppxPackage *Microsoft.Windows.Photos* | Remove-AppxPackage -AllUsers'; 'Power Automate' = 'Get-AppxPackage *Microsoft.PowerAutomateDesktop* | Remove-AppxPackage -AllUsers'; 'PowerShell 2.0' = 'Disable-WindowsOptionalFeature -Online -FeatureName "MicrosoftWindowsPowerShellV2" -NoRestart'; 'PowerShell ISE' = 'Remove-WindowsCapability -Online -Name "PowerShell-ISE-v2~~~~0.0.1.0" -NoRestart'; 'Quick Assist' = 'Get-AppxPackage *Microsoft.QuickAssist* | Remove-AppxPackage -AllUsers'; 'Recall' = 'reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" /v DisableAllScreenshotCapture /t REG_DWORD /d 1 /f'; 'Remote Desktop Client' = '# Core component, removal not recommended.'; 'Skype' = 'Get-AppxPackage *Microsoft.SkypeApp* | Remove-AppxPackage -AllUsers'; 'Snipping Tool' = 'Get-AppxPackage *Microsoft.ScreenSketch* | Remove-AppxPackage -AllUsers'; 'Solitaire Collection' = 'Get-AppxPackage *Microsoft.MicrosoftSolitaireCollection* | Remove-AppxPackage -AllUsers'; 'Speech (all languages)' = 'Get-WindowsCapability -Online | Where-Object { $_.Name -like "Language.Speech*" } | ForEach-Object { Remove-WindowsCapability -Online -Name $_.Name -NoRestart }'; 'Steps Recorder' = 'Disable-WindowsOptionalFeature -Online -FeatureName "StepsRecorder" -NoRestart'; 'Sticky Notes' = 'Get-AppxPackage *Microsoft.MicrosoftStickyNotes* | Remove-AppxPackage -AllUsers'; 'Teams' = 'Get-AppxPackage *MicrosoftTeams* | Remove-AppxPackage -AllUsers'; 'Tips' = 'Get-AppxPackage *Microsoft.Getstarted* | Remove-AppxPackage -AllUsers'; 'To Do' = 'Get-AppxPackage *Microsoft.Todos* | Remove-AppxPackage -AllUsers'; 'Voice Recorder' = 'Get-AppxPackage *Microsoft.WindowsSoundRecorder* | Remove-AppxPackage -AllUsers'; 'Wallet' = 'Get-AppxPackage *Microsoft.Wallet* | Remove-AppxPackage -AllUsers'; 'Weather' = 'Get-AppxPackage *Microsoft.BingWeather* | Remove-AppxPackage -AllUsers'; 'Windows Fax and Scan' = 'Disable-WindowsOptionalFeature -Online -FeatureName "Windows-Fax-And-Scan" -NoRestart'; 'Windows Hello' = 'reg add "HKLM\SOFTWARE\Policies\Microsoft\Biometrics" /v Enabled /t REG_DWORD /d 0 /f; reg add "HKLM\SOFTWARE\Policies\Microsoft\Biometrics\CredentialProviders" /v Enabled /t REG_DWORD /d 0 /f'; 'Windows Media Player (classic)' = 'Disable-WindowsOptionalFeature -Online -FeatureName "WindowsMediaPlayer" -NoRestart'; 'Windows Media Player (modern)' = 'Get-AppxPackage *Microsoft.ZuneMusic* | Remove-AppxPackage -AllUsers'; 'Windows Terminal' = 'Get-AppxPackage *Microsoft.WindowsTerminal* | Remove-AppxPackage -AllUsers'; 'WordPad' = 'Remove-WindowsCapability -Online -Name "WordPad~~~~0.0.1.0" -NoRestart'; 'Xbox Apps' = 'Get-AppxPackage *Microsoft.Xbox* | Remove-AppxPackage -AllUsers; Get-AppxPackage *Microsoft.GamingApp* | Remove-AppxPackage -AllUsers'; 'Your Phone / Phone Link' = 'Get-AppxPackage *Microsoft.YourPhone* | Remove-AppxPackage -AllUsers'
    }

    foreach ($bloat in $formData.BloatwareToRemove) {
        if ($bloatwareCommands.ContainsKey($bloat)) {
            $command = $bloatwareCommands[$bloat]; if ($command.StartsWith("#")) { continue }
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
        $searchOutput = & choco search $PackageName --exact --limit-output --source="https://community.chocolatey.org/api/v2/" --no-progress 2>&1
        $searchOutput | Out-File -FilePath $script:logFile -Append -Encoding UTF8

        if ($LASTEXITCODE -ne 0) {
             Write-LogAndHost "Error during 'choco search' for '$PackageName'. Exit code: $LASTEXITCODE. Output: $($searchOutput | Out-String)" -HostColor Red -LogPrefix "Test-ChocolateyPackage"
             return $false
        }
        
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
    $menuLines.Add(" Windows & Software Manager [PSS v1.6] ($(Get-Date -Format "dd.MM.yyyy HH:mm"))") 
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
        $consoleWidth = [System.Console]::WindowWidth
    } catch {}

    if ($consoleWidth -le 1) { 
        try {
            $consoleWidth = $Host.UI.RawUI.WindowSize.Width
        } catch {}
    }
    if ($consoleWidth -le 1) {
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

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

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
        if (-not (Install-Chocolatey)) {
            Write-LogAndHost "Chocolatey is not installed. Proceeding without Chocolatey." -HostColor Yellow -LogPrefix "Choco-Init"
        } else {
            if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
                Write-LogAndHost "Chocolatey command not found after installation attempt. Proceeding without Chocolatey." -HostColor Yellow -LogPrefix "Choco-Init"
            } else {
                if (-not $env:ChocolateyInstall) {
                    $env:ChocolateyInstall = "$($env:ProgramData)\chocolatey"
                    Write-LogAndHost "ChocolateyInstall environment variable set to: $env:ChocolateyInstall" -NoHost
                }
                $chocoVersion = & choco --version 2>&1
                if ($LASTEXITCODE -ne 0) {
                    Write-LogAndHost "Chocolatey is not functioning correctly. Exit code: $LASTEXITCODE. Proceeding without Chocolatey." -HostColor Yellow -LogPrefix "Choco-Init"
                } else {
                    Write-LogAndHost "Found Chocolatey version: $($chocoVersion -join ' ')"
                }
            }
        }
    } else {
        if (-not $env:ChocolateyInstall) {
            $env:ChocolateyInstall = "$($env:ProgramData)\chocolatey"
            Write-LogAndHost "ChocolateyInstall environment variable set to: $env:ChocolateyInstall" -NoHost
        }
        $chocoVersion = & choco --version 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-LogAndHost "Chocolatey is not functioning correctly. Exit code: $LASTEXITCODE. Proceeding without Chocolatey." -HostColor Yellow -LogPrefix "Choco-Init"
        } else {
            Write-LogAndHost "Found Chocolatey version: $($chocoVersion -join ' ')"
        }
    }
}
catch {
    Write-LogAndHost "Exception occurred while checking Chocolatey. $($_.Exception.Message). Proceeding without Chocolatey." -HostColor Yellow -LogPrefix "Choco-Init"
}

Write-Host ""; Write-LogAndHost "Perdanga Forever!"
try {
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        # Enable global confirmation to prevent prompts during installations.
        & choco feature enable -n allowGlobalConfirmation 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8
        if ($LASTEXITCODE -eq 0) {
            Write-LogAndHost "Automatic confirmation enabled."
        } else {
            Write-LogAndHost "Failed to enable automatic confirmation. $($LASTEXITCODE)" -HostColor Yellow -LogPrefix "Choco-Init"
        }
    } else {
        Write-LogAndHost "Skipping Chocolatey global confirmation setup as Chocolatey is not installed." -HostColor Yellow -LogPrefix "Choco-Init"
    }
}
catch {
    Write-LogAndHost "Exception enabling automatic confirmation. $($_.Exception.Message). Proceeding without Chocolatey configuration." -HostColor Yellow -LogPrefix "Choco-Init"
}
Write-Host ""

$script:firstRun = $true

# --- MAIN LOOP ---
do {
    Show-Menu
    try { $userInput = Read-Host } catch { Write-LogAndHost "Could not read user input. $($_.Exception.Message)" -HostColor Red -LogPrefix "Main-Loop"; Start-Sleep -Seconds 2; continue }
    $userInput = $userInput.Trim().ToLower()

    if ($userInput -eq 'perdanga') {
        Show-PerdangaArt
        continue
    }

    if ([string]::IsNullOrEmpty($userInput)) {
        Clear-Host; Write-LogAndHost "No input detected. Please enter an option." -HostColor Yellow
        Write-LogAndHost "Press any key to return to the menu..." -HostColor DarkGray -NoLog; $null = Read-Host; continue
    }
    
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
    else {
        Clear-Host
        $validOptions = ($script:mainMenuLetters | Sort-Object | ForEach-Object { $_.ToUpper() }) -join ','
        $errorMessage = "Invalid input: '$userInput'. Use options [$validOptions], program numbers, or a secret word."
        Write-LogAndHost $errorMessage -HostColor Red -LogPrefix "Main-Loop"
        Start-Sleep -Seconds 2
    }
} while ($true)


