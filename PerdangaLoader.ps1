<#
.SYNOPSIS
    Author: Roman Zhdanov
    Version: 1.7
    Last Modified: 04.12.2025
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

function Invoke-DisableTelemetry {
    Write-LogAndHost "Checking Windows Telemetry status..." -HostColor Cyan

    # Define services and registry keys to manage
    $servicesToDisable = @("DiagTrack", "dmwappushservice", "WerSvc")
    $regKeys = @{
        "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" = @{
            "AllowTelemetry" = @{ Value = 0; Type = "DWord" }
        }
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection" = @{
            "AllowTelemetry" = @{ Value = 0; Type = "DWord" }
        }
        "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" = @{
            "DisableWindowsConsumerFeatures" = @{ Value = 1; Type = "DWord" }
        }
        "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo" = @{
            "DisabledByGroupPolicy" = @{ Value = 1; Type = "DWord" }
        }
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Privacy" = @{
            "TailoredExperiencesWithDiagnosticDataEnabled" = @{ Value = 0; Type = "DWord" }
            "AdvertisingID" = @{ Value = 0; Type = "DWord" }
        }
        "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting" = @{
            "DisableCollection" = @{ Value = 1; Type = "DWord" }
        }
    }

    # Check if telemetry is already disabled
    $servicesDisabled = $true
    foreach ($serviceName in $servicesToDisable) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($service) {
            if ($service.StartType -ne 'Disabled' -or $service.Status -ne 'Stopped') {
                $servicesDisabled = $false
                break
            }
        }
    }

    $regDisabled = $true
    foreach ($path in $regKeys.Keys) {
        $values = $regKeys[$path]
        foreach ($valueName in $values.Keys) {
            try {
                $currentValue = Get-ItemProperty -Path $path -Name $valueName -ErrorAction Stop
                if ($currentValue.$valueName -ne $values[$valueName].Value) {
                    $regDisabled = $false
                    break
                }
            } catch {
                $regDisabled = $false
                break
            }
        }
        if (-not $regDisabled) { break }
    }

    if ($servicesDisabled -and $regDisabled) {
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

    # Process services
    foreach ($serviceName in $servicesToDisable) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if (-not $service) {
            Write-LogAndHost "Service '$serviceName' not found, skipping." -HostColor DarkGray -NoLog
            continue
        }

        $changed = $false
        # Stop service if running
        if ($service.Status -ne 'Stopped') {
            try {
                Stop-Service -Name $serviceName -Force -ErrorAction Stop
                $changed = $true
                Write-LogAndHost "Stopped service: $serviceName" -NoLog
            } catch {
                Write-LogAndHost "Could not stop service '$serviceName'. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Invoke-DisableTelemetry"
            }
        }

        # Disable startup type
        if ($service.StartType -ne 'Disabled') {
            try {
                Set-Service -Name $serviceName -StartupType Disabled -ErrorAction Stop
                $changed = $true
                Write-LogAndHost "Disabled service: $serviceName" -NoLog
            } catch {
                Write-LogAndHost "Could not disable service '$serviceName'. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Invoke-DisableTelemetry"
            }
        }

        if (-not $changed) {
            Write-LogAndHost "Service '$serviceName' is already stopped and disabled." -HostColor DarkGray -NoLog
        }
    }

    # Process registry keys
    foreach ($path in $regKeys.Keys) {
        $values = $regKeys[$path]
        foreach ($valueName in $values.Keys) {
            $valueInfo = $values[$valueName]
            try {
                $currentValue = Get-ItemProperty -Path $path -Name $valueName -ErrorAction Stop
                if ($currentValue.$valueName -eq $valueInfo.Value) {
                    Write-LogAndHost "Registry value '$valueName' at '$path' is already set to $($valueInfo.Value). Skipping." -HostColor DarkGray -NoLog
                    continue
                }
            } catch {
                # Value doesn't exist or path doesn't exist - proceed to create/set
            }

            try {
                if (-not (Test-Path $path)) {
                    Write-LogAndHost "Creating registry path: $path" -NoLog
                    New-Item -Path $path -Force -ErrorAction Stop | Out-Null
                }
                Set-ItemProperty -Path $path -Name $valueName -Value $valueInfo.Value -Type $valueInfo.Type -Force -ErrorAction Stop
                Write-LogAndHost "Successfully set registry value '$valueName' at '$path'." -NoHost
            } catch {
                Write-LogAndHost "Failed to set registry key at '$path' for value '$valueName'. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Invoke-DisableTelemetry"
            }
        }
    }

    Write-LogAndHost "Telemetry has been successfully disabled." -HostColor Green
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

    Write-LogAndHost "Launching System Information GUI..." -HostColor Cyan -LogPrefix "Show-SystemInfo"
    
    try {
        # --- GUI Setup ---
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Perdanga System Information"
        $form.Size = New-Object System.Drawing.Size(950, 750)
        $form.StartPosition = "CenterScreen"
        $form.FormBorderStyle = "FixedDialog"
        $form.MaximizeBox = $false

        # Gemini Theme Colors
        $geminiDarkBg = [System.Drawing.Color]::FromArgb(20, 20, 25)
        $geminiPanelBg = [System.Drawing.Color]::FromArgb(35, 35, 40)
        $geminiBlue = [System.Drawing.Color]::FromArgb(60, 100, 180)
        $geminiYellow = [System.Drawing.Color]::FromArgb(230, 180, 50)
        $geminiAccent = [System.Drawing.Color]::FromArgb(0, 200, 255)
        $geminiGrayText = [System.Drawing.Color]::Gainsboro
        $geminiWhiteText = [System.Drawing.Color]::White

        $form.BackColor = $geminiDarkBg
        $form.Opacity = 0
        
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
        $mainTableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 60))) | Out-Null
        $mainTableLayout.BackColor = $geminiDarkBg
        $mainTableLayout.Padding = New-Object System.Windows.Forms.Padding(10)
        $form.Controls.Add($mainTableLayout)

        # --- Define improved fonts ---
        $headerFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $labelFont = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
        $valueFont = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
        $buttonFont = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

        # --- Helper function to create styled GroupBoxes ---
        function New-InfoGroupBox($Text, $HeaderColor) {
            $groupbox = New-Object System.Windows.Forms.GroupBox
            $groupbox.Text = $Text
            $groupbox.Font = $headerFont
            $groupbox.ForeColor = $HeaderColor
            $groupbox.Dock = "Fill"
            $groupbox.Padding = New-Object System.Windows.Forms.Padding(8, 20, 8, 8)
            $groupbox.Margin = New-Object System.Windows.Forms.Padding(6)
            $groupbox.BackColor = $geminiPanelBg

            # Create a nested panel with AutoScroll
            $scrollPanel = New-Object System.Windows.Forms.Panel
            $scrollPanel.Dock = "Fill"
            $scrollPanel.AutoScroll = $true
            $scrollPanel.BackColor = $geminiPanelBg
            $scrollPanel.Padding = New-Object System.Windows.Forms.Padding(4)
            $groupbox.Controls.Add($scrollPanel)

            # Create a nested TableLayoutPanel
            $tlp = New-Object System.Windows.Forms.TableLayoutPanel
            $tlp.Dock = "Top"
            $tlp.AutoSize = $true
            $tlp.ColumnCount = 2
            $tlp.BackColor = $geminiPanelBg
            $tlp.Padding = New-Object System.Windows.Forms.Padding(4)
            
            $tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
            $tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
            
            $scrollPanel.Controls.Add($tlp)

            return $groupbox
        }

        # --- Create all GroupBoxes ---
        $gbOs = New-InfoGroupBox "Operating System" $geminiBlue
        $gbCpu = New-InfoGroupBox "Processor" $geminiYellow
        $gbRam = New-InfoGroupBox "Memory (RAM)" $geminiBlue
        $gbHardware = New-InfoGroupBox "System Hardware" $geminiYellow
        $gbGpu = New-InfoGroupBox "Video Card(s)" $geminiBlue
        $gbNetwork = New-InfoGroupBox "Network Adapters" $geminiYellow
        $gbDisk = New-InfoGroupBox "Disk Drives" $geminiBlue

        # --- Add GroupBoxes to the layout ---
        $mainTableLayout.Controls.Add($gbOs, 0, 0)
        $mainTableLayout.Controls.Add($gbCpu, 0, 1)
        $mainTableLayout.Controls.Add($gbRam, 0, 2)
        $mainTableLayout.Controls.Add($gbHardware, 0, 3)
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

        # --- Progress Bar for operations ---
        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Size = New-Object System.Drawing.Size(200, 18)
        $progressBar.Style = "Continuous"
        $progressBar.Visible = $false
        $buttonPanel.Controls.Add($progressBar)

        # --- Helper function to create styled Buttons with hover effects ---
        function New-ActionButton($Text, $BackColor, $HoverColor) {
            $button = New-Object System.Windows.Forms.Button
            $button.Text = $Text
            $button.Size = "130,30"
            $button.Font = $buttonFont
            $button.ForeColor = $geminiWhiteText
            $button.BackColor = $BackColor
            $button.FlatStyle = "Flat"
            $button.FlatAppearance.BorderSize = 0
            $button.Cursor = [System.Windows.Forms.Cursors]::Hand
            $button.Margin = New-Object System.Windows.Forms.Padding(4, 0, 4, 0)
            
            # Store original colors in a simple way
            if (-not $script:originalColors) { $script:originalColors = @{} }
            $script:originalColors[$button] = @{
                OriginalColor = $BackColor
                HoverColor = $HoverColor
                OriginalFont = $buttonFont
            }
            
            # Add subtle hover effects
            $button.Add_MouseEnter({
                try {
                    if ($script:originalColors.ContainsKey($this)) {
                        $colors = $script:originalColors[$this]
                        $this.BackColor = $colors.HoverColor
                        $this.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
                    }
                } catch {
                    # Silently ignore hover errors
                }
            })
            $button.Add_MouseLeave({
                try {
                    if ($script:originalColors.ContainsKey($this)) {
                        $colors = $script:originalColors[$this]
                        $this.BackColor = $colors.OriginalColor
                        $this.Font = $colors.OriginalFont
                    }
                } catch {
                    # Silently ignore hover errors
                }
            })
            
            return $button
        }

        $buttonCopy = New-ActionButton "Copy to Clipboard" $geminiYellow ([System.Drawing.Color]::FromArgb(240, 200, 70))
        $buttonRefresh = New-ActionButton "Refresh" $geminiBlue ([System.Drawing.Color]::FromArgb(80, 120, 200))
        $buttonExport = New-ActionButton "Export to File" $geminiAccent ([System.Drawing.Color]::FromArgb(50, 220, 255))
        
        $buttonPanel.Controls.AddRange(@($buttonCopy, $buttonRefresh, $buttonExport))

        # --- Data Population Logic ---
        $script:infoStore = @{}
        
        function Update-SystemInfo {
            try {
                $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
                $progressBar.Visible = $true
                $progressBar.Value = 10
                $script:infoStore.Clear()
                
                $populateGroupBox = {
                    param($GroupBox, $Data)
                    try {
                        $tlp = $GroupBox.Controls[0].Controls[0]
                        $tlp.Controls.Clear()
                        $tlp.RowCount = 0
                        
                        foreach ($item in $Data.GetEnumerator()) {
                            $label = New-Object System.Windows.Forms.Label
                            $label.Text = $item.Key
                            $label.Font = $labelFont
                            $label.ForeColor = $geminiGrayText
                            $label.AutoSize = $true
                            $label.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 3)
                            
                            $value = New-Object System.Windows.Forms.Label
                            $value.Text = $item.Value
                            $value.Font = $valueFont
                            $value.ForeColor = $geminiWhiteText
                            $value.AutoSize = $true
                            $value.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 3)
                            
                            # Add tooltip for long values
                            if ($value.Text.Length -gt 50) {
                                $tooltip = New-Object System.Windows.Forms.ToolTip
                                $tooltip.SetToolTip($value, $value.Text)
                            }
                            
                            $tlp.Controls.Add($label, 0, $tlp.RowCount)
                            $tlp.Controls.Add($value, 1, $tlp.RowCount)
                            
                            $tlp.RowCount += 1
                            $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
                        }
                    } catch {
                        Write-LogAndHost "Error populating group box: $($_.Exception.Message)" -HostColor Red -LogPrefix "Show-SystemInfo"
                    }
                }
                
                # --- OS Information ---
                $progressBar.Value = 20
                $osData = [ordered]@{}
                $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
                if ($osInfo) {
                    $osData["Name:"] = $osInfo.Caption
                    $script:infoStore["OS Name"] = $osInfo.Caption
                    $osData["Version:"] = $osInfo.Version
                    $script:infoStore["OS Version"] = $osInfo.Version
                    $osData["Build:"] = $osInfo.BuildNumber
                    $script:infoStore["OS Build"] = $osInfo.BuildNumber
                    $osData["Architecture:"] = $osInfo.OSArchitecture
                    $script:infoStore["OS Architecture"] = $osInfo.OSArchitecture
                    $osData["Install Date:"] = $osInfo.InstallDate.ToString("yyyy-MM-dd")
                    $script:infoStore["Install Date"] = $osInfo.InstallDate.ToString("yyyy-MM-dd")
                    $osData["Last Boot:"] = $osInfo.LastBootUpTime.ToString("yyyy-MM-dd HH:mm:ss")
                    $script:infoStore["Last Boot"] = $osInfo.LastBootUpTime.ToString("yyyy-MM-dd HH:mm:ss")
                    try { 
                        $productID = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue).ProductId
                        if ($productID) {
                            $osData["Product ID:"] = $productID
                            $script:infoStore["Product ID"] = $productID 
                        }
                    } catch {}
                    try { 
                        $editionID = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue).EditionID
                        if ($editionID) {
                            $osData["Edition:"] = $editionID
                            $script:infoStore["Edition"] = $editionID 
                        }
                    } catch {}
                }
                & $populateGroupBox $gbOs $osData
                
                # --- CPU Information ---
                $progressBar.Value = 30
                $cpuData = [ordered]@{}
                $cpuInfo = Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue
                if ($cpuInfo) {
                    $cpuData["Name:"] = $cpuInfo.Name.Trim()
                    $script:infoStore["CPU"] = $cpuInfo.Name.Trim()
                    $cpuData["Cores (Logical):"] = "$($cpuInfo.NumberOfCores) ($($cpuInfo.NumberOfLogicalProcessors))"
                    $script:infoStore["Cores"] = "$($cpuInfo.NumberOfCores) ($($cpuInfo.NumberOfLogicalProcessors))"
                    
                    # Removed clock speed and cache information as requested
                    
                    $virtEnabled = if ($cpuInfo.VirtualizationFirmwareEnabled) { "Enabled" } else { "Disabled" }
                    $cpuData["Virtualization:"] = "Firmware $virtEnabled"
                    $script:infoStore["Virtualization"] = "Firmware $virtEnabled"
                }
                & $populateGroupBox $gbCpu $cpuData

                # --- System Hardware ---
                $progressBar.Value = 40
                $hwData = [ordered]@{}
                $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
                if ($computerSystem) {
                    $hwData["Manufacturer:"] = $computerSystem.Manufacturer
                    $script:infoStore["Manufacturer"] = $computerSystem.Manufacturer
                    $hwData["Model:"] = $computerSystem.Model
                    $script:infoStore["Model"] = $computerSystem.Model
                    $hwData["System Type:"] = $computerSystem.SystemType
                    $script:infoStore["System Type"] = $computerSystem.SystemType
                }
                $boardInfo = Get-CimInstance -ClassName Win32_BaseBoard -ErrorAction SilentlyContinue
                if ($boardInfo) {
                    $hwData["Motherboard:"] = "$($boardInfo.Manufacturer) $($boardInfo.Product)"
                    $script:infoStore["Motherboard"] = "$($boardInfo.Manufacturer) $($boardInfo.Product)"
                }
                $biosInfo = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue
                if ($biosInfo) {
                    $hwData["BIOS Version:"] = $biosInfo.SMBIOSBIOSVersion
                    $script:infoStore["BIOS Version"] = $biosInfo.SMBIOSBIOSVersion
                    $hwData["BIOS Release Date:"] = $biosInfo.ReleaseDate.ToString("yyyy-MM-dd")
                    $script:infoStore["BIOS Release Date"] = $biosInfo.ReleaseDate.ToString("yyyy-MM-dd")
                }
                
                $secureBootStatus = try { if (Confirm-SecureBootUEFI -ErrorAction SilentlyContinue) { "Enabled" } else { "Disabled" } } catch { "Unsupported / Error" }
                $hwData["Secure Boot:"] = $secureBootStatus
                $script:infoStore["Secure Boot"] = $secureBootStatus
                
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
                $hwData["TPM Status:"] = $tpmStatus
                $script:infoStore["TPM Status"] = $tpmStatus

                & $populateGroupBox $gbHardware $hwData

                # --- Memory (RAM) ---
                $progressBar.Value = 50
                $ramData = [ordered]@{}
                if ($computerSystem) {
                    $ramGB = [math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2)
                    $ramData["Total Installed:"] = "$($ramGB) GB"
                    $script:infoStore["Total RAM"] = "$($ramGB) GB"
                }
                $osInfoForRam = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
                if ($osInfoForRam) {
                    $ramData["Available:"] = "$([math]::Round($osInfoForRam.FreePhysicalMemory / 1MB, 2)) GB"
                    $script:infoStore["Available RAM"] = "$([math]::Round($osInfoForRam.FreePhysicalMemory / 1MB, 2)) GB"
                }
                $ramData["Modules:"] = ""
                $memoryModules = Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction SilentlyContinue
                if ($memoryModules) {
                    $i = 1
                    foreach ($module in @($memoryModules)) {
                        $capacityGB = [math]::Round($module.Capacity / 1GB, 2)
                        $key = "- Slot $i ($($module.DeviceLocator)):"
                        $value = "$($capacityGB) GB, $($module.ConfiguredClockSpeed) MHz, $($module.Manufacturer)"
                        $ramData[$key] = $value
                        $script:infoStore["RAM Module $i"] = $value
                        $i++
                    }
                }
                & $populateGroupBox $gbRam $ramData
                
                # --- Video Cards ---
                $progressBar.Value = 60
                $gpuData = [ordered]@{}
                $videoControllers = Get-CimInstance -ClassName Win32_VideoController -ErrorAction SilentlyContinue
                if ($videoControllers) {
                    $i = 1
                    foreach ($video in @($videoControllers)) {
                        $gpuData["Name:"] = $video.Name
                        $script:infoStore["GPU $i Name"] = $video.Name
                        $adapterRamGB = $null
                        try {
                            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\"
                            $matchingKey = (Get-ChildItem -Path $regPath -EA SilentlyContinue) | Where-Object { 
                                ($_.GetValue("DriverDesc") -eq $video.Name) -or ($_.GetValue("Description") -eq $video.Name) 
                            } | Select-Object -First 1
                            if ($matchingKey) { 
                                $vramBytes = $matchingKey.GetValue("HardwareInformation.qwMemorySize")
                                if ($vramBytes -gt 0) { $adapterRamGB = [math]::Round($vramBytes / 1GB, 2) } 
                            }
                        } catch {}
                        if (-not $adapterRamGB -and $video.AdapterRAM) { 
                            $adapterRamGB = [math]::Round($video.AdapterRAM / 1GB, 2) 
                        }
                        $gpuData["Adapter RAM:"] = "$($adapterRamGB) GB"
                        $script:infoStore["GPU $i VRAM"] = "$($adapterRamGB) GB"
                        $gpuData["Driver Version:"] = $video.DriverVersion
                        $script:infoStore["GPU $i Driver"] = $video.DriverVersion
                        $gpuData["Driver Date:"] = $video.DriverDate.ToString("yyyy-MM-dd")
                        $script:infoStore["GPU $i Driver Date"] = $video.DriverDate.ToString("yyyy-MM-dd")
                        # Fixed resolution display with standard "x" instead of "×" to avoid encoding issues
                        $gpuData["Current Resolution:"] = "$($video.CurrentHorizontalResolution)x$($video.CurrentVerticalResolution) , $($video.CurrentRefreshRate)Hz"
                        $script:infoStore["GPU $i Resolution"] = "$($video.CurrentHorizontalResolution)x$($video.CurrentVerticalResolution) , $($video.CurrentRefreshRate)Hz"
                        $gpuData[" "] = ""
                        $i++
                    }
                }
                & $populateGroupBox $gbGpu $gpuData

                # --- Disk Information ---
                $progressBar.Value = 70
                $diskTlp = $gbDisk.Controls[0].Controls[0]
                $diskTlp.Controls.Clear()
                $diskTlp.RowCount = 0
                
                $disks = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue
                if ($disks) {
                    foreach ($disk in @($disks)) {
                        $sizeGB = [math]::Round($disk.Size / 1GB, 2)
                        $diskType = try { (Get-PhysicalDisk -DeviceNumber $disk.Index -EA SilentlyContinue).MediaType } catch { "Unknown" }
                        
                        $diskLabel = New-Object System.Windows.Forms.Label
                        $diskLabel.Text = "$($disk.Model) ($($sizeGB) GB) - $diskType"
                        $diskLabel.Font = $headerFont
                        $diskLabel.ForeColor = $gbDisk.ForeColor
                        $diskLabel.AutoSize = $true
                        $diskLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 3)
                        $diskTlp.Controls.Add($diskLabel, 0, $diskTlp.RowCount)
                        $diskTlp.SetColumnSpan($diskLabel, 2)
                        $diskTlp.RowCount += 1
                        $diskTlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
                        $script:infoStore["Disk $($disk.Index)"] = $diskLabel.Text

                        $partitions = Get-CimAssociatedInstance -InputObject $disk -ResultClassName Win32_DiskPartition -ErrorAction SilentlyContinue
                        if ($partitions) {
                            foreach ($partition in @($partitions)) {
                                $logicalDisk = Get-CimAssociatedInstance -InputObject $partition -ResultClassName Win32_LogicalDisk -ErrorAction SilentlyContinue
                                if ($logicalDisk) {
                                    $freeGB = [math]::Round($logicalDisk.FreeSpace / 1GB, 2)
                                    $usedGB = [math]::Round(($logicalDisk.Size - $logicalDisk.FreeSpace) / 1GB, 2)
                                    $percentFree = [math]::Round(($logicalDisk.FreeSpace / $logicalDisk.Size) * 100, 2)
                                    
                                    $driveLabel = New-Object System.Windows.Forms.Label
                                    $driveLabel.Text = "  - Drive $($logicalDisk.DeviceID):"
                                    $driveLabel.Font = $labelFont
                                    $driveLabel.ForeColor = $geminiGrayText
                                    $driveLabel.AutoSize = $true
                                    $driveLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 3)
                                    
                                    $sizeLabel = New-Object System.Windows.Forms.Label
                                    $sizeLabel.Text = "Size: $($usedGB) GB used / $($freeGB) GB free ($($percentFree)%)"
                                    $sizeLabel.Font = $valueFont
                                    $sizeLabel.ForeColor = $geminiWhiteText
                                    $sizeLabel.AutoSize = $true
                                    $sizeLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 3)
                                    
                                    # Add a mini progress bar for disk usage
                                    $diskUsageBar = New-Object System.Windows.Forms.ProgressBar
                                    $diskUsageBar.Value = 100 - $percentFree
                                    $diskUsageBar.Size = New-Object System.Drawing.Size(120, 8)
                                    $diskUsageBar.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 3)
                                    
                                    $diskTlp.Controls.Add($driveLabel, 0, $diskTlp.RowCount)
                                    $diskTlp.Controls.Add($sizeLabel, 1, $diskTlp.RowCount)
                                    $diskTlp.RowCount += 1
                                    $diskTlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
                                    
                                    $diskTlp.Controls.Add($diskUsageBar, 1, $diskTlp.RowCount)
                                    $diskTlp.SetColumnSpan($diskUsageBar, 2)
                                    $diskTlp.RowCount += 1
                                    $diskTlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
                                    
                                    $script:infoStore["Partition $($logicalDisk.DeviceID)"] = "Size: $($usedGB) GB used / $($freeGB) GB free ($($percentFree)%)"
                                }
                            }
                        }
                    }
                }
                
                # --- Network Information ---
                $progressBar.Value = 80
                $netData = [ordered]@{}
                $netAdapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue | Where-Object { $_.IPEnabled }
                if ($netAdapters) {
                    $i = 1
                    foreach ($adapter in @($netAdapters)) {
                        $netData["Description:"] = $adapter.Description
                        $script:infoStore["NIC $i Desc"] = $adapter.Description
                        $netData["IP Address:"] = ($adapter.IPAddress -join ', ')
                        $script:infoStore["NIC $i IP"] = ($adapter.IPAddress -join ', ')
                        $netData["MAC Address:"] = $adapter.MACAddress
                        $script:infoStore["NIC $i MAC"] = $adapter.MACAddress
                        $netData["Default Gateway:"] = ($adapter.DefaultIPGateway -join ', ')
                        $script:infoStore["NIC $i Gateway"] = ($adapter.DefaultIPGateway -join ', ')
                        $netData["DNS Servers:"] = ($adapter.DNSServerSearchOrder -join ', ')
                        $script:infoStore["NIC $i DNS"] = ($adapter.DNSServerSearchOrder -join ', ')
                        $netData["DHCP Enabled:"] = if ($adapter.DHCPEnabled) { "Yes" } else { "No" }
                        $script:infoStore["NIC $i DHCP"] = if ($adapter.DHCPEnabled) { "Yes" } else { "No" }
                        $netData[" "] = ""
                        $i++
                    }
                }
                & $populateGroupBox $gbNetwork $netData
                
                $progressBar.Value = 100
                Start-Sleep -Milliseconds 500
                $progressBar.Visible = $false

            } catch {
                Write-LogAndHost "Failed to gather system information. Error: $($_.Exception.Message)" -HostColor Red -LogPrefix "Show-SystemInfo"
                $progressBar.Visible = $false
            } finally {
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
            }
        }

        # --- Button Event Handlers ---
        $buttonRefresh.Add_Click({ Update-SystemInfo })
        
        $buttonCopy.Add_Click({
            try {
                $clipboardText = ""
                $sections = @(
                    "--- Operating System ---", "OS Name", "OS Version", "OS Build", "OS Architecture", "Install Date", "Last Boot", "Product ID", "Edition",
                    "--- Processor ---", "CPU", "Cores", "Virtualization",
                    "--- System Hardware ---", "Manufacturer", "Model", "System Type", "Motherboard", "BIOS Version", "BIOS Release Date", "Secure Boot", "TPM Status",
                    "--- Memory (RAM) ---", "Total RAM", "Available RAM",
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
                    $gpuNum = $_.Split(' ')[1]
                    $clipboardText += "`r`nGPU: $($script:infoStore[$_])`r`n"
                    $clipboardText += "  VRAM: $($script:infoStore["GPU $gpuNum VRAM"])`r`n"
                    $clipboardText += "  Driver: $($script:infoStore["GPU $gpuNum Driver"])`r`n"
                    $clipboardText += "  Driver Date: $($script:infoStore["GPU $gpuNum Driver Date"])`r`n"
                    $clipboardText += "  Resolution: $($script:infoStore["GPU $gpuNum Resolution"])`r`n"
                }
                ($script:infoStore.Keys | Where-Object { $_ -like "Disk *" } | Sort-Object) | ForEach-Object { 
                    $clipboardText += "`r`nDisk: $($script:infoStore[$_])`r`n"
                    ($script:infoStore.Keys | Where-Object { $_ -like "Partition*" } | Sort-Object) | ForEach-Object {
                         $clipboardText += "  - $($script:infoStore[$_])`r`n"
                    }
                }
                ($script:infoStore.Keys | Where-Object { $_ -like "NIC * Desc" } | Sort-Object) | ForEach-Object { 
                    $nicNum = $_.Split(' ')[1]
                    $clipboardText += "`r`nNIC: $($script:infoStore[$_])`r`n"
                    $clipboardText += "  IP: $($script:infoStore["NIC $nicNum IP"])`r`n"
                    $clipboardText += "  MAC: $($script:infoStore["NIC $nicNum MAC"])`r`n"
                    $clipboardText += "  Gateway: $($script:infoStore["NIC $nicNum Gateway"])`r`n"
                    $clipboardText += "  DNS: $($script:infoStore["NIC $nicNum DNS"])`r`n"
                    $clipboardText += "  DHCP: $($script:infoStore["NIC $nicNum DHCP"])`r`n"
                }
                Set-Clipboard -Value $clipboardText.Trim()
                [System.Windows.Forms.MessageBox]::Show("System information copied to clipboard.", "Success", "OK", "Information") | Out-Null
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to copy to clipboard. Error: $($_.Exception.Message)", "Error", "OK", "Error") | Out-Null
            }
        })
        
        $buttonExport.Add_Click({
            try {
                $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
                $saveFileDialog.Filter = "Text files (*.txt)|*.txt|JSON files (*.json)|*.json|CSV files (*.csv)|*.csv"
                $saveFileDialog.Title = "Export System Information"
                $result = $saveFileDialog.ShowDialog()
                
                if ($result -eq "OK") {
                    $extension = [System.IO.Path]::GetExtension($saveFileDialog.FileName)
                    
                    if ($extension -eq ".json") {
                        # Export as JSON
                        $json = $script:infoStore | ConvertTo-Json
                        $json | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
                    } elseif ($extension -eq ".csv") {
                        # Export as CSV
                        $script:infoStore.GetEnumerator() | Select-Object @{Name="Property";Expression={$_.Key}}, @{Name="Value";Expression={$_.Value}} | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
                    } else {
                        # Export as text (default)
                        $clipboardText = ""
                        $sections = @(
                            "--- Operating System ---", "OS Name", "OS Version", "OS Build", "OS Architecture", "Install Date", "Last Boot", "Product ID", "Edition",
                            "--- Processor ---", "CPU", "Cores", "Virtualization",
                            "--- System Hardware ---", "Manufacturer", "Model", "System Type", "Motherboard", "BIOS Version", "BIOS Release Date", "Secure Boot", "TPM Status",
                            "--- Memory (RAM) ---", "Total RAM", "Available RAM",
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
                            $gpuNum = $_.Split(' ')[1]
                            $clipboardText += "`r`nGPU: $($script:infoStore[$_])`r`n"
                            $clipboardText += "  VRAM: $($script:infoStore["GPU $gpuNum VRAM"])`r`n"
                            $clipboardText += "  Driver: $($script:infoStore["GPU $gpuNum Driver"])`r`n"
                            $clipboardText += "  Driver Date: $($script:infoStore["GPU $gpuNum Driver Date"])`r`n"
                            $clipboardText += "  Resolution: $($script:infoStore["GPU $gpuNum Resolution"])`r`n"
                        }
                        ($script:infoStore.Keys | Where-Object { $_ -like "Disk *" } | Sort-Object) | ForEach-Object { 
                            $clipboardText += "`r`nDisk: $($script:infoStore[$_])`r`n"
                            ($script:infoStore.Keys | Where-Object { $_ -like "Partition*" } | Sort-Object) | ForEach-Object {
                                 $clipboardText += "  - $($script:infoStore[$_])`r`n"
                            }
                        }
                        ($script:infoStore.Keys | Where-Object { $_ -like "NIC * Desc" } | Sort-Object) | ForEach-Object { 
                            $nicNum = $_.Split(' ')[1]
                            $clipboardText += "`r`nNIC: $($script:infoStore[$_])`r`n"
                            $clipboardText += "  IP: $($script:infoStore["NIC $nicNum IP"])`r`n"
                            $clipboardText += "  MAC: $($script:infoStore["NIC $nicNum MAC"])`r`n"
                            $clipboardText += "  Gateway: $($script:infoStore["NIC $nicNum Gateway"])`r`n"
                            $clipboardText += "  DNS: $($script:infoStore["NIC $nicNum DNS"])`r`n"
                            $clipboardText += "  DHCP: $($script:infoStore["NIC $nicNum DHCP"])`r`n"
                        }
                        $clipboardText | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
                    }
                    
                    [System.Windows.Forms.MessageBox]::Show("System information exported to $($saveFileDialog.FileName)", "Export Successful", "OK", "Information") | Out-Null
                }
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to export system information. Error: $($_.Exception.Message)", "Export Failed", "OK", "Error") | Out-Null
            }
        })

        # --- Initial Load and Show Form ---
        $form.Add_Shown({ 
            try {
                # Fade in effect
                $fadeIn = {
                    if ($form.Opacity -lt 1) {
                        $form.Opacity += 0.05
                        $form.Refresh()
                        Start-Sleep -Milliseconds 20
                        &$fadeIn
                    }
                }
                &$fadeIn
            } catch {
                # Ignore fade-in errors
            }
        })
        
        Update-SystemInfo

        $null = $form.ShowDialog()
        Write-LogAndHost "System information GUI closed by user." -NoHost
        
    } catch {
        Write-LogAndHost "An unexpected error occurred with the System Information GUI. Details: $($_.Exception.Message)" -HostColor Red
    } finally {
        try {
            if ($form) { $form.Dispose() }
        } catch {
            # Ignore disposal errors
        }
    }

    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}



#  FUNCTION : Imports a list of programs from a file and installs them.
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

function Invoke-TempFileCleanup {
    # Nested function to find dynamic application caches (enhanced)
    function Find-DynamicAppCaches {
        Write-LogAndHost "Scanning for application caches..." -NoHost -LogPrefix "Find-DynamicAppCaches"
        
        # Return structure with two categories: Application and System caches
        $result = @{
            ApplicationCaches = [ordered]@{}
            SystemCaches = [ordered]@{}
        }
        
        $processedApps = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    
        # --- Enhanced list of well-known application caches ---
        $wellKnownApps = [ordered]@{
            "NVIDIA Cache"             = @("$env:LOCALAPPDATA\NVIDIA\GLCache", "$env:ProgramData\NVIDIA Corporation\Downloader", "$env:LOCALAPPDATA\NVIDIA\DXCache");
            "DirectX Shader Cache"     = @("$env:LOCALAPPDATA\D3DSCache", "$env:LOCALAPPDATA\AMD\DxCache");
            "Steam Cache"              = @("$env:LOCALAPPDATA\Steam\appcache", "$env:LOCALAPPDATA\Steam\htmlcache", "$env:ProgramFiles\Steam\package");
            "Discord Cache"            = @("$env:APPDATA\discord\Cache", "$env:APPDATA\discord\Code Cache", "$env:APPDATA\discord\GPUCache", "$env:APPDATA\discord\Local Storage");
            "EA App Cache"             = @("$env:LOCALAPPDATA\Electronic Arts\EA Desktop\cache", "$env:ProgramData\Electronic Arts\EA Desktop\cache");
            "Spotify Cache"            = @("$env:LOCALAPPDATA\Spotify\Storage", "$env:LOCALAPPDATA\Spotify\Browser", "$env:LOCALAPPDATA\Spotify\Data");
            "Visual Studio Code Cache" = @("$env:APPDATA\Code\Cache", "$env:APPDATA\Code\GPUCache", "$env:APPDATA\Code\CachedData", "$env:APPDATA\Code\Local Storage");
            "Slack Cache"              = @("$env:APPDATA\Slack\Cache", "$env:APPDATA\Slack\GPUCache", "$env:APPDATA\Slack\Service Worker\CacheStorage", "$env:APPDATA\Slack\IndexedDB");
            "Zoom Cache"               = @("$env:APPDATA\Zoom\data\Cache", "$env:APPDATA\Zoom\data\logs");
            "Adobe Cache"              = @("$env:LOCALAPPDATA\Adobe\Common\Media Cache Files", "$env:LOCALAPPDATA\Adobe\Common\Media Cache", "$env:APPDATA\Adobe\Logs");
            "Telegram Cache"           = @("$env:APPDATA\Telegram Desktop\tdata\user_data\cache", "$env:APPDATA\Telegram Desktop\tdata\temp");
            "Microsoft Teams Cache"    = @("$env:APPDATA\Microsoft\Teams\Cache", "$env:APPDATA\Microsoft\Teams\blob_storage", "$env:APPDATA\Microsoft\Teams\databases");
            "OneDrive Cache"           = @("$env:LOCALAPPDATA\Microsoft\OneDrive\cache", "$env:LOCALAPPDATA\Microsoft\OneDrive\logs");
            "Dropbox Cache"            = @("$env:LOCALAPPDATA\Dropbox\cache", "$env:LOCALAPPDATA\Dropbox\logs");
            "Google Drive Cache"       = @("$env:LOCALAPPDATA\Google\DriveFS\cache", "$env:LOCALAPPDATA\Google\DriveFS\logs");
            "Firefox Cache"            = @("$env:APPDATA\Mozilla\Firefox\Profiles\*\cache2", "$env:APPDATA\Mozilla\Firefox\Profiles\*\startupCache");
        }
    
        Write-LogAndHost "Checking for well-known application caches..." -NoHost
        foreach ($appName in $wellKnownApps.Keys) {
            $existingPaths = @()
            foreach ($path in $wellKnownApps[$appName]) {
                # Check if path exists and has content
                if (Test-Path $path -PathType Container) {
                    $size = (Get-ChildItem $path -Recurse -Force -ErrorAction SilentlyContinue | 
                            Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                    if ($size -gt 0) {
                        $existingPaths += $path
                        Write-LogAndHost "Found cache for '$appName' at $path (Size: $([math]::Round($size/1MB, 2)) MB)" -NoHost
                    }
                }
            }
            
            if ($existingPaths.Count -gt 0) {
                $result.ApplicationCaches[$appName] = @{ Paths = $existingPaths; Type = 'Folder' }
                [void]$processedApps.Add($appName)
            }
        }
    
        # --- Enhanced browser cache detection ---
        Write-LogAndHost "Scanning for installed web browser caches..." -NoHost
        $browserCacheConfigs = [ordered]@{
            "Google Chrome" = @{
                InstallPaths = @(
                    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
                    "$env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe",
                    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
                );
                CacheDirs = @(
                    "$env:LOCALAPPDATA\Google\Chrome\User Data\*\Cache",
                    "$env:LOCALAPPDATA\Google\Chrome\User Data\*\Code Cache",
                    "$env:LOCALAPPDATA\Google\Chrome\User Data\*\GPUCache",
                    "$env:LOCALAPPDATA\Google\Chrome\User Data\*\Service Worker"
                )
            };
            "Microsoft Edge" = @{
                InstallPaths = @(
                    "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
                    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
                );
                CacheDirs = @(
                    "$env:LOCALAPPDATA\Microsoft\Edge\User Data\*\Cache",
                    "$env:LOCALAPPDATA\Microsoft\Edge\User Data\*\Code Cache",
                    "$env:LOCALAPPDATA\Microsoft\Edge\User Data\*\GPUCache",
                    "$env:LOCALAPPDATA\Microsoft\Edge\User Data\*\Service Worker"
                )
            };
            "Brave Browser" = @{
                InstallPaths = @(
                    "$env:ProgramFiles\BraveSoftware\Brave-Browser\Application\brave.exe",
                    "$env:ProgramFiles(x86)\BraveSoftware\Brave-Browser\Application\brave.exe"
                );
                CacheDirs = @(
                    "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data\*\Cache",
                    "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data\*\Code Cache",
                    "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data\*\GPUCache"
                )
            };
            "Opera Browser" = @{
                InstallPaths = @(
                    "$env:ProgramFiles\Opera\launcher.exe",
                    "$env:ProgramFiles(x86)\Opera\launcher.exe"
                );
                CacheDirs = @(
                    "$env:LOCALAPPDATA\Opera Software\Opera*\Cache",
                    "$env:LOCALAPPDATA\Opera Software\Opera*\Code Cache",
                    "$env:LOCALAPPDATA\Opera Software\Opera*\GPUCache"
                )
            };
            "Vivaldi Browser" = @{
                InstallPaths = @(
                    "$env:ProgramFiles\Vivaldi\Application\vivaldi.exe",
                    "$env:ProgramFiles(x86)\Vivaldi\Application\vivaldi.exe"
                );
                CacheDirs = @(
                    "$env:LOCALAPPDATA\Vivaldi\User Data\*\Cache",
                    "$env:LOCALAPPDATA\Vivaldi\User Data\*\Code Cache",
                    "$env:LOCALAPPDATA\Vivaldi\User Data\*\GPUCache"
                )
            };
            "Mozilla Firefox" = @{
                InstallPaths = @(
                    "$env:ProgramFiles\Mozilla Firefox\firefox.exe",
                    "$env:ProgramFiles(x86)\Mozilla Firefox\firefox.exe"
                );
                CacheDirs = @() # Handled separately
            }
        }
    
        foreach ($browserName in $browserCacheConfigs.Keys) {
            $config = $browserCacheConfigs[$browserName]
            $isBrowserInstalled = $false
            
            # Check if browser is installed
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
                            $cachePaths = @(
                                "$($_.FullName)\cache2",
                                "$($_.FullName)\startupCache",
                                "$($_.FullName)\offlineCache"
                            )
                            foreach ($cachePath in $cachePaths) {
                                if (Test-Path $cachePath) {
                                    $size = (Get-ChildItem $cachePath -Recurse -Force -ErrorAction SilentlyContinue | 
                                            Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                                    if ($size -gt 0) {
                                        [void]$foundBrowserPaths.Add($cachePath)
                                        Write-LogAndHost "Found Firefox cache at $cachePath (Size: $([math]::Round($size/1MB, 2)) MB)" -NoHost
                                    }
                                }
                            }
                        }
                    }
                } else {
                    # For Chromium-based browsers - resolve wildcards
                    foreach ($cacheDir in $config.CacheDirs) {
                        try {
                            $resolvedPaths = Resolve-Path $cacheDir -ErrorAction SilentlyContinue
                            if ($resolvedPaths) {
                                foreach ($path in $resolvedPaths) {
                                    if (Test-Path $path -PathType Container) {
                                        $size = (Get-ChildItem $path -Recurse -Force -ErrorAction SilentlyContinue | 
                                                Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                                        if ($size -gt 0) {
                                            [void]$foundBrowserPaths.Add($path.Path)
                                            Write-LogAndHost "Found $browserName cache at $($path.Path) (Size: $([math]::Round($size/1MB, 2)) MB)" -NoHost
                                        }
                                    }
                                }
                            }
                        } catch {}
                    }
                }
    
                if ($foundBrowserPaths.Count -gt 0) {
                    $uniquePaths = $foundBrowserPaths | Select-Object -Unique
                    $result.ApplicationCaches[$browserName + " Cache"] = @{ Paths = $uniquePaths; Type = 'Folder' }
                    [void]$processedApps.Add($browserName)
                }
            }
        }
    
        # --- Enhanced registry scanning for other applications ---
        Write-LogAndHost "Scanning registry for other installed applications..." -NoHost
        $uninstallPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        )
        
        # Extended list of cache folder names
        $cacheFolderNames = @(
            "Cache", "cache", "Code Cache", "GPUCache", "ShaderCache", 
            "temp", "tmp", "logs", "Logs", "CrashDumps", "dumps", 
            "Blob Storage", "blob_storage", "Local Storage", "localstorage", 
            "IndexedDB", "databases", "Service Worker", "ServiceWorker", 
            "Application Cache", "appcache", "offlineCache"
        )
        
        # Extended search roots
        $searchRoots = @($env:LOCALAPPDATA, $env:APPDATA, $env:ProgramData, "$env:ProgramFiles", "${env:ProgramFiles(x86)}")
    
        foreach ($path in $uninstallPaths) {
            $regKeys = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
            if (-not $regKeys) { continue }
    
            foreach ($key in $regKeys) {
                $appName = $key.GetValue("DisplayName")
                $publisher = $key.GetValue("Publisher")
                $installLocation = $key.GetValue("InstallLocation")
    
                if ([string]::IsNullOrWhiteSpace($appName) -or $appName -match "^KB[0-9]{6,}$" -or $processedApps.Contains($appName)) { 
                    continue 
                }
                [void]$processedApps.Add($appName)
    
                $potentialNames = New-Object System.Collections.Generic.List[string]
                $potentialNames.Add($appName)
                if (-not [string]::IsNullOrWhiteSpace($publisher)) { 
                    $potentialNames.Add($publisher) 
                }
                
                # Also check install location for caches
                if (-not [string]::IsNullOrWhiteSpace($installLocation) -and (Test-Path $installLocation)) {
                    $potentialNames.Add($installLocation)
                }
                
                $foundPaths = New-Object System.Collections.Generic.List[string]
    
                foreach ($name in ($potentialNames | Select-Object -Unique)) {
                    foreach ($root in $searchRoots) {
                        # Handle both direct paths and publisher subfolders
                        $basePath = Join-Path -Path $root -ChildPath $name
                        if (Test-Path $basePath) {
                            foreach ($cacheName in $cacheFolderNames) {
                                $cachePath = Join-Path -Path $basePath -ChildPath $cacheName
                                if (Test-Path $cachePath) {
                                    $size = (Get-ChildItem $cachePath -Recurse -Force -ErrorAction SilentlyContinue | 
                                            Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                                    if ($size -gt 0) {
                                        $foundPaths.Add($cachePath)
                                        Write-LogAndHost "Found cache for '$appName' at $cachePath (Size: $([math]::Round($size/1MB, 2)) MB)" -NoHost
                                    }
                                }
                            }
                        }
                        
                        # Also check for app-specific cache folders directly in root
                        $appCachePath = Join-Path -Path $root -ChildPath "$name\Cache"
                        if (Test-Path $appCachePath) {
                            $size = (Get-ChildItem $appCachePath -Recurse -Force -ErrorAction SilentlyContinue | 
                                    Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                            if ($size -gt 0) {
                                $foundPaths.Add($appCachePath)
                                Write-LogAndHost "Found cache for '$appName' at $appCachePath (Size: $([math]::Round($size/1MB, 2)) MB)" -NoHost
                            }
                        }
                    }
                }
                
                if ($foundPaths.Count -gt 0) {
                    $uniquePaths = $foundPaths | Select-Object -Unique
                    $result.ApplicationCaches[$appName] = @{ Paths = $uniquePaths; Type = 'Folder' }
                }
            }
        }
        
        # --- Additional system-wide cache locations (removed Startup and Temp) ---
        Write-LogAndHost "Scanning for additional system cache locations..." -NoHost
        # Note: Removed Startup and Temp folders to avoid duplication
        # Temp folders are now handled by "Windows Temporary Files"
        # Startup folders are not included as they typically don't contain significant cache files
        
        Write-LogAndHost "Finished cache scan. Found $($result.ApplicationCaches.Count) application caches and $($result.SystemCaches.Count) system caches." -NoHost
        return $result
    }

    # Main function body
    if (-not $script:guiAvailable) {
        Write-LogAndHost "GUI is not available, cannot launch the System Cleanup tool." -HostColor Red -LogPrefix "Invoke-TempFileCleanup"
        Start-Sleep -Seconds 2
        return
    }
    Write-LogAndHost "Launching System Cleanup GUI..." -HostColor Cyan

    # --- Enhanced DATA STRUCTURE for all cleanup items ---
    $cleanupItems = [ordered]@{
        "System Items" = [ordered]@{
            "Windows Temporary Files" = @{ Paths = @("$env:TEMP", "$env:windir\Temp", "$env:ProgramData\Temp"); Type = 'Folder' }
            "Windows Update Cache"    = @{ Paths = @("$env:windir\SoftwareDistribution\Download", "$env:windir\SoftwareDistribution\DataStore"); Type = 'Folder' }
            "Delivery Optimization"   = @{ Paths = @("$env:windir\SoftwareDistribution\DeliveryOptimization"); Type = 'Folder' }
            "Windows Log Files"       = @{ Paths = @("$env:windir\Logs", "$env:windir\Logs\CBS"); Type = 'Folder' }
            "System Minidump Files"   = @{ Paths = @("$env:windir\Minidump", "$env:LOCALAPPDATA\CrashDumps"); Type = 'Folder' }
            "Windows Prefetch Files"  = @{ Paths = @("$env:windir\Prefetch"); Type = 'Folder' }
            "Thumbnail Cache"         = @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"; Filter = "thumbcache_*.db"; Type = 'File' }
            "Windows Icon Cache"      = @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"; Filter = "iconcache_*.db"; Type = 'File' }
            "Windows Font Cache"      = @{ Paths = @("$env:windir\ServiceProfiles\LocalService\AppData\Local\FontCache"); Type = 'Folder' }
            "Microsoft Store Cache"   = @{ Paths = @("$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache"); Type = 'Folder' }
            "DNS Cache"               = @{ Type = 'Special' }
            "Recycle Bin"             = @{ Type = 'Special' }
            "Windows Error Reporting" = @{ Paths = @("$env:ProgramData\Microsoft\Windows\WER", "$env:LOCALAPPDATA\CrashDumps"); Type = 'Folder' }
            "Windows Search Data"     = @{ Paths = @("$env:ProgramData\Microsoft\Search\Data"); Type = 'Folder' }
            "Windows Defender History"= @{ Paths = @("$env:ProgramData\Microsoft\Windows Defender\Scans\History"); Type = 'Folder' }
            "Temporary Internet Files"= @{ Paths = @("$env:LOCALAPPDATA\Microsoft\Windows\INetCache"); Type = 'Folder' }
            "Windows Metro Apps Cache"= @{ 
                Paths = @(
                    "$env:LOCALAPPDATA\Packages\*\AC\INetCache",
                    "$env:LOCALAPPDATA\Packages\*\TempState",
                    "$env:LOCALAPPDATA\Packages\*\LocalCache"
                ); 
                Type = 'Folder' 
            }
            "Windows Upgrade Log Files" = @{ Paths = @("$env:windir\Panther"); Type = 'Folder' }
        }
        "Discovered Application Caches" = [ordered]@{}
    }

    # --- GUI Setup ---
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Perdanga System Cleanup"
    $form.Size = New-Object System.Drawing.Size(600, 720)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)

    $commonFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $controlBackColor = [System.Drawing.Color]::FromArgb(60, 60, 63)
    $controlForeColor = [System.Drawing.Color]::White

    # --- ToolTip Setup for displaying paths ---
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.AutoPopDelay = 20000
    $toolTip.InitialDelay = 700
    $toolTip.ReshowDelay = 500
    $toolTip.UseFading = $true
    $toolTip.UseAnimation = $true

    # --- TreeView Setup ---
    $treeView = New-Object System.Windows.Forms.TreeView
    $treeView.Location = New-Object System.Drawing.Point(15, 55)
    $treeView.Size = New-Object System.Drawing.Size(560, 360)
    $treeView.CheckBoxes = $true
    $treeView.Font = $commonFont
    $treeView.BackColor = $controlBackColor
    $treeView.ForeColor = $controlForeColor
    $treeView.BorderStyle = "FixedSingle"
    $treeView.FullRowSelect = $true
    $treeView.ShowNodeToolTips = $false
    $form.Controls.Add($treeView) | Out-Null

    # --- Populate TreeView with static items (only if they have content) ---
    foreach ($categoryName in $cleanupItems.Keys) {
        $parentNode = $treeView.Nodes.Add($categoryName, $categoryName)
        if ($categoryName -eq "System Items") {
            foreach ($itemName in $cleanupItems[$categoryName].Keys) {
                $itemConfig = $cleanupItems[$categoryName][$itemName]
                
                # Skip folder items with no paths
                if ($itemConfig.Type -eq 'Folder' -and $itemConfig.Paths.Count -eq 0) { 
                    continue 
                }
                
                $addItem = $false
                
                if ($itemConfig.Type -eq 'Folder') {
                    # Special case: Windows Metro Apps Cache - always add (will be checked during analysis)
                    if ($itemName -eq "Windows Metro Apps Cache") {
                        $addItem = $true
                    } else {
                        $hasContent = $false
                        foreach ($path in $itemConfig.Paths) {
                            if (Test-Path $path -PathType Container) {
                                # Check if the folder has any files or subfolders with content
                                $items = Get-ChildItem $path -Recurse -Force -ErrorAction SilentlyContinue
                                if ($items) {
                                    $size = ($items | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                                    if ($size -gt 0) {
                                        $hasContent = $true
                                        break
                                    }
                                }
                            }
                        }
                        $addItem = $hasContent
                    }
                }
                elseif ($itemConfig.Type -eq 'File') {
                    $filePath = Join-Path $itemConfig.Path $itemConfig.Filter
                    if (Test-Path $filePath) {
                        $file = Get-Item $filePath -Force -ErrorAction SilentlyContinue
                        if ($file.Length -gt 0) {
                            $addItem = $true
                        }
                    }
                }
                elseif ($itemConfig.Type -eq 'Special') {
                    $addItem = $true
                }
                
                if ($addItem) {
                    $childNode = $parentNode.Nodes.Add($itemName, $itemName)
                    $childNode.Tag = $itemConfig
                    $childNode.Checked = $false 
                }
            }
            $parentNode.Checked = $false
            if ($parentNode.Nodes.Count -gt 0) {
                $parentNode.Expand()
            }
        }
    }

    # --- Log Box Setup ---
    $logBox = New-Object System.Windows.Forms.RichTextBox
    $logBox.Location = New-Object System.Drawing.Point(15, 455)
    $logBox.Size = New-Object System.Drawing.Size(560, 150)
    $logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $logBox.BackColor = $controlBackColor
    $logBox.ForeColor = $controlForeColor
    $logBox.ReadOnly = $true
    $logBox.BorderStyle = "FixedSingle"
    $logBox.ScrollBars = "Vertical"
    $form.Controls.Add($logBox) | Out-Null
    
    # --- Progress Bar Setup ---
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(15, 425)
    $progressBar.Size = New-Object System.Drawing.Size(560, 20)
    $progressBar.Visible = $false
    $form.Controls.Add($progressBar) | Out-Null

    # --- Button Setup ---
    $buttonAnalyze = New-Object System.Windows.Forms.Button
    $buttonAnalyze.Text = "Analyze"
    $buttonAnalyze.Size = "120,30"
    $buttonAnalyze.Location = "100,625"
    
    $buttonClean = New-Object System.Windows.Forms.Button
    $buttonClean.Text = "Clean"
    $buttonClean.Size = "120,30"
    $buttonClean.Location = "235,625"
    
    $buttonClose = New-Object System.Windows.Forms.Button
    $buttonClose.Text = "Exit"
    $buttonClose.Size = "120,30"
    $buttonClose.Location = "370,625"
    
    # --- Select/Deselect All Buttons ---
    $buttonSelectAll = New-Object System.Windows.Forms.Button
    $buttonSelectAll.Text = "Select All"
    $buttonSelectAll.Size = "120,30"
    $buttonSelectAll.Location = "15,15"
    
    $buttonDeselectAll = New-Object System.Windows.Forms.Button
    $buttonDeselectAll.Text = "Deselect All"
    $buttonDeselectAll.Size = "120,30"
    $buttonDeselectAll.Location = "145,15"
    
    @( $buttonAnalyze, $buttonClean, $buttonClose, $buttonSelectAll, $buttonDeselectAll ) | ForEach-Object {
        $_.Font = $commonFont
        $_.ForeColor = [System.Drawing.Color]::White
        $_.FlatStyle = "Flat"
        $_.FlatAppearance.BorderSize = 0
        $form.Controls.Add($_) | Out-Null
    }
    
    $buttonAnalyze.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $buttonClean.BackColor = [System.Drawing.Color]::FromArgb(200, 70, 70)
    $buttonClose.BackColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $buttonSelectAll.BackColor = [System.Drawing.Color]::FromArgb(80, 150, 80)
    $buttonDeselectAll.BackColor = [System.Drawing.Color]::FromArgb(150, 80, 80)
    
    # --- Event Handlers & Logic ---
    $analyzedData = @{}
    $logWriter = { 
        param($Message, $Color = 'White')
        $logBox.SelectionStart = $logBox.TextLength
        $logBox.SelectionLength = 0
        $logBox.SelectionColor = $Color
        $logBox.AppendText("$(Get-Date -Format 'HH:mm:ss') - $Message`n")
        $logBox.ScrollToCaret()
    }
    $form.Tag = @{ logWriter = $logWriter }

    # --- DYNAMICALLY FIND AND POPULATE APP CACHES ---
    $form.Add_Shown({
        $logWriterFunc = $form.Tag.logWriter
        & $logWriterFunc "Scanning for caches (this may take a moment)..." 'LightBlue'
        $form.Update()

        $discoveredCaches = Find-DynamicAppCaches
        $appCacheNode = $treeView.Nodes["Discovered Application Caches"]
        $systemItemsNode = $treeView.Nodes["System Items"]

        $treeView.BeginUpdate()
        
        # Populate Application Caches
        if ($discoveredCaches.ApplicationCaches.Count -gt 0) {
            foreach ($appName in ($discoveredCaches.ApplicationCaches.Keys | Sort-Object)) {
                $itemConfig = $discoveredCaches.ApplicationCaches[$appName]
                $cleanupItems["Discovered Application Caches"][$appName] = $itemConfig
                $childNode = $appCacheNode.Nodes.Add($appName, $appName)
                $childNode.Tag = $itemConfig
                $childNode.Checked = $false 
                
                # Add size info to node text if available
                if ($itemConfig.Size) {
                    $sizeText = " ($([math]::Round($itemConfig.Size/1MB, 2)) MB)"
                    $childNode.Text = $appName + $sizeText
                }
            }
            $appCacheNode.Expand()
            $appCacheNode.Checked = $false 
        } else {
            $treeView.Nodes.Remove($appCacheNode)
        }
        
        # Note: System caches are no longer discovered to avoid duplication
        # All system items are now predefined in the "System Items" section
        
        $treeView.EndUpdate()
        
        $appCount = $discoveredCaches.ApplicationCaches.Count
        & $logWriterFunc "Scan complete. Found $appCount application caches. Review and select them for cleaning." 'Green'
        
        # Ensure the view is scrolled to the top after loading
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
                $allChecked = $true
                $noneChecked = $true
                foreach ($sibling in $parent.Nodes) { 
                    if ($sibling.Checked) { $noneChecked = $false } else { $allChecked = $false } 
                }
                if ($allChecked) { $parent.Checked = $true } elseif ($noneChecked) { $parent.Checked = $false } else { $parent.Checked = $false }
            }
        } finally {
            $updatingChecks = $false
        }
    })

    # --- Button Click Handlers ---
    $buttonSelectAll.add_Click({ 
        foreach ($node in $treeView.Nodes) { 
            $node.Checked = $true 
        } 
    }) | Out-Null
    
    $buttonDeselectAll.add_Click({ 
        foreach ($node in $treeView.Nodes) { 
            $node.Checked = $false 
        } 
    }) | Out-Null

    $buttonAnalyze.add_Click({
        $logBox.Clear()
        & $logWriter "Starting analysis..." 'Cyan'
        $analyzedData.Clear()
        $progressBar.Value = 0
        $progressBar.Visible = $true
        $form.Update()
        
        $nodesToAnalyze = @()
        foreach ($parentNode in $treeView.Nodes) {
            foreach ($childNode in $parentNode.Nodes) {
                if ($childNode.Checked) { $nodesToAnalyze += $childNode }
            }
        }
        if ($nodesToAnalyze.Count -eq 0) { 
            & $logWriter "No items selected for analysis." 'Yellow'
            $progressBar.Visible = $false
            return 
        }

        $progressBar.Maximum = $nodesToAnalyze.Count
        $treeView.BeginUpdate()
        $totalSize = 0
        
        foreach ($childNode in $nodesToAnalyze) {
            $itemConfig = $childNode.Tag
            $itemSize = 0
            switch ($itemConfig.Type) {
                'Folder' { 
                    foreach ($p in $itemConfig.Paths) {
                        if(Test-Path $p) {
                            $itemSize += (Get-ChildItem $p -Recurse -Force -EA SilentlyContinue | Measure-Object -Property Length -Sum -EA SilentlyContinue).Sum 
                        } 
                    } 
                }
                'File' { 
                    if(Test-Path $itemConfig.Path) {
                        $itemSize = (Get-ChildItem $itemConfig.Path -Filter $itemConfig.Filter -Force -EA SilentlyContinue | Measure-Object -Property Length -Sum -EA SilentlyContinue).Sum 
                    } 
                }
                'Special' { 
                    if ($childNode.Name -eq "Recycle Bin") {
                        $shell = New-Object -ComObject Shell.Application
                        $recycleBin = $shell.NameSpace(0xA)
                        $itemSize = ($recycleBin.Items() | ForEach-Object { $_.Size } | Measure-Object -Sum).Sum 
                    }
                    # Special items like DNS cache have no size to measure.
                }
            }
            $analyzedData[$childNode.Name] = $itemSize
            $totalSize += $itemSize
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
        if ($analyzedData.Count -eq 0) { 
            [System.Windows.Forms.MessageBox]::Show("Please run an analysis first.", "Analysis Required", "OK", "Information") | Out-Null
            return 
        }
        
        $confirmResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to permanently delete these files?", "Confirm Deletion", "YesNo", "Warning")
        if ($confirmResult -ne 'Yes') { 
            & $logWriter "Cleanup cancelled by user." 'Yellow'
            return 
        }
        
        & $logWriter "Starting cleanup..." 'Cyan'
        $progressBar.Value = 0
        $progressBar.Visible = $true
        @( $buttonAnalyze, $buttonClean, $buttonSelectAll, $buttonDeselectAll ) | ForEach-Object { $_.Enabled = $false }
        $form.Update()
        
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
                    'Folder' { 
                        foreach ($p in $itemConfig.Paths) { 
                            if (Test-Path $p) { 
                                # Enhanced folder cleanup with wildcard support
                                $itemsToRemove = Get-ChildItem -Path $p -Recurse -Force -ErrorAction SilentlyContinue
                                if ($itemsToRemove) {
                                    Remove-Item -Path $itemsToRemove.FullName -Recurse -Force -ErrorAction SilentlyContinue
                                }
                            } 
                        } 
                    }
                    'File' { 
                        if (Test-Path $itemConfig.Path) { 
                            Remove-Item -Path (Join-Path $itemConfig.Path $itemConfig.Filter) -Force -ErrorAction SilentlyContinue 
                        } 
                    }
                    'Special' { 
                        if ($baseNodeName -eq "Recycle Bin") { 
                            Clear-RecycleBin -Force -ErrorAction SilentlyContinue 
                        }
                        if ($baseNodeName -eq "DNS Cache") { 
                            Start-Process -FilePath "ipconfig" -ArgumentList "/flushdns" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue 
                        }
                    }
                }
                if ($analyzedData.ContainsKey($childNode.Name)) {
                    $totalDeleted += $analyzedData[$childNode.Name]
                }
            } catch { 
                & $logWriter "Failed to clean $($baseNodeName). Error: $($_.Exception.Message)" 'Red' 
            }
            $progressBar.PerformStep()
        }
        
        $deletedInMB = [math]::Round($totalDeleted / 1MB, 2)
        & $logWriter "Cleanup complete. Freed approximately $deletedInMB MB of space." 'Green'
        & $logWriter "Perdanga Forever!" 'Magenta'
        
        $analyzedData.Clear()
        $progressBar.Visible = $false
        @( $buttonAnalyze, $buttonClean, $buttonSelectAll, $buttonDeselectAll ) | ForEach-Object { $_.Enabled = $true }
        
        foreach ($node in $treeView.Nodes) { 
            $node.Text = $node.Name -replace " \(\d+(\.\d+)? (MB|KB)\)$"
            foreach($child in $node.Nodes) { 
                $child.Text = $child.Name -replace " \(\d+(\.\d+)? (MB|KB)\)$" 
            } 
        }
    })

    $buttonClose.add_Click({ $form.Close() }) | Out-Null

    try { 
        $null = $form.ShowDialog() 
    } catch { 
        Write-LogAndHost "An unexpected error occurred with the System Cleanup GUI. Details: $($_.Exception.Message)" -HostColor Red 
    } finally { 
        $form.Dispose() 
    }
    
    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}

function Invoke-PerdangaSystemManager {
    <#
    .SYNOPSIS
        A comprehensive GUI-based manager for Power Plans, System Tweaks, Maintenance, Software, and Windows Update.
    .DESCRIPTION
        VERSION 5.8:
        - FIX: "Export List" no longer hangs the process (moved MessageBox to UI thread).
        - UPDATE: "MENU" text is now centered in the sidebar.
        - NEW: Added "Export List" to Install Apps (Saves detected apps to Desktop).
        - NEW: Added "Rebuild Icon Cache" to Maintenance tools.
        - NEW: Added "Refresh List" button to Install Apps tab.
        - NEW: Added "Update Selected" button to Install Apps tab.
        - FIX: Chocolatey detection logic improved (Path/Output parsing).
        - UPDATE: Installed software highlighting (Green/Gray status).
    #>
    
    # --- 1. PREREQUISITES CHECK ---
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "[ERROR] Script requires Administrator privileges." -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show("This tool requires Administrator privileges.`nPlease re-run PowerShell as an Administrator.", "Privilege Error", "OK", "Error") | Out-Null
        return
    }

    # Load Assemblies
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    Write-Host "[INIT] Starting Perdanga System Manager" -ForegroundColor Cyan

    # --- 2. SHARED STATE (SYNCHRONIZED) ---
    $sync = [Hashtable]::Synchronized(@{})
    $sync.Form = $null
    $sync.PowerPlans = @()
    $sync.InstalledApps = @()
    $sync.AppliedTweaks = @()
    $sync.StatusMessage = "Ready"
    $sync.IsBusy = $false
    $sync.PackageManager = "winget" 
    $sync.SoftwareSearchText = ""
    $sync.WUStatus = "Checking..."
    $sync.MenuButtons = @{}
    $sync.ActionQueue = @() 
    
    # --- 3. DATABASE DEFINITIONS ---
    $sync.Tweaks = @(
        # --- ESSENTIAL / PRIVACY ---
        @{ 
            Id="WPFTweaksTele"; Name="Disable Telemetry"; Category="Privacy"; Description="Disables DiagTrack, WAP Push, and extensive data collection."; 
            ScheduledTask=@(
                @{ Name="Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser"; State="Disabled"; OriginalState="Enabled" },
                @{ Name="Microsoft\Windows\Application Experience\ProgramDataUpdater"; State="Disabled"; OriginalState="Enabled" },
                @{ Name="Microsoft\Windows\Autochk\Proxy"; State="Disabled"; OriginalState="Enabled" },
                @{ Name="Microsoft\Windows\Customer Experience Improvement Program\Consolidator"; State="Disabled"; OriginalState="Enabled" },
                @{ Name="Microsoft\Windows\Customer Experience Improvement Program\UsbCeip"; State="Disabled"; OriginalState="Enabled" },
                @{ Name="Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector"; State="Disabled"; OriginalState="Enabled" },
                @{ Name="Microsoft\Windows\Feedback\Siuf\DmClient"; State="Disabled"; OriginalState="Enabled" },
                @{ Name="Microsoft\Windows\Feedback\Siuf\DmClientOnScenarioDownload"; State="Disabled"; OriginalState="Enabled" },
                @{ Name="Microsoft\Windows\Windows Error Reporting\QueueReporting"; State="Disabled"; OriginalState="Enabled" },
                @{ Name="Microsoft\Windows\Application Experience\MareBackup"; State="Disabled"; OriginalState="Enabled" },
                @{ Name="Microsoft\Windows\Application Experience\StartupAppTask"; State="Disabled"; OriginalState="Enabled" },
                @{ Name="Microsoft\Windows\Application Experience\PcaPatchDbTask"; State="Disabled"; OriginalState="Enabled" },
                @{ Name="Microsoft\Windows\Maps\MapsUpdateTask"; State="Disabled"; OriginalState="Enabled" }
            );
            Registry=@(
                @{ Path="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection"; Name="AllowTelemetry"; Value="0"; OriginalValue="<RemoveEntry>"; Type="DWord" },
                @{ Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"; Name="AllowTelemetry"; Value="0"; OriginalValue="<RemoveEntry>"; Type="DWord" },
                @{ Path="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name="ContentDeliveryAllowed"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name="OemPreInstalledAppsEnabled"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name="PreInstalledAppsEnabled"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name="PreInstalledAppsEverEnabled"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name="SilentInstalledAppsEnabled"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name="SubscribedContent-338387Enabled"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name="SubscribedContent-338388Enabled"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name="SubscribedContent-338389Enabled"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name="SubscribedContent-353698Enabled"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name="SystemPaneSuggestionsEnabled"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"; Name="DoNotShowFeedbackNotifications"; Value="1"; OriginalValue="<RemoveEntry>"; Type="DWord" },
                @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"; Name="DisableTailoredExperiencesWithDiagnosticData"; Value="1"; OriginalValue="<RemoveEntry>"; Type="DWord" },
                @{ Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo"; Name="DisabledByGroupPolicy"; Value="1"; OriginalValue="<RemoveEntry>"; Type="DWord" }
            )
        },
        @{ 
            Id="WPFTweaksConsumerFeatures"; Name="Disable Consumer Features"; Category="Privacy"; Description="Stops Windows from auto-installing sponsored apps."; 
            Registry=@( @{ Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"; Name="DisableWindowsConsumerFeatures"; Value="1"; OriginalValue="<RemoveEntry>"; Type="DWord" } )
        },
        @{ 
            Id="WPFTweaksActivity"; Name="Disable Activity History"; Category="Privacy"; Description="Stops Windows from tracking launched apps and docs."; 
            Registry=@( 
                @{ Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"; Name="EnableActivityFeed"; Value="0"; OriginalValue="<RemoveEntry>"; Type="DWord" },
                @{ Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"; Name="PublishUserActivities"; Value="0"; OriginalValue="<RemoveEntry>"; Type="DWord" },
                @{ Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"; Name="UploadUserActivities"; Value="0"; OriginalValue="<RemoveEntry>"; Type="DWord" }
            ) 
        },
        @{ 
            Id="WPFTweaksLoc"; Name="Disable Location Tracking"; Category="Privacy"; Description="Disables system-wide location tracking services."; 
            Registry=@( 
                @{ Path="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location"; Name="Value"; Value="Deny"; OriginalValue="Allow"; Type="String" },
                @{ Path="HKLM:\SYSTEM\CurrentControlSet\Services\lfsvc\Service\Configuration"; Name="Status"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKLM:\SYSTEM\Maps"; Name="AutoUpdateEnabled"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Sensor\Overrides\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}"; Name="SensorPermissionState"; Value="0"; OriginalValue="1"; Type="DWord" }
            )
        },
        @{ 
            Id="WPFTweaksWifi"; Name="Disable Wi-Fi Sense"; Category="Privacy"; Description="Prevents sharing Wi-Fi credentials with Facebook/Outlook contacts."; 
            Registry=@( 
                @{ Path="HKLM:\Software\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting"; Name="Value"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKLM:\Software\Microsoft\PolicyManager\default\WiFi\AllowAutoConnectToWiFiSenseHotspots"; Name="Value"; Value="0"; OriginalValue="1"; Type="DWord" }
            ) 
        },
        @{ 
            Id="WPFTweaksDisableNotifications"; Name="Disable Notifications"; Category="Privacy"; Description="Disables the Notification Center and Toast popups."; 
            Registry=@( 
                @{ Path="HKCU:\Software\Policies\Microsoft\Windows\Explorer"; Name="DisableNotificationCenter"; Value="1"; OriginalValue="<RemoveEntry>"; Type="DWord" },
                @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\PushNotifications"; Name="ToastEnabled"; Value="0"; OriginalValue="1"; Type="DWord" }
            )
        },

        # --- SYSTEM OPTIMIZATION ---
        @{ 
            Id="WPFTweaksServices"; Name="Optimize Services"; Category="System"; Description="Sets 100+ unused services to MANUAL (Safe Optimization)."; 
            Service=@(
                @{Name="AJRouter"; StartupType="Disabled"; OriginalType="Manual"},
                @{Name="ALG"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="AppIDSvc"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="AppMgmt"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="AppReadiness"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="AppXSvc"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="Appinfo"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="AudioEndpointBuilder"; StartupType="Automatic"; OriginalType="Automatic"},
                @{Name="AudioSrv"; StartupType="Automatic"; OriginalType="Automatic"},
                @{Name="BITS"; StartupType="AutomaticDelayedStart"; OriginalType="Automatic"},
                @{Name="Browser"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="DiagTrack"; StartupType="Disabled"; OriginalType="Automatic"},
                @{Name="DialogBlockingService"; StartupType="Disabled"; OriginalType="Disabled"},
                @{Name="Dnscache"; StartupType="Automatic"; OriginalType="Automatic"},
                @{Name="DPS"; StartupType="Automatic"; OriginalType="Automatic"},
                @{Name="FDResPub"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="Fax"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="HomeGroupListener"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="HomeGroupProvider"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="InstallService"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="InventorySvc"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="LanmanServer"; StartupType="Automatic"; OriginalType="Automatic"},
                @{Name="LanmanWorkstation"; StartupType="Automatic"; OriginalType="Automatic"},
                @{Name="MapsBroker"; StartupType="AutomaticDelayedStart"; OriginalType="Automatic"},
                @{Name="Netlogon"; StartupType="Automatic"; OriginalType="Automatic"},
                @{Name="Netman"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="PcaSvc"; StartupType="Manual"; OriginalType="Automatic"},
                @{Name="PlugPlay"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="PrintNotify"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="RemoteRegistry"; StartupType="Disabled"; OriginalType="Disabled"},
                @{Name="RpcSs"; StartupType="Automatic"; OriginalType="Automatic"},
                @{Name="Spooler"; StartupType="Automatic"; OriginalType="Automatic"},
                @{Name="SysMain"; StartupType="Automatic"; OriginalType="Automatic"},
                @{Name="TrkWks"; StartupType="Automatic"; OriginalType="Automatic"},
                @{Name="W32Time"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="WSearch"; StartupType="AutomaticDelayedStart"; OriginalType="Automatic"},
                @{Name="WaaSMedicSvc"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="WinDefend"; StartupType="Automatic"; OriginalType="Automatic"},
                @{Name="WinHttpAutoProxySvc"; StartupType="Manual"; OriginalType="Manual"},
                @{Name="Winmgmt"; StartupType="Automatic"; OriginalType="Automatic"},
                @{Name="WlanSvc"; StartupType="Automatic"; OriginalType="Automatic"},
                @{Name="wuauserv"; StartupType="Manual"; OriginalType="Manual"}
            )
        },
        @{ 
            Id="WPFTweaksDVR"; Name="Disable GameDVR"; Category="Gaming"; Description="Disables Xbox Game Recording."; 
            Registry=@( 
                @{ Path="HKCU:\System\GameConfigStore"; Name="GameDVR_Enabled"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKCU:\System\GameConfigStore"; Name="GameDVR_FSEBehavior"; Value="2"; OriginalValue="0"; Type="DWord" },
                @{ Path="HKCU:\System\GameConfigStore"; Name="GameDVR_HonorUserFSEBehaviorMode"; Value="1"; OriginalValue="0"; Type="DWord" },
                @{ Path="HKCU:\System\GameConfigStore"; Name="GameDVR_EFSEFeatureFlags"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR"; Name="AllowGameDVR"; Value="0"; OriginalValue="1"; Type="DWord" }
            ) 
        },
        @{ 
            Id="WPFTweaksHiber"; Name="Disable Hibernation"; Category="System"; Description="Disables hiberfil.sys to save disk space."; 
            Registry=@(
                @{ Path="HKLM:\System\CurrentControlSet\Control\Session Manager\Power"; Name="HibernateEnabled"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FlyoutMenuSettings"; Name="ShowHibernateOption"; Value="0"; OriginalValue="1"; Type="DWord" }
            );
            InvokeScript="powercfg /hibernate off";
            UndoScript="powercfg /hibernate on";
        },
        @{ 
            Id="WPFToggleNumLock"; Name="Enable NumLock on Boot"; Category="System"; Description="Forces NumLock to be on when Windows starts."; 
            Registry=@( 
                @{ Path="HKU:\.Default\Control Panel\Keyboard"; Name="InitialKeyboardIndicators"; Value="2"; OriginalValue="0"; Type="String" },
                @{ Path="HKCU:\Control Panel\Keyboard"; Name="InitialKeyboardIndicators"; Value="2"; OriginalValue="0"; Type="String" }
            ) 
        },
        @{ 
            Id="WPFTweaksUTC"; Name="Set Time to UTC"; Category="System"; Description="Fixes time sync issues when dual-booting with Linux."; 
            Registry=@( @{ Path="HKLM:\SYSTEM\CurrentControlSet\Control\TimeZoneInformation"; Name="RealTimeIsUniversal"; Value="1"; OriginalValue="0"; Type="DWord" } ) 
        },
        @{ 
            Id="WPFTweaksStorage"; Name="Disable Storage Sense"; Category="System"; Description="Prevents Windows from automatically deleting temp files."; 
            Registry=@( @{ Path="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy"; Name="01"; Value="0"; OriginalValue="1"; Type="DWord" } ) 
        },

        # --- NETWORK ---
        @{ 
            Id="WPFTweaksDisableipsix"; Name="Disable IPv6"; Category="Network"; Description="Disables IPv6 to force IPv4."; 
            Registry=@( @{ Path="HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters"; Name="DisabledComponents"; Value="255"; OriginalValue="0"; Type="DWord" } );
            InvokeScript="Disable-NetAdapterBinding -Name '*' -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue";
            UndoScript="Enable-NetAdapterBinding -Name '*' -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue"
        },

        # --- APPEARANCE & EXPLORER ---
        @{ 
            Id="WPFToggleDarkMode"; Name="Enable Dark Mode"; Category="Appearance"; Description="Sets Apps and System to Dark Theme"; 
            Registry=@( 
                @{ Path="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"; Name="AppsUseLightTheme"; Value="0"; OriginalValue="1"; Type="DWord" }, 
                @{ Path="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"; Name="SystemUsesLightTheme"; Value="0"; OriginalValue="1"; Type="DWord" } 
            ) 
        },
        @{ 
            Id="WPFToggleBingSearch"; Name="Disable Bing Search"; Category="Appearance"; Description="Removes Bing results from Start Menu search."; 
            Registry=@( @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"; Name="BingSearchEnabled"; Value="0"; OriginalValue="1"; Type="DWord" } ) 
        },
        @{ 
            Id="WPFTweaksRightClickMenu"; Name="Classic Right-Click"; Category="Appearance"; Description="Restores Windows 10 style context menu (Win 11)."; 
            InvokeScript="New-Item -Path 'HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}' -Name 'InprocServer32' -Force -Value '' | Out-Null; Stop-Process -Name 'explorer' -Force";
            UndoScript="Remove-Item -Path 'HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}' -Recurse -Force -ErrorAction SilentlyContinue; Stop-Process -Name 'explorer' -Force";
        },
        @{ 
            Id="WPFToggleTaskbarAlignment"; Name="Align Taskbar Left"; Category="Appearance"; Description="Moves Windows 11 Taskbar icons to the left."; 
            Registry=@( @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name="TaskbarAl"; Value="0"; OriginalValue="1"; Type="DWord" } ) 
        },
        @{ 
            Id="WPFTweaksDisableExplorerAutoDiscovery"; Name="Disable Explorer Folder Discovery"; Category="Explorer"; Description="Stops Explorer from trying to guess folder content type."; 
            InvokeScript="Remove-Item 'HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags' -Recurse -Force -ErrorAction SilentlyContinue; Remove-Item 'HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\BagMRU' -Recurse -Force -ErrorAction SilentlyContinue; $allFolders = 'HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags\AllFolders\Shell'; if (!(Test-Path $allFolders)) { New-Item -Path $allFolders -Force | Out-Null }; New-ItemProperty -Path $allFolders -Name 'FolderType' -Value 'NotSpecified' -PropertyType String -Force | Out-Null; Stop-Process -Name 'explorer' -Force";
            UndoScript="Remove-Item 'HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags' -Recurse -Force -ErrorAction SilentlyContinue; Remove-Item 'HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\BagMRU' -Recurse -Force -ErrorAction SilentlyContinue; Stop-Process -Name 'explorer' -Force"
        },
        @{ 
            Id="WPFToggleHiddenFiles"; Name="Show Hidden Files"; Category="Explorer"; Description="Makes hidden files visible in Explorer."; 
            Registry=@( @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name="Hidden"; Value="1"; OriginalValue="0"; Type="DWord" } );
            InvokeScript="Stop-Process -Name 'explorer' -Force"
        },
        @{ 
            Id="WPFToggleShowExt"; Name="Show File Extensions"; Category="Explorer"; Description="Always show file extensions in Explorer."; 
            Registry=@( @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name="HideFileExt"; Value="0"; OriginalValue="1"; Type="DWord" } );
            InvokeScript="Stop-Process -Name 'explorer' -Force"
        },
        @{ 
            Id="WPFTweaksEndTaskOnTaskbar"; Name="Enable End Task"; Category="Explorer"; Description="Adds 'End Task' to Taskbar right-click menu."; 
            Registry=@( @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings"; Name="TaskbarEndTask"; Value="1"; OriginalValue="0"; Type="DWord" } ) 
        },

        # --- INPUT ---
        @{ 
            Id="WPFToggleMouseAcceleration"; Name="Disable Mouse Accel"; Category="Input"; Description="Disables 'Enhance Pointer Precision'."; 
            Registry=@( 
                @{ Path="HKCU:\Control Panel\Mouse"; Name="MouseSpeed"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKCU:\Control Panel\Mouse"; Name="MouseThreshold1"; Value="0"; OriginalValue="6"; Type="DWord" },
                @{ Path="HKCU:\Control Panel\Mouse"; Name="MouseThreshold2"; Value="0"; OriginalValue="10"; Type="DWord" }
            ) 
        },
        @{ 
            Id="WPFToggleStickyKeys"; Name="Disable Sticky Keys"; Category="Input"; Description="Disables the Sticky Keys shortcut."; 
            Registry=@( @{ Path="HKCU:\Control Panel\Accessibility\StickyKeys"; Name="Flags"; Value="506"; OriginalValue="510"; Type="String" } ) 
        },

        # --- ADVANCED / DEBLOAT ---
        @{ 
            Id="WPFTweaksRemoveCopilot"; Name="Disable Copilot"; Category="Advanced Debloat"; Description="Disables the Microsoft Copilot AI integration (Win 11)."; 
            Registry=@( 
                @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name="ShowCopilotButton"; Value="0"; OriginalValue="1"; Type="DWord" },
                @{ Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot"; Name="TurnOffWindowsCopilot"; Value="1"; OriginalValue="0"; Type="DWord" },
                @{ Path="HKCU:\Software\Policies\Microsoft\Windows\WindowsCopilot"; Name="TurnOffWindowsCopilot"; Value="1"; OriginalValue="0"; Type="DWord" },
                @{ Path="HKLM:\SOFTWARE\Microsoft\Windows\Shell\Copilot"; Name="IsCopilotAvailable"; Value="0"; OriginalValue="<RemoveEntry>"; Type="DWord" },
                @{ Path="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsCopilot"; Name="AllowCopilotRuntime"; Value="0"; OriginalValue="<RemoveEntry>"; Type="DWord" }
            ) 
        },
        @{ 
            Id="WPFTweaksRecallOff"; Name="Disable Recall (AI)"; Category="Advanced Debloat"; Description="Disables Windows Recall AI features (Win 11 24H2+)."; 
            Registry=@( 
                @{ Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI"; Name="DisableAIDataAnalysis"; Value="1"; OriginalValue="<RemoveEntry>"; Type="DWord" },
                @{ Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI"; Name="AllowRecallEnablement"; Value="0"; OriginalValue="<RemoveEntry>"; Type="DWord" },
                @{ Path="HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy"; Name="VerifiedAndReputablePolicyState"; Value="0"; OriginalValue="<RemoveEntry>"; Type="DWord" }
            );
            InvokeScript="DISM /Online /Disable-Feature /FeatureName:Recall /Quiet /NoRestart";
            UndoScript="DISM /Online /Enable-Feature /FeatureName:Recall /Quiet /NoRestart"
        },
        @{ 
            Id="WPFTweaksDisableEdge"; Name="Disable Edge"; Category="Advanced Debloat"; Description="Prevents Microsoft Edge from running via the 'DisallowRun' policy."; 
            Registry=@( 
                @{ Path="HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"; Name="DisallowRun"; Value="1"; OriginalValue="0"; Type="DWord" },
                @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun"; Name="DisableEdge"; Value="msedge.exe"; OriginalValue="<RemoveEntry>"; Type="String" }
            )
        }
    )

    $sync.Software = @(
        # Browsers
        @{ Id="Google.Chrome"; Name="Google Chrome"; Category="Browsers"; Description="Web Browser" },
        @{ Id="Mozilla.Firefox"; Name="Firefox"; Category="Browsers"; Description="Web Browser" },
        @{ Id="Brave.Brave"; Name="Brave"; Category="Browsers"; Description="Privacy-focused Browser" },
        @{ Id="Microsoft.Edge"; Name="Edge"; Category="Browsers"; Description="Microsoft Edge" },
        
        # Utilities
        @{ Id="7zip.7zip"; Name="7-Zip"; Category="Utilities"; Description="File Archiver" },
        @{ Id="AdrienAllard.FileConverter"; Name="File Converter"; Category="Utilities"; Description="Context menu file converter." },
        @{ Id="Nilesoft.Shell"; Name="Nilesoft Shell"; Category="Utilities"; Description="Fixes Windows 11 context menu." },
        @{ Id="OCCT.OCCT"; Name="OCCT"; Category="Utilities"; Description="OverClock Checking Tool (Stress Test)." },
        @{ Id="RevoUninstaller.RevoUninstaller"; Name="Revo Uninstaller"; Category="Utilities"; Description="Advanced uninstaller." },
        @{ Id="RARLab.WinRAR"; Name="WinRAR"; Category="Utilities"; Description="Archive manager." },
        @{ Id="AntibodySoftware.WizTree"; Name="WizTree"; Category="Utilities"; Description="Disk Space Analyzer" },
        @{ Id="Microsoft.PowerToys"; Name="PowerToys"; Category="Utilities"; Description="System Utilities" },
        @{ Id="PuTTY.PuTTY"; Name="PuTTY"; Category="Utilities"; Description="SSH/Telnet Client" },
        @{ Id="WinSCP.WinSCP"; Name="WinSCP"; Category="Utilities"; Description="SFTP/FTP Client" },
        
        # Development
        @{ Id="Anysphere.Cursor"; Name="Cursor IDE"; Category="Development"; Description="AI-first code editor." },
        @{ Id="Microsoft.VisualStudioCode"; Name="VS Code"; Category="Development"; Description="Code Editor" },
        @{ Id="Git.Git"; Name="Git"; Category="Development"; Description="Version Control" },
        @{ Id="Microsoft.PowerShell"; Name="PowerShell 7"; Category="Development"; Description="Latest Cross-platform Shell" },
        @{ Id="Notepad++.Notepad++"; Name="Notepad++"; Category="Development"; Description="Text Editor" },
        @{ Id="Python.Python.3.12"; Name="Python 3"; Category="Development"; Description="Programming Language" },

        # Multimedia
        @{ Id="DuongDieuPhap.ImageGlass"; Name="ImageGlass"; Category="Multimedia"; Description="Lightweight image viewer." },
        @{ Id="Spotify.Spotify"; Name="Spotify"; Category="Multimedia"; Description="Music streaming." },
        @{ Id="VideoLAN.VLC"; Name="VLC Media Player"; Category="Multimedia"; Description="Plays anything" },
        @{ Id="OBSProject.OBSStudio"; Name="OBS Studio"; Category="Multimedia"; Description="Streaming Software" },
        @{ Id="Audacity.Audacity"; Name="Audacity"; Category="Multimedia"; Description="Audio Editor" },
        @{ Id="HandBrake.HandBrake"; Name="HandBrake"; Category="Multimedia"; Description="Video Transcoder" },
        @{ Id="GIMP.GIMP"; Name="GIMP"; Category="Multimedia"; Description="Image Editor" },

        # Communications
        @{ Id="Discord.Discord"; Name="Discord"; Category="Social"; Description="Chat & Streaming" },
        @{ Id="Telegram.TelegramDesktop"; Name="Telegram"; Category="Social"; Description="Messaging App" },
        @{ Id="Zoom.Zoom"; Name="Zoom"; Category="Social"; Description="Video Conferencing" },
        @{ Id="SlackTechnologies.Slack"; Name="Slack"; Category="Social"; Description="Team Collaboration" },

        # Gaming
        @{ Id="Nvidia.NvidiaApp"; Name="Nvidia App"; Category="Gaming"; Description="Modern replacement for GeForce Experience." },
        @{ Id="Valve.Steam"; Name="Steam"; Category="Gaming"; Description="Game Store" },
        @{ Id="EpicGames.EpicGamesLauncher"; Name="Epic Games"; Category="Gaming"; Description="Game Store" },

        # Internet
        @{ Id="qBittorrent.qBittorrent"; Name="qBittorrent"; Category="Internet"; Description="Free BitTorrent client." },

        # System Runtimes
        @{ Id="Microsoft.VCRedist.2015+.x64"; Name="Visual C++ Redists"; Category="System"; Description="Common runtimes for games/apps." }
    )

    # --- 4. RUNSPACE POOL SETUP ---
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, [int]$env:NUMBER_OF_PROCESSORS)
    $runspacePool.ApartmentState = "STA"
    $runspacePool.ThreadOptions = "ReuseThread"
    $runspacePool.Open()
    $sync.RunspacePool = $runspacePool

    # --- 5. WORKER SCRIPT BLOCK ---
    $workerScript = {
        param($SyncHash, $Action, $Payload)
        
        # --- HELPER FUNCTIONS ---
        function Set-RegValue {
            param ($Name, $Path, $Type, $Value)
            try {
                if(!(Test-Path 'HKU:\')) { New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS -ErrorAction SilentlyContinue | Out-Null }
                
                If (!(Test-Path $Path)) { New-Item -Path $Path -Force -ErrorAction Stop | Out-Null }

                if ($Value -ne "<RemoveEntry>") {
                    Set-ItemProperty -Path $Path -Name $Name -Type $Type -Value $Value -Force -ErrorAction Stop | Out-Null
                } else {
                    Remove-ItemProperty -Path $Path -Name $Name -Force -ErrorAction Stop | Out-Null
                }
            } catch {}
        }

        function Set-SvcConfig {
            param ($Name, $StartupType)
            try {
                $service = Get-Service -Name $Name -ErrorAction Stop
                $service | Set-Service -StartupType $StartupType -ErrorAction Stop
            } catch {}
        }

        function Set-TaskState {
            param ($Name, $State)
            try {
                if($State -eq "Disabled") { Disable-ScheduledTask -TaskName $Name -ErrorAction Stop }
                if($State -eq "Enabled") { Enable-ScheduledTask -TaskName $Name -ErrorAction Stop }
            } catch {}
        }

        function Invoke-CustomScript {
            param ($scriptblock)
            try { Invoke-Command $scriptblock -ErrorAction Stop } catch {}
        }

        # --- CORE TWEAK ENGINE ---
        function Process-TweakBatch {
            param($TweakIds, $Undo=$false) 
            if (-not $TweakIds) { return }
            
            # Map Undo/Apply fields
            if($Undo) {
                $Values = @{ Registry="OriginalValue"; ScheduledTask="OriginalState"; Service="OriginalType"; ScriptType="UndoScript" }
            } else {
                $Values = @{ Registry="Value"; ScheduledTask="State"; Service="StartupType"; ScriptType="InvokeScript" }
            }

            foreach ($tId in $TweakIds) {
                if ([string]::IsNullOrWhiteSpace($tId)) { continue }
                $tweak = $SyncHash.Tweaks | Where-Object { $_.Id -eq $tId }
                if (-not $tweak) { continue }
                
                $statusMsg = if ($Undo) { "UNDO" } else { "APPLY" }
                [Console]::WriteLine("[$statusMsg] $($tweak.Name)...")
                
                # Registry
                if($tweak.Registry) {
                    foreach($reg in $tweak.Registry) {
                        Set-RegValue -Name $reg.Name -Path $reg.Path -Type $reg.Type -Value $reg.$($Values.Registry)
                    }
                }
                # Services
                if($tweak.Service) {
                    foreach($svc in $tweak.Service) {
                        Set-SvcConfig -Name $svc.Name -StartupType $svc.$($Values.Service)
                    }
                }
                # Scheduled Tasks
                if($tweak.ScheduledTask) {
                    foreach($task in $tweak.ScheduledTask) {
                        Set-TaskState -Name $task.Name -State $task.$($Values.ScheduledTask)
                    }
                }
                # Scripts
                if($tweak.$($Values.ScriptType)) {
                    $sb = [scriptblock]::Create($tweak.$($Values.ScriptType))
                    Invoke-CustomScript -scriptblock $sb
                }
            }
        }

        function Manage-Package {
            param($Id, $Mode) 
            $cmd = if ($SyncHash.PackageManager -eq "winget") { "winget" } else { "choco" }
            $argList = @()
            if ($SyncHash.PackageManager -eq "winget") {
                switch ($Mode) {
                    "Install"   { $argList = @("install", "--id", $Id, "-e", "--silent", "--accept-source-agreements", "--accept-package-agreements") }
                    "Uninstall" { $argList = @("uninstall", "--id", $Id, "-e", "--silent") }
                    "Upgrade"   { $argList = @("upgrade", "--id", $Id, "-e", "--silent", "--accept-source-agreements", "--accept-package-agreements") }
                }
            } else {
                # Heuristic attempt to convert Winget ID to Choco ID (Grab last part)
                $Id = $Id.Split(".")[-1]; $argList = @($Mode.ToLower(), $Id, "-y")
            }
            Start-Process -FilePath $cmd -ArgumentList $argList -Wait -NoNewWindow
        }

        try {
            switch ($Action) {
                "LoadPower" { 
                    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
                    $raw = powercfg.exe -list
                    $plans = @()
                    foreach ($l in $raw) {
                        if ($l -match '([a-fA-F0-9]{8}-(?:[a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12})\s*\(([^)]+)\)') {
                            $plans += @{ Guid=$matches[1]; Name=$matches[2]; Status=$(if ($l -match '\*\s*$') { "Active" } else { "Inactive" }) }
                        }
                    }
                    $SyncHash.PowerPlans = $plans
                }
                "SetPower" { powercfg.exe -setactive $Payload; $SyncHash.ActionQueue += "LoadPower" }
                "AddUltimate" { powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61; $SyncHash.ActionQueue += "LoadPower" }
                "DeletePower" { powercfg -delete $Payload; $SyncHash.ActionQueue += "LoadPower" }

                "ApplyTweak" { Process-TweakBatch -TweakIds @($Payload) -Undo $false }
                "UndoTweak" { Process-TweakBatch -TweakIds @($Payload) -Undo $true }
                "ApplyTweakBatch" { Process-TweakBatch -TweakIds $Payload -Undo $false }
                "UndoTweakBatch" { Process-TweakBatch -TweakIds $Payload -Undo $true }
                
                "CreateRestorePoint" {
                    [Console]::WriteLine("[INFO] Creating System Restore Point...")
                    try {
                        Enable-ComputerRestore -Drive "$env:SystemDrive" -ErrorAction SilentlyContinue
                        Checkpoint-Computer -Description "Perdanga Manager Restore Point" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
                        [Console]::WriteLine("   -> Success.")
                    } catch { throw "Could not create restore point. $($_.Exception.Message)" }
                }
                
                "WU_GetStatus" {
                    $auPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
                    $noAutoUpdate = (Get-ItemProperty -Path $auPath -Name "NoAutoUpdate" -ErrorAction SilentlyContinue).NoAutoUpdate
                    if ($noAutoUpdate -eq 1) { $SyncHash.WUStatus = "Disabled" }
                    else {
                        $uxPath = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"
                        $deferral = (Get-ItemProperty -Path $uxPath -Name "DeferFeatureUpdatesPeriodInDays" -ErrorAction SilentlyContinue).DeferFeatureUpdatesPeriodInDays
                        if ($deferral -eq 365) { $SyncHash.WUStatus = "Security Only" }
                        else { $SyncHash.WUStatus = "Default (Enabled)" }
                    }
                }

                "WU_SetDefault" {
                    [Console]::WriteLine("[WU] Restoring Default Settings...")
                    if (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU")) { New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Force | Out-Null }
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "NoAutoUpdate" -Type DWord -Value 0 -Force
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "AUOptions" -Type DWord -Value 3 -Force
                    
                    if (!(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config")) { New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config" -Force | Out-Null }
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config" -Name "DODownloadMode" -Type DWord -Value 1 -Force

                    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\WaaSMedicSvc" -Name "Start" -Type DWord -Value 3 -ErrorAction SilentlyContinue
                    Remove-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\WaaSMedicSvc" -Name "FailureActions" -ErrorAction SilentlyContinue

                    $services = @(
                        @{Name = "BITS"; StartupType = "Manual"},
                        @{Name = "wuauserv"; StartupType = "Manual"},
                        @{Name = "UsoSvc"; StartupType = "Automatic"},
                        @{Name = "uhssvc"; StartupType = "Disabled"},
                        @{Name = "WaaSMedicSvc"; StartupType = "Manual"}
                    )
                    foreach ($service in $services) {
                        try {
                            Set-Service -Name $service.Name -StartupType $service.StartupType -ErrorAction SilentlyContinue
                            Start-Process -FilePath "sc.exe" -ArgumentList "failure `"$($service.Name)`" reset= 86400 actions= restart/60000/restart/60000/restart/60000" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
                            if ($service.StartupType -eq "Automatic") { Start-Service -Name $service.Name -ErrorAction SilentlyContinue }
                        } catch {}
                    }
                    
                    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" -Name "DeferFeatureUpdatesPeriodInDays" -ErrorAction SilentlyContinue
                    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" -Name "DeferQualityUpdatesPeriodInDays" -ErrorAction SilentlyContinue
                    Start-Process -FilePath "gpupdate" -ArgumentList "/force" -Wait -WindowStyle Hidden
                    [Console]::WriteLine("   -> Defaults Restored.")
                }

                "WU_SetSecurity" {
                    [Console]::WriteLine("[WU] Setting Security-Only Updates...")
                    if (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverSearching")) { New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverSearching" -Force | Out-Null }
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverSearching" -Name "DontPromptForWindowsUpdate" -Type DWord -Value 1 -Force
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverSearching" -Name "DontSearchWindowsUpdate" -Type DWord -Value 1 -Force
                    if (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU")) { New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Force | Out-Null }
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "NoAutoRebootWithLoggedOnUsers" -Type DWord -Value 1 -Force
                    if (!(Test-Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings")) { New-Item -Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" -Force | Out-Null }
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" -Name "BranchReadinessLevel" -Type DWord -Value 20 -Force
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" -Name "DeferFeatureUpdatesPeriodInDays" -Type DWord -Value 365 -Force
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" -Name "DeferQualityUpdatesPeriodInDays" -Type DWord -Value 4 -Force
                    [Console]::WriteLine("   -> Security Mode Applied.")
                }

                "WU_SetDisabled" {
                    [Console]::WriteLine("[WU] Disabling Windows Updates...")
                    if (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU")) { New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Force | Out-Null }
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "NoAutoUpdate" -Type DWord -Value 1 -Force
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "AUOptions" -Type DWord -Value 1 -Force
                    if (!(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config")) { New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config" -Force | Out-Null }
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config" -Name "DODownloadMode" -Type DWord -Value 0 -Force
                    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\WaaSMedicSvc" -Name "Start" -Type DWord -Value 4 -ErrorAction SilentlyContinue
                    $failureActions = [byte[]](0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x03,0x00,0x00,0x00,0x14,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xc0,0xd4,0x01,0x00,0x00,0x00,0x00,0x00,0xe0,0x93,0x04,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00)
                    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\WaaSMedicSvc" -Name "FailureActions" -Type Binary -Value $failureActions -ErrorAction SilentlyContinue
                    $services = @("BITS", "wuauserv", "UsoSvc", "uhssvc", "WaaSMedicSvc")
                    foreach ($service in $services) {
                        try {
                            Stop-Service -Name $service -Force -ErrorAction SilentlyContinue
                            Set-Service -Name $service -StartupType Disabled -ErrorAction SilentlyContinue
                            Start-Process -FilePath "sc.exe" -ArgumentList "failure `"$service`" reset= 0 actions= `"`"" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
                        } catch {}
                    }
                    $taskPaths = @('\Microsoft\Windows\InstallService\*', '\Microsoft\Windows\UpdateOrchestrator\*', '\Microsoft\Windows\UpdateAssistant\*', '\Microsoft\Windows\WaaSMedic\*', '\Microsoft\Windows\WindowsUpdate\*', '\Microsoft\WindowsUpdate\*')
                    foreach ($taskPath in $taskPaths) { try { Get-ScheduledTask -TaskPath $taskPath -ErrorAction SilentlyContinue | Disable-ScheduledTask -ErrorAction SilentlyContinue } catch {} }
                    [Console]::WriteLine("   -> Updates Disabled.")
                }

                "CheckTweaks" {
                    $appliedIds = @()
                    if (!(Test-Path 'HKU:\')) { New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS -ErrorAction SilentlyContinue | Out-Null }

                    # --- OPTIMIZATION: BATCH FETCHING ---
                    # Fetching Services and Tasks once is much faster than querying CIM/WMI/COM inside the loop.
                    
                    # 1. Cache Services
                    $serviceMap = @{}
                    try {
                        Get-Service -ErrorAction SilentlyContinue | ForEach-Object { $serviceMap[$_.Name] = $_ }
                    } catch {}

                    # 2. Cache Scheduled Tasks (The biggest performance bottleneck)
                    $taskMap = @{}
                    try {
                        Get-ScheduledTask -ErrorAction SilentlyContinue | ForEach-Object { 
                            # Key format: "\Path\Name" (Normalized)
                            $key = ($_.TaskPath + $_.TaskName).Replace("\\", "\").Trim()
                            $taskMap[$key] = $_ 
                        }
                    } catch {}

                    foreach ($tweak in $SyncHash.Tweaks) {
                        $isApplied = $true
                        
                        # Registry Check (Fast enough to do individually)
                        if ($tweak.Registry) {
                            foreach ($reg in $tweak.Registry) {
                                try {
                                    $regVal = Get-ItemProperty -Path $reg.Path -Name $reg.Name -ErrorAction Stop | Select-Object -ExpandProperty $reg.Name
                                    if ($regVal -ne $reg.Value) { $isApplied = $false; break }
                                } catch { 
                                    if ($reg.OriginalValue -eq "<RemoveEntry>") { if (Test-Path $reg.Path) { $isApplied = $false; break } } else { $isApplied = $false; break }
                                }
                            }
                        }
                        
                        # Service Check (Using Cache)
                        if ($isApplied -and $tweak.Service) {
                            foreach ($svc in $tweak.Service) {
                                $s = $serviceMap[$svc.Name]
                                if ($s) {
                                    if ($svc.StartupType -eq "Disabled" -and $s.StartType -ne "Disabled") { $isApplied = $false; break }
                                    if ($svc.StartupType -eq "Manual" -and $s.StartType -eq "Automatic") { $isApplied = $false; break }
                                }
                            }
                        }
                        
                        # Task Check (Using Cache)
                        if ($isApplied -and $tweak.ScheduledTask) {
                            foreach ($task in $tweak.ScheduledTask) {
                                # Construct lookup key based on definition
                                # Definition example: "Microsoft\Windows\..." 
                                # We need to ensure it matches the cache key format "\Microsoft\Windows\..."
                                $lookupKey = "\" + $task.Name.TrimStart("\")
                                $lookupKey = $lookupKey.Replace("\\", "\") # Normalize
                                
                                $t = $taskMap[$lookupKey]
                                if ($t -and $t.State -ne $task.State) { $isApplied = $false; break }
                            }
                        }
                        
                        if ($isApplied) { $appliedIds += $tweak.Id }
                    }
                    $SyncHash.AppliedTweaks = $appliedIds
                }

                "CheckInstalled" {
                    try {
                        $installedIds = @()
                        
                        # --- OPTIMIZED DETECTION LOGIC (FAST REGISTRY + FILE SCAN) ---
                        # Instead of running slow CLI commands (winget list/choco list), we scan the Registry and Files.
                        # This eliminates "infinite loading" and timeouts.
                        
                        # 1. GATHER ALL INSTALLED DISPLAY NAMES FROM REGISTRY
                        $regNames = New-Object System.Collections.Generic.HashSet[string]
                        $uninstallPaths = @(
                            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
                            "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
                        )
                        foreach ($path in $uninstallPaths) {
                            Get-ItemProperty $path -ErrorAction SilentlyContinue | ForEach-Object {
                                if ($_.DisplayName) { [void]$regNames.Add($_.DisplayName.Trim()) }
                            }
                        }
                        
                        # 2. GATHER CHOCO LIBS (IF APPLICABLE)
                        # Scanning folders in ProgramData is infinitely faster than 'choco list'
                        $chocoLibs = New-Object System.Collections.Generic.HashSet[string]
                        if ($SyncHash.PackageManager -eq "chocolatey") {
                            $chocoPath = "$env:ProgramData\chocolatey\lib"
                            if (Test-Path $chocoPath) {
                                Get-ChildItem -Path $chocoPath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                                    [void]$chocoLibs.Add($_.Name.ToLower())
                                }
                            }
                        }

                        # 3. MATCHING LOGIC
                        foreach ($sw in $SyncHash.Software) {
                            $isFound = $false
                            $swName = $sw.Name
                            $swId   = $sw.Id # e.g. "Google.Chrome"
                            
                            # Heuristic A: Registry Name Match (Contains)
                            # We iterate the hashset to find substring matches
                            foreach ($regName in $regNames) {
                                if ($regName -like "*$swName*") { $isFound = $true; break }
                                if ($regName -like "*$swId*") { $isFound = $true; break }
                                # Specific fix for VS Code (Reg: "Microsoft Visual Studio Code", Name: "VS Code")
                                if ($swName -eq "VS Code" -and $regName -like "*Visual Studio Code*") { $isFound = $true; break }
                            }

                            # Heuristic B: Chocolatey Folder Match (Exact or Suffix)
                            if (!$isFound -and $SyncHash.PackageManager -eq "chocolatey") {
                                $chocoId = $sw.Id.Split(".")[-1].ToLower() # "Google.Chrome" -> "chrome"
                                if ($chocoLibs.Contains($chocoId)) { $isFound = $true }
                                if ($chocoLibs.Contains($sw.Id.ToLower())) { $isFound = $true }
                            }

                            if ($isFound) { $installedIds += $sw.Id }
                        }
                        
                        $SyncHash.InstalledApps = $installedIds
                    } catch { [Console]::WriteLine("[ERROR] CheckInstalled failed: $_") }
                }

                "InstallApp" { Manage-Package -Id $Payload -Mode "Install"; $SyncHash.ActionQueue += "CheckInstalled" }
                "UninstallApp" { Manage-Package -Id $Payload -Mode "Uninstall"; $SyncHash.ActionQueue += "CheckInstalled" }
                "UpgradeApp" { Manage-Package -Id $Payload -Mode "Upgrade"; $SyncHash.ActionQueue += "CheckInstalled" }
                
                "ExportList" {
                    $path = "$([Environment]::GetFolderPath('Desktop'))\Perdanga_SoftwareList.txt"
                    $content = "--- Perdanga System Manager: Installed Software Report ---`r`n"
                    $content += "Date: $(Get-Date)`r`n`r`n"
                    $content += "--- Detected Database Matches ---`r`n"
                    if ($SyncHash.InstalledApps) { $content += ($SyncHash.InstalledApps -join "`r`n") }
                    else { $content += "None detected." }
                    
                    # Also include a raw dump of registry keys for complete backup utility
                    $content += "`r`n`r`n--- Raw System Application List (Registry Dump) ---`r`n"
                    $regKeys = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher
                    $regKeys += Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher
                    $regKeys += Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher
                    $sorted = $regKeys | Where-Object { $_.DisplayName } | Sort-Object DisplayName | Select-Object -Unique DisplayName, DisplayVersion, Publisher
                    foreach($item in $sorted) { $content += "$($item.DisplayName) ($($item.DisplayVersion))`r`n" }

                    $content | Out-File -FilePath $path -Encoding UTF8
                }

                "RunCommand" {
                    $cmd = $Payload; $args = $null
                    if ($Payload -match "^(\S+)\s+(.+)$") { $cmd = $matches[1]; $args = $matches[2] }
                    if ($args) { Start-Process -FilePath $cmd -ArgumentList $args -ErrorAction SilentlyContinue } 
                    else { Start-Process -FilePath $cmd -ErrorAction SilentlyContinue }
                }

                "SetDNS" {
                    $provider = $Payload
                    $dnsMap = @{
                        "Google" = @("8.8.8.8", "8.8.4.4", "2001:4860:4860::8888", "2001:4860:4860::8844")
                        "Cloudflare" = @("1.1.1.1", "1.0.0.1", "2606:4700:4700::1111", "2606:4700:4700::1001")
                    }
                    try {
                        $adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
                        [Console]::WriteLine("[DNS] Configuring $provider DNS on active adapters...")
                        foreach ($adapter in $adapters) {
                            if ($provider -eq "DHCP") {
                                Set-DnsClientServerAddress -InterfaceIndex $adapter.ifIndex -ResetServerAddresses -ErrorAction SilentlyContinue
                                [Console]::WriteLine("   -> DHCP set on $($adapter.Name)")
                            } 
                            elseif ($dnsMap.ContainsKey($provider)) {
                                $ips = $dnsMap[$provider]
                                Set-DnsClientServerAddress -InterfaceIndex $adapter.ifIndex -ServerAddresses ($ips[0], $ips[1]) -ErrorAction SilentlyContinue
                                try { Set-DnsClientServerAddress -InterfaceIndex $adapter.ifIndex -ServerAddresses ($ips[2], $ips[3]) -ErrorAction SilentlyContinue } catch {}
                                [Console]::WriteLine("   -> $provider DNS set on $($adapter.Name)")
                            }
                        }
                        Clear-DnsClientCache
                    } catch { [Console]::WriteLine("[ERROR] Failed to set DNS: $($_.Exception.Message)") }
                }

                "Maint_ResetUpdates" {
                    $services = @("BITS", "wuauserv", "appidsvc", "cryptsvc")
                    foreach ($svc in $services) { Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue }
                    if (Test-Path "$env:systemroot\SoftwareDistribution") { Rename-Item "$env:systemroot\SoftwareDistribution" "$env:systemroot\SoftwareDistribution.bak" -Force -ErrorAction SilentlyContinue }
                    if (Test-Path "$env:systemroot\System32\Catroot2") { Rename-Item "$env:systemroot\System32\Catroot2" "$env:systemroot\System32\Catroot2.bak" -Force -ErrorAction SilentlyContinue }
                    foreach ($svc in $services) { Start-Service -Name $svc -ErrorAction SilentlyContinue }
                    Start-Process "wuauclt" -ArgumentList "/resetauthorization", "/detectnow" -NoNewWindow -Wait
                }
                "Maint_SystemScan" {
                    Start-Process "chkdsk.exe" -ArgumentList "/scan" -NoNewWindow -Wait
                    Start-Process "sfc.exe" -ArgumentList "/scannow" -NoNewWindow -Wait
                    Start-Process "dism.exe" -ArgumentList "/Online /Cleanup-Image /RestoreHealth" -NoNewWindow -Wait
                    Start-Process "sfc.exe" -ArgumentList "/scannow" -NoNewWindow -Wait
                }
                "Maint_Cleanup" {
                    Start-Process cleanmgr.exe -ArgumentList "/d C: /VERYLOWDISK" -NoNewWindow -Wait
                    dism /online /Cleanup-Image /StartComponentCleanup /ResetBase
                }
                "Maint_Network" { 
                    netsh winsock reset; netsh int ip reset; netsh winhttp reset proxy; ipconfig /flushdns 
                }
                "Maint_IconCache" {
                    Stop-Process -Name "explorer" -Force -ErrorAction SilentlyContinue
                    Remove-Item "$env:localappdata\IconCache.db" -Force -ErrorAction SilentlyContinue
                    Remove-Item "$env:localappdata\Microsoft\Windows\Explorer\iconcache*" -Force -ErrorAction SilentlyContinue
                    Start-Process "explorer.exe"
                }
                "Maint_ReinstallPackageManager" {
                    if ($SyncHash.PackageManager -eq "winget") { Start-Process -FilePath "winget" -ArgumentList "install -e --accept-source-agreements --accept-package-agreements Microsoft.AppInstaller" -Wait -NoNewWindow } 
                    elseif ($SyncHash.PackageManager -eq "chocolatey") { Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1')) }
                }
                "Maint_RemoveOneDrive" {
                    # Robust Cleanup Logic (Based on External Source)
                    $OneDrivePath = $env:OneDrive
                    [Console]::WriteLine("[INFO] Removing OneDrive (Deep Clean)...")
                    
                    # Kill Processes
                    Stop-Process -Name "*OneDrive*" -Force -ErrorAction SilentlyContinue
                    
                    # Uninstall Command (Check)
                    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\OneDriveSetup.exe"
                    $msStorePath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\Applications\*OneDrive*"

                    if (Test-Path $regPath) {
                        $uninstallStr = (Get-ItemProperty $regPath).UninstallString
                        $exe = $uninstallStr.Split(" ")[0]
                        $args = $uninstallStr.Substring($exe.Length) + " /silent"
                        Start-Process -FilePath $exe -ArgumentList $args -Wait -NoNewWindow
                    } elseif (Test-Path $msStorePath) {
                        Start-Process -FilePath winget -ArgumentList "uninstall -e --purge --accept-source-agreements Microsoft.OneDrive" -NoNewWindow -Wait
                    } else {
                        if (Test-Path "$env:SystemRoot\System32\OneDriveSetup.exe") { Start-Process "$env:SystemRoot\System32\OneDriveSetup.exe" -ArgumentList "/uninstall" -Wait -NoNewWindow } 
                        elseif (Test-Path "$env:SystemRoot\SysWOW64\OneDriveSetup.exe") { Start-Process "$env:SystemRoot\SysWOW64\OneDriveSetup.exe" -ArgumentList "/uninstall" -Wait -NoNewWindow }
                    }
                    
                    # Cleanup Directories
                    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "$env:localappdata\Microsoft\OneDrive"
                    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "$env:programdata\Microsoft OneDrive"
                    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "$env:OneDrive"
                    Remove-Item -Path "HKCU:\Software\Microsoft\OneDrive" -Recurse -Force -ErrorAction SilentlyContinue
                    
                    # Cleanup Explorer Namespace
                    Set-RegValue -Name "System.IsPinnedToNameSpaceTree" -Path "HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" -Value 0 -Type "DWord"
                    Set-RegValue -Name "System.IsPinnedToNameSpaceTree" -Path "HKCR:\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" -Value 0 -Type "DWord"
                    
                    # Remove Run hooks
                    Set-RegValue -Name "OneDrive" -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Value "<RemoveEntry>" -Type "String"
                    
                    # Remove Environment Variable
                    [Environment]::SetEnvironmentVariable("OneDrive", $null, "User")
                    
                    taskkill.exe /F /IM "explorer.exe"
                    Start-Process "explorer.exe"
                    [Console]::WriteLine("   -> OneDrive Removed.")
                }
                "Maint_OOSU10" {
                    $url = "https://dl5.oo-software.com/files/ooshutup10/OOSU10.exe"; $dest = "$env:TEMP\OOSU10.exe"
                    Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing; Start-Process -FilePath $dest -Wait
                }
                "Maint_Autoruns" {
                    $zip = "$env:TEMP\Autoruns.zip"; $dest = "$env:TEMP\Autoruns"
                    Invoke-WebRequest "https://download.sysinternals.com/files/Autoruns.zip" -OutFile $zip -UseBasicParsing
                    Expand-Archive -Path $zip -DestinationPath $dest -Force; Start-Process -FilePath "$dest\Autoruns64.exe" -Wait
                }
                "Maint_EnableSSH" {
                     $cap = Get-WindowsCapability -Online | Where-Object { $_.Name -like "OpenSSH.Server*" }
                     if ($cap.State -ne "Installed") { Add-WindowsCapability -Online -Name $cap.Name }
                     Start-Service sshd; Set-Service sshd -StartupType Automatic
                     Start-Service ssh-agent; Set-Service ssh-agent -StartupType Automatic
                     if (!(Get-NetFirewallRule -Name "OpenSSH-Server-In-TCP" -ErrorAction SilentlyContinue)) { New-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -DisplayName 'OpenSSH Server (sshd)' -Enabled True -Direction Inbound -Protocol TCP -Action Allow -LocalPort 22 }
                }
                "Maint_AdobeClean" {
                    $url = "https://swupmf.adobe.com/webfeed/CleanerTool/win/AdobeCreativeCloudCleanerTool.exe"; $dest = "$env:TEMP\AdobeCreativeCloudCleanerTool.exe"
                    Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
                    Start-Process -FilePath $dest -Wait
                }
                "RestartExplorer" {
                    Stop-Process -Name "explorer" -Force -ErrorAction SilentlyContinue
                    [Console]::WriteLine("[INFO] Explorer restarted.")
                }
            }
            if ($Action -notmatch "WU_GetStatus") { $SyncHash.Result = "Success" }
        }
        catch {
            $SyncHash.Result = "Error"
            $SyncHash.ErrorMessage = $_.Exception.Message
        }
    }

    # --- 6. GUI SETUP ---
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Perdanga System Manager"
    $form.Size = New-Object System.Drawing.Size(1280, 750)
    $form.StartPosition = "CenterScreen"
    $form.MinimumSize = New-Object System.Drawing.Size(950, 600)
    $sync.Form = $form

    # --- Theme Colors (Minimalist Modern) ---
    $colDarkBg      = [System.Drawing.Color]::FromArgb(30, 30, 35)
    $colSideBar     = [System.Drawing.Color]::FromArgb(20, 20, 25)
    $colPanelBg     = [System.Drawing.Color]::FromArgb(40, 40, 45)
    $colBtnSurface  = [System.Drawing.Color]::FromArgb(50, 50, 55) # New Neutral Button Surface
    
    # Pastel / Matte Palette for Minimalist Look
    $colAccentBlue  = [System.Drawing.Color]::FromArgb(100, 140, 210)  # Soft Slate Blue
    $colAccentGreen = [System.Drawing.Color]::FromArgb(100, 180, 130)  # Soft Sage Green
    $colAccentPurple= [System.Drawing.Color]::FromArgb(170, 110, 200)  # Muted Lavender
    $colAccentRed   = [System.Drawing.Color]::FromArgb(210, 100, 100)  # Soft Coral/Red
    $colAccentOrange= [System.Drawing.Color]::FromArgb(220, 150, 80)   # Muted Amber
    $colAccentCyan  = [System.Drawing.Color]::FromArgb(80, 200, 210)   # Soft Teal
    $colAccentGray  = [System.Drawing.Color]::FromArgb(140, 140, 145)  # Medium Gray

    $colTextWhite   = [System.Drawing.Color]::White
    $colTextGray    = [System.Drawing.Color]::Gainsboro
    $colBorder      = [System.Drawing.Color]::FromArgb(60, 60, 65)

    $form.BackColor = $colDarkBg
    $fontStd  = New-Object System.Drawing.Font("Segoe UI", 10)
    $fontBold = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $fontHead = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $fontMenu = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)

    # --- TIMER ---
    $uiTimer = New-Object System.Windows.Forms.Timer
    $uiTimer.Interval = 100
    $uiTimer.Tag = $null
    
    # --- CONTROLS ---
    $splitMain = New-Object System.Windows.Forms.SplitContainer; $splitMain.Dock = "Fill"; $splitMain.FixedPanel = "Panel1"; $splitMain.SplitterDistance = 260; $splitMain.Panel1.BackColor = $colSideBar; $splitMain.Panel2.BackColor = $colDarkBg; $form.Controls.Add($splitMain)

    # Sidebar
    $lblMenuTitle = New-Object System.Windows.Forms.Label; $lblMenuTitle.Text = "MENU"; $lblMenuTitle.Dock = "Top"; $lblMenuTitle.Height = 70; $lblMenuTitle.Font = $fontMenu; $lblMenuTitle.ForeColor = [System.Drawing.Color]::Gray; $lblMenuTitle.TextAlign = "MiddleCenter"; $splitMain.Panel1.Controls.Add($lblMenuTitle)
    $flowMenu = New-Object System.Windows.Forms.FlowLayoutPanel; $flowMenu.Dock = "Fill"; $flowMenu.FlowDirection = "TopDown"; $flowMenu.WrapContents = $false; $flowMenu.Padding = New-Object System.Windows.Forms.Padding(0, 10, 0, 0); $splitMain.Panel1.Controls.Add($flowMenu); $flowMenu.BringToFront()
    
    $sync.ViewColors = @{ "Tweaks"=$colAccentPurple; "Installs"=$colAccentGreen; "Maintenance"=$colAccentOrange; "WindowsUpdate"=$colAccentBlue; "Power"=$colAccentRed; "Tools"=$colAccentGray }

    function New-MenuButton($Text, $Tag, $ToolTipText) {
        $BaseColor = $sync.ViewColors[$Tag]
        $btn = New-Object System.Windows.Forms.Button; $btn.Text = "  $Text"; $btn.Tag = $Tag; $btn.Width = 260; $btn.Height = 55; $btn.FlatStyle = "Flat"; $btn.FlatAppearance.BorderSize = 0; $btn.Font = $fontStd; $btn.ForeColor = $colTextGray; $btn.TextAlign = "MiddleLeft"; $btn.Cursor = [System.Windows.Forms.Cursors]::Hand; $btn.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 2); $sync.MenuButtons[$Tag] = $btn
        $tooltip = New-Object System.Windows.Forms.ToolTip
        $tooltip.SetToolTip($btn, $ToolTipText)
        $btn.Add_Click({ if ($script:currentView -ne $this.Tag) { Set-View $this.Tag } })
        return $btn
    }

    $flowMenu.Controls.Add((New-MenuButton "System Tweaks" "Tweaks" "Apply privacy and performance tweaks"))
    $flowMenu.Controls.Add((New-MenuButton "Install Apps" "Installs" "Batch install software via Winget or Chocolatey"))
    $flowMenu.Controls.Add((New-MenuButton "Maintenance" "Maintenance" "Repair and cleanup tools"))
    $flowMenu.Controls.Add((New-MenuButton "Windows Update" "WindowsUpdate" "Manage Windows Update policies"))
    $flowMenu.Controls.Add((New-MenuButton "Power Config" "Power" "Manage Power Plans"))
    # Removed Windows Features as requested
    $flowMenu.Controls.Add((New-MenuButton "Quick Tools" "Tools" "Shortcuts to common system panels"))

    # --- Status Strip ---
    $statusStrip = New-Object System.Windows.Forms.StatusStrip; $statusStrip.BackColor = $colSideBar; $statusStrip.ForeColor = $colTextWhite
    $lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel; $lblStatus.Text = "Ready"; $lblStatus.Spring = $true; $lblStatus.TextAlign = "MiddleLeft"
    
    # ProgressBar REMOVED here
    
    $statusStrip.Items.AddRange(@($lblStatus)); $form.Controls.Add($statusStrip)

    # --- PANELS FOR CONTENT ---
    
    # Header Panel
    $pnlHeader = New-Object System.Windows.Forms.Panel; $pnlHeader.Dock = "Top"; $pnlHeader.Height = 60; $pnlHeader.BackColor = $colDarkBg; $pnlHeader.Padding = New-Object System.Windows.Forms.Padding(20, 15, 0, 0)
    $lblHeaderTitle = New-Object System.Windows.Forms.Label; $lblHeaderTitle.Text = "Dashboard"; $lblHeaderTitle.Font = $fontHead; $lblHeaderTitle.ForeColor = $colTextWhite; $lblHeaderTitle.AutoSize = $true; $pnlHeader.Controls.Add($lblHeaderTitle)
    $splitMain.Panel2.Controls.Add($pnlHeader)

    $pnlActions = New-Object System.Windows.Forms.FlowLayoutPanel; $pnlActions.Dock = "Bottom"; $pnlActions.Height = 60; $pnlActions.BackColor = $colDarkBg; $pnlActions.FlowDirection = "LeftToRight"; $splitMain.Panel2.Controls.Add($pnlActions)
    
    $flowMaint = New-Object System.Windows.Forms.FlowLayoutPanel; $flowMaint.Dock = "Fill"; $flowMaint.AutoScroll = $true; $flowMaint.Visible = $false; $flowMaint.Padding = New-Object System.Windows.Forms.Padding(20)
    $flowTools = New-Object System.Windows.Forms.FlowLayoutPanel; $flowTools.Dock = "Fill"; $flowTools.AutoScroll = $true; $flowTools.Visible = $false; $flowTools.Padding = New-Object System.Windows.Forms.Padding(20)
    $pnlWU = New-Object System.Windows.Forms.Panel; $pnlWU.Dock = "Fill"; $pnlWU.Visible = $false
    
    # Windows Update UI
    $grpStatus = New-Object System.Windows.Forms.GroupBox; $grpStatus.Text = "Update Status"; $grpStatus.ForeColor = $colTextWhite; $grpStatus.Dock = "Top"; $grpStatus.Height = 100; $grpStatus.Font = $fontBold; $pnlWU.Controls.Add($grpStatus)
    $lblCurMode = New-Object System.Windows.Forms.Label; $lblCurMode.Text = "Mode: Checking..."; $lblCurMode.Location = "20,30"; $lblCurMode.AutoSize = $true; $lblCurMode.Font = $fontStd; $grpStatus.Controls.Add($lblCurMode)
    $grpOpts = New-Object System.Windows.Forms.GroupBox; $grpOpts.Text = "Configuration"; $grpOpts.ForeColor = $colTextWhite; $grpOpts.Dock = "Fill"; $grpOpts.Font = $fontBold; $pnlWU.Controls.Add($grpOpts); $grpOpts.BringToFront()
    $rbDef = New-Object System.Windows.Forms.RadioButton; $rbDef.Text = "Default Settings (Enable All)"; $rbDef.Location = "20,40"; $rbDef.AutoSize = $true; $rbDef.ForeColor = $colTextWhite; $grpOpts.Controls.Add($rbDef)
    $rbSec = New-Object System.Windows.Forms.RadioButton; $rbSec.Text = "Security Only (Defer Features 1 Year)"; $rbSec.Location = "20,80"; $rbSec.AutoSize = $true; $rbSec.ForeColor = $colTextWhite; $grpOpts.Controls.Add($rbSec)
    $rbDis = New-Object System.Windows.Forms.RadioButton; $rbDis.Text = "Disable Updates (Not Recommended)"; $rbDis.Location = "20,120"; $rbDis.AutoSize = $true; $rbDis.ForeColor = [System.Drawing.Color]::Salmon; $grpOpts.Controls.Add($rbDis)
    
    # Install Apps UI
    $pnlPackageManager = New-Object System.Windows.Forms.Panel; $pnlPackageManager.Height = 40; $pnlPackageManager.Dock = "Top"; $pnlPackageManager.BackColor = $colPanelBg; $pnlPackageManager.Padding = New-Object System.Windows.Forms.Padding(10, 5, 10, 5); $pnlPackageManager.Visible = $false
    $lblPackageManager = New-Object System.Windows.Forms.Label; $lblPackageManager.Text = "Package Manager:"; $lblPackageManager.AutoSize = $true; $lblPackageManager.ForeColor = $colTextWhite; $lblPackageManager.Location = "10, 10"; $pnlPackageManager.Controls.Add($lblPackageManager)
    $rbWinget = New-Object System.Windows.Forms.RadioButton; $rbWinget.Text = "Winget"; $rbWinget.AutoSize = $true; $rbWinget.Checked = ($sync.PackageManager -eq "winget"); $rbWinget.ForeColor = $colTextWhite; $rbWinget.Location = "120, 10"
    $rbWinget.Add_CheckedChanged({ if ($this.Checked) { $sync.PackageManager = "winget"; Start-BackgroundTask "CheckInstalled" $null } }); $pnlPackageManager.Controls.Add($rbWinget)
    $rbChocolatey = New-Object System.Windows.Forms.RadioButton; $rbChocolatey.Text = "Chocolatey"; $rbChocolatey.AutoSize = $true; $rbChocolatey.Checked = ($sync.PackageManager -eq "chocolatey"); $rbChocolatey.ForeColor = $colTextWhite; $rbChocolatey.Location = "200, 10"
    $rbChocolatey.Add_CheckedChanged({ if ($this.Checked) { $sync.PackageManager = "chocolatey"; Start-BackgroundTask "CheckInstalled" $null } }); $pnlPackageManager.Controls.Add($rbChocolatey)
    
    $pnlSearch = New-Object System.Windows.Forms.Panel; $pnlSearch.Height = 50; $pnlSearch.Dock = "Top"; $pnlSearch.BackColor = $colPanelBg; $pnlSearch.Padding = New-Object System.Windows.Forms.Padding(20, 10, 20, 10); $pnlSearch.Visible = $false
    $txtSearch = New-Object System.Windows.Forms.TextBox; $txtSearch.Location = "20, 15"; $txtSearch.Size = "300, 25"; $txtSearch.Font = $fontStd; $txtSearch.ForeColor = $colTextGray; $txtSearch.BackColor = $colDarkBg; $txtSearch.BorderStyle = "FixedSingle"; $txtSearch.Text = "Search software..."
    $txtSearch.Add_Enter({ if ($this.Text -eq "Search software...") { $this.Text = ""; $this.ForeColor = $colTextWhite } })
    $txtSearch.Add_Leave({ if ([string]::IsNullOrWhiteSpace($this.Text)) { $this.Text = "Search software..."; $this.ForeColor = $colTextGray } })
    $txtSearch.Add_TextChanged({ if ($this.Text -ne "Search software...") { $sync.SoftwareSearchText = $this.Text } else { $sync.SoftwareSearchText = "" }; Update-ListView "Installs" })
    $pnlSearch.Controls.Add($txtSearch)
    
    function Start-BackgroundTask {
        param($Action, $Payload)
        if ($sync.IsBusy) { Write-Host "[WARN] System is busy. Ignoring Action: $Action" -ForegroundColor Yellow; return }
        Write-Host "[TASK] Starting: $Action" -NoNewline -ForegroundColor Cyan; if ($Payload) { Write-Host " | Target: $Payload" -ForegroundColor Gray } else { Write-Host "" }
        $sync.IsBusy = $true; $lblStatus.Text = "Processing $Action..."; $pnlActions.Enabled = $false
        # ProgressBar logic REMOVED here
        $ps = [powershell]::Create(); $ps.RunspacePool = $sync.RunspacePool; [void]$ps.AddScript($workerScript).AddArgument($sync).AddArgument($Action).AddArgument($Payload)
        $handle = $ps.BeginInvoke(); $uiTimer.Tag = @{ Handle=$handle; Shell=$ps; Action=$Action }; $uiTimer.Start()
    }

    $uiTimer.Add_Tick({
        $state = $uiTimer.Tag
        if ($state -and $state.Handle.IsCompleted) {
            $uiTimer.Stop(); $ps = $state.Shell
            try { $ps.EndInvoke($state.Handle); $ps.Dispose() } catch { Write-Host "[ERROR] Exception: $($_.Exception.Message)" -ForegroundColor Red }
            if ($sync.Result -eq "Error") { Write-Host "[FAIL] Task Failed: $($state.Action)" -ForegroundColor Red; [System.Windows.Forms.MessageBox]::Show($sync.ErrorMessage, "Error", "OK", "Error") } 
            else { 
                Write-Host "[DONE] Task Completed: $($state.Action)" -ForegroundColor Green 
                if ($state.Action -eq "ExportList") { [System.Windows.Forms.MessageBox]::Show("Software list exported to Desktop.", "Export Success", "OK", "Information") }
            }

            # Queue Processing
            foreach ($qItem in $sync.ActionQueue) {
                if ($qItem -eq "CheckInstalled" -and $script:currentView -eq "Installs") { Update-ListView "Installs" }
                if ($qItem -eq "CheckTweaks" -and $script:currentView -eq "Tweaks") { Update-ListView "Tweaks" }
                if ($qItem -eq "LoadPower") { Update-ListView "Power" }
            }
            $sync.ActionQueue = @()

            # Single Action Updates
            if ($state.Action -eq "CheckTweaks" -and $script:currentView -eq "Tweaks") { Update-ListView "Tweaks" }
            if ($state.Action -eq "CheckInstalled" -and $script:currentView -eq "Installs") { Update-ListView "Installs" }
            if ($state.Action -eq "LoadPower" -and $script:currentView -eq "Power") { Update-ListView "Power" }
            if ($state.Action -eq "WU_GetStatus") { $lblCurMode.Text = "Current Status: " + $sync.WUStatus }
            
            $sync.IsBusy = $false; $lblStatus.Text = "Ready"; $pnlActions.Enabled = $true
            # ProgressBar logic REMOVED here
        }
    })

    function New-DataList($ColsOrdered, $CheckBoxes) {
        $lv = New-Object System.Windows.Forms.ListView
        $lv.Dock = "Fill"; $lv.View = "Details"; $lv.FullRowSelect = $true; $lv.GridLines = $false; $lv.HeaderStyle = "Nonclickable"; $lv.CheckBoxes = $CheckBoxes; 
        # Note: AllowColumnReorder affects user dragging columns to new positions, not resizing.
        $lv.AllowColumnReorder = $false
        $lv.BackColor = $colPanelBg; $lv.ForeColor = $colTextGray; $lv.BorderStyle = "None"; $lv.Font = $fontStd; $lv.OwnerDraw = $true; $lv.Visible = $false
        
        # --- STORE WIDTHS IN TAG FOR LOCKING ---
        $lv.Tag = @{ 
            HeaderBrush = [System.Drawing.SolidBrush]::new($colDarkBg); 
            TextBrush = [System.Drawing.SolidBrush]::new($colTextWhite); 
            SelBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(60,60,70)); 
            BgBrush = [System.Drawing.SolidBrush]::new($colPanelBg); 
            BorderPen = [System.Drawing.Pen]::new($colBorder, 1);
            ColumnWidths = @{} # Store authorized widths
        }

        # Add columns and populate widths
        $i = 0
        foreach ($k in $ColsOrdered.Keys) { 
            $w = $ColsOrdered[$k]
            $lv.Columns.Add($k, $w) | Out-Null 
            $lv.Tag.ColumnWidths[$i] = $w
            $i++
        }

        # --- STRICT LOCKING & VISUAL FIX ---
        
        # 1. Prevent Resizing (Lock)
        $lv.Add_ColumnWidthChanging({ 
            param($sender, $e)
            $e.Cancel = $true
            # Force back to stored width immediately to prevent visual jitter
            $e.NewWidth = $sender.Tag.ColumnWidths[$e.ColumnIndex]
        })

        # 2. Fill Empty Space (Fix White Header Bug)
        $lv.Add_Resize({ 
            param($sender, $e)
            if ($sender.Columns.Count -gt 0) {
                $totalW = 0
                # Sum all columns except the last one
                for($j=0; $j -lt $sender.Columns.Count - 1; $j++) {
                    $totalW += $sender.Columns[$j].Width
                }
                # Calculate remaining width
                $newLastW = $sender.ClientRectangle.Width - $totalW
                # Apply to last column if valid
                if ($newLastW -gt 50) {
                    $lastIdx = $sender.Columns.Count - 1
                    $sender.Columns[$lastIdx].Width = $newLastW
                    # IMPORTANT: Update the "locked" width so the locking event accepts the new size
                    $sender.Tag.ColumnWidths[$lastIdx] = $newLastW
                }
            }
        })

        $lv.Add_DrawColumnHeader({ param($s,$e) $e.Graphics.FillRectangle($s.Tag.HeaderBrush, $e.Bounds); $sf = [System.Drawing.StringFormat]::new(); $sf.LineAlignment = "Center"; $rectF = [System.Drawing.RectangleF]::new($e.Bounds.X, $e.Bounds.Y, $e.Bounds.Width, $e.Bounds.Height); $e.Graphics.DrawString($e.Header.Text, $fontBold, $s.Tag.TextBrush, $rectF, $sf) })
        $lv.Add_DrawItem({ param($s,$e) $e.DrawDefault = $true })
        $lv.Add_DrawSubItem({ 
            param($s,$e) 
            $color = $colTextGray
            if ($e.ColumnIndex -eq 0) { if (($sync.InstalledApps -contains $e.Item.Tag) -or ($sync.AppliedTweaks -contains $e.Item.Tag)) { $color = [System.Drawing.Color]::LightGreen } else { $color = $colTextWhite } }
            elseif ($e.ColumnIndex -ge 1) { if ($e.SubItem.Text -match "Active" -or $e.SubItem.Text -match "Enabled" -or $e.SubItem.Text -eq "Installed") { $color = [System.Drawing.Color]::LightGreen } elseif ($e.SubItem.Text -match "Inactive" -or $e.SubItem.Text -match "Disabled" -or $e.SubItem.Text -eq "Not Installed") { $color = [System.Drawing.Color]::Gray } } 
            if ($e.Item.Selected) { $e.Graphics.FillRectangle($s.Tag.SelBrush, $e.Bounds); $color = $colTextWhite } else { $e.Graphics.FillRectangle($s.Tag.BgBrush, $e.Bounds) }
            $sf = [System.Drawing.StringFormat]::new(); $sf.LineAlignment = "Center"; $sf.Trimming = "EllipsisCharacter"; $x = $e.Bounds.X; if ($e.ColumnIndex -eq 0 -and $s.CheckBoxes) { $x += 22 } else { $x += 5 }
            $brush = [System.Drawing.SolidBrush]::new($color); $rectF = [System.Drawing.RectangleF]::new($x, $e.Bounds.Y, $e.Bounds.Width, $e.Bounds.Height); $e.Graphics.DrawString($e.SubItem.Text, $fontStd, $brush, $rectF, $sf); $brush.Dispose(); $e.Graphics.DrawLine($s.Tag.BorderPen, $e.Bounds.Left, $e.Bounds.Bottom - 1, $e.Bounds.Right, $e.Bounds.Bottom - 1)
        })
        return $lv
    }

    $lvPower    = New-DataList ([ordered]@{ "Plan Name"=500; "Status"=120; "GUID"=200 }) $false
    $lvTweaks   = New-DataList ([ordered]@{ "Tweak Name"=300; "Status"=100; "Category"=150; "Description"=500 }) $true
    $lvInstalls = New-DataList ([ordered]@{ "Software Name"=300; "Status"=120; "Category"=150; "Description"=500 }) $true 

    $pnlListContainer = New-Object System.Windows.Forms.Panel; $pnlListContainer.Dock = "Fill"; $pnlListContainer.Padding = New-Object System.Windows.Forms.Padding(20, 20, 20, 0); $pnlListContainer.BackColor = $colDarkBg
    $pnlListContainer.Controls.Add($lvPower); $pnlListContainer.Controls.Add($lvTweaks); $pnlListContainer.Controls.Add($lvInstalls); $pnlListContainer.Controls.Add($flowMaint); $pnlListContainer.Controls.Add($flowTools); $pnlListContainer.Controls.Add($pnlWU); $pnlListContainer.Controls.Add($pnlPackageManager); $pnlListContainer.Controls.Add($pnlSearch)
    $splitMain.Panel2.Controls.Add($pnlListContainer); $pnlListContainer.BringToFront()

    function Update-ListView($Type) {
        if ($Type -eq "Power") {
            $lvPower.BeginUpdate(); $lvPower.Items.Clear()
            foreach ($p in $sync.PowerPlans) { $item = $lvPower.Items.Add($p.Name); $item.SubItems.Add($p.Status) | Out-Null; $item.SubItems.Add($p.Guid) | Out-Null; $item.Tag = $p.Guid; if ($p.Status -eq "Active") { $item.SubItems[1].ForeColor = $colAccentGreen } }
            $lvPower.EndUpdate()
        } elseif ($Type -eq "Tweaks") {
            $lvTweaks.BeginUpdate(); $lvTweaks.Items.Clear()
            foreach ($t in $sync.Tweaks) { $item = $lvTweaks.Items.Add($t.Name); if ($sync.AppliedTweaks -contains $t.Id) { $item.SubItems.Add("Active") | Out-Null; $item.ForeColor = $colAccentGreen } else { $item.SubItems.Add("Inactive") | Out-Null }; $item.SubItems.Add($t.Category) | Out-Null; $item.SubItems.Add($t.Description) | Out-Null; $item.Tag = $t.Id }
            $lvTweaks.EndUpdate()
        } elseif ($Type -eq "Installs") {
            $lvInstalls.BeginUpdate(); $lvInstalls.Items.Clear()
            $filteredSoftware = $sync.Software
            if ($sync.SoftwareSearchText -and $sync.SoftwareSearchText.Trim() -ne "") { $searchText = $sync.SoftwareSearchText.ToLower(); $filteredSoftware = $sync.Software | Where-Object { $_.Name.ToLower().Contains($searchText) -or $_.Category.ToLower().Contains($searchText) } }
            foreach ($s in $filteredSoftware) { 
                $item = $lvInstalls.Items.Add($s.Name)
                if ($sync.InstalledApps -contains $s.Id) { $item.SubItems.Add("Installed") | Out-Null; $item.ForeColor = $colAccentGreen } 
                else { $item.SubItems.Add("Not Installed") | Out-Null }
                $item.SubItems.Add($s.Category) | Out-Null; $item.SubItems.Add($s.Description) | Out-Null; $item.Tag = $s.Id 
            }
            $lvInstalls.EndUpdate()
        }
    }

    # Minimalist Button: Neutral BG, Colored Text, Colored Border
    function New-ActionButton($Text, $Color, $ActionScript) { 
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $Text
        $btn.BackColor = $colBtnSurface
        $btn.ForeColor = $Color
        $btn.FlatStyle = "Flat"
        $btn.FlatAppearance.BorderSize = 1
        $btn.FlatAppearance.BorderColor = $Color
        $btn.Size = New-Object System.Drawing.Size(140, 35)
        $btn.Margin = New-Object System.Windows.Forms.Padding(0, 10, 10, 0)
        $btn.Add_Click($ActionScript)
        return $btn 
    }
    
    function New-ActionButtonWithTooltip($Text, $Color, $ToolTipText, $ActionScript) { 
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $Text
        $btn.BackColor = $colBtnSurface
        $btn.ForeColor = $Color
        $btn.FlatStyle = "Flat"
        $btn.FlatAppearance.BorderSize = 1
        $btn.FlatAppearance.BorderColor = $Color
        $btn.Size = New-Object System.Drawing.Size(140, 35)
        $btn.Margin = New-Object System.Windows.Forms.Padding(0, 10, 10, 0)
        $tooltip = New-Object System.Windows.Forms.ToolTip
        $tooltip.SetToolTip($btn, $ToolTipText)
        $btn.Add_Click($ActionScript)
        return $btn 
    }

    function New-MaintCard($Title, $Desc, $ActionId, $Color) {
        $p = New-Object System.Windows.Forms.Panel; $p.Size = New-Object System.Drawing.Size(300, 150); $p.BackColor = $colPanelBg; $p.Margin = New-Object System.Windows.Forms.Padding(0, 0, 20, 20)
        $l = New-Object System.Windows.Forms.Label; $l.Text = $Title; $l.Font = $fontBold; $l.ForeColor = $colTextWhite; $l.Location = "10,10"; $l.AutoSize = $true; $p.Controls.Add($l)
        $d = New-Object System.Windows.Forms.Label; $d.Text = $Desc; $d.ForeColor = $colTextGray; $d.Location = "10,40"; $d.Size = "280, 60"; $p.Controls.Add($d)
        
        $b = New-Object System.Windows.Forms.Button; $b.Text = "Run"
        # Minimalist "Ghost" Button Style
        $b.BackColor = $colBtnSurface; $b.ForeColor = $Color; $b.FlatStyle = "Flat"; $b.FlatAppearance.BorderSize = 1; $b.FlatAppearance.BorderColor = $Color
        $b.Location = "10, 105"; $b.Tag = $ActionId
        $b.Add_Click({ $tag = $this.Tag; if ($tag -match "^DirectRun:(.+)$") { $cmd = $matches[1]; $args=$null; if ($cmd -match "^(\S+)\s+(.+)$") { $cmd=$matches[1]; $args=$matches[2] }; if ($args) { Start-Process $cmd $args } else { Start-Process $cmd } } else { $parts=$tag-split":",2; if($parts.Count-eq2){Start-BackgroundTask $parts[0] $parts[1]}else{Start-BackgroundTask $tag $null} } })
        $p.Controls.Add($b); return $p
    }

    # Maintenance Cards
    $flowMaint.Controls.Add((New-MaintCard "DNS: Google" "Set DNS to 8.8.8.8 & IPv6." "SetDNS:Google" $colAccentCyan))
    $flowMaint.Controls.Add((New-MaintCard "DNS: Cloudflare" "Set DNS to 1.1.1.1 & IPv6." "SetDNS:Cloudflare" $colAccentCyan))
    $flowMaint.Controls.Add((New-MaintCard "DNS: DHCP" "Reset DNS to automatic." "SetDNS:DHCP" $colAccentCyan))
    $flowMaint.Controls.Add((New-MaintCard "Update Repair" "Reset WU components." "Maint_ResetUpdates" $colAccentRed))
    $flowMaint.Controls.Add((New-MaintCard "System Scan" "Chkdsk -> SFC -> DISM -> SFC." "Maint_SystemScan" $colAccentRed))
    $flowMaint.Controls.Add((New-MaintCard "Deep Cleanup" "Disk Cleanup + Component Store." "Maint_Cleanup" $colAccentBlue))
    $flowMaint.Controls.Add((New-MaintCard "Network Reset" "Flush DNS, Reset IP/Winsock." "Maint_Network" $colAccentRed))
    $flowMaint.Controls.Add((New-MaintCard "Rebuild Icon Cache" "Fixes broken/blank icons." "Maint_IconCache" $colAccentOrange))
    $flowMaint.Controls.Add((New-MaintCard "Reinstall Package Manager" "Reinstall Winget/Choco." "Maint_ReinstallPackageManager" $colAccentPurple))
    $flowMaint.Controls.Add((New-MaintCard "Remove OneDrive" "Uninstall OneDrive." "Maint_RemoveOneDrive" $colAccentRed))
    $flowMaint.Controls.Add((New-MaintCard "O&O ShutUp10" "Run O&O ShutUp10." "Maint_OOSU10" $colAccentBlue))
    $flowMaint.Controls.Add((New-MaintCard "Enable SSH Server" "Install OpenSSH Server." "Maint_EnableSSH" $colAccentGreen))
    $flowMaint.Controls.Add((New-MaintCard "Run Adobe Cleaner" "Run Adobe CC Cleaner Tool." "Maint_AdobeClean" $colAccentOrange))
    $flowMaint.Controls.Add((New-MaintCard "Sysinternals Autoruns" "Run Autoruns." "Maint_Autoruns" $colAccentBlue))

    # Tools Cards
    $flowTools.Controls.Add((New-MaintCard "Control Panel" "Open Legacy Control Panel." "DirectRun:control" $colAccentGray))
    $flowTools.Controls.Add((New-MaintCard "Network Connections" "Open Network Adapter settings." "DirectRun:explorer ncpa.cpl" $colAccentGray))
    $flowTools.Controls.Add((New-MaintCard "Power Options" "Open Power Options." "DirectRun:explorer powercfg.cpl" $colAccentGray))
    $flowTools.Controls.Add((New-MaintCard "Sound Settings" "Open Sound settings." "DirectRun:explorer mmsys.cpl" $colAccentGray))
    $flowTools.Controls.Add((New-MaintCard "System Properties" "Open System Properties." "DirectRun:explorer sysdm.cpl" $colAccentGray))
    $flowTools.Controls.Add((New-MaintCard "User Accounts" "Open User Accounts." "DirectRun:netplwiz" $colAccentGray))
    $flowTools.Controls.Add((New-MaintCard "Registry Editor" "Open Registry Editor." "DirectRun:regedit" $colAccentGray))
    $flowTools.Controls.Add((New-MaintCard "Services" "Open Services." "DirectRun:mmc services.msc" $colAccentGray))
    $flowTools.Controls.Add((New-MaintCard "God Mode" "Open God Mode." "DirectRun:explorer shell:::{ED7BA470-8E54-465E-825C-99712043E01C}" $colAccentGray))
    $flowTools.Controls.Add((New-MaintCard "Task Manager" "Manage running processes." "DirectRun:taskmgr" $colAccentGray))
    $flowTools.Controls.Add((New-MaintCard "Restart Explorer" "Reload shell (Fix glitches)." "RestartExplorer" $colAccentGray))

    $script:currentView = $null
    function Set-View($View) {
        $script:currentView = $View
        $lvPower.Visible = $false; $lvTweaks.Visible = $false; $lvInstalls.Visible = $false; $flowMaint.Visible = $false; $flowTools.Visible = $false; $pnlWU.Visible = $false; $pnlPackageManager.Visible = $false; $pnlSearch.Visible = $false; $pnlActions.Controls.Clear()
        foreach ($key in $sync.MenuButtons.Keys) { $btn = $sync.MenuButtons[$key]; if ($key -eq $View) { $btn.BackColor = $colPanelBg; $btn.ForeColor = $colTextWhite; $btn.FlatAppearance.BorderSize = 4; $btn.FlatAppearance.BorderColor = $sync.ViewColors[$key] } else { $btn.BackColor = $colSideBar; $btn.ForeColor = $colTextGray; $btn.FlatAppearance.BorderSize = 0 } }
        
        switch ($View) {
            "Power" { 
                $lblHeaderTitle.Text = "Power Plans"; $lblHeaderTitle.ForeColor = $colAccentRed
                $lvPower.Visible = $true
                $pnlActions.Controls.Add((New-ActionButton "Set Active" $colAccentGreen { if ($lvPower.SelectedItems.Count -gt 0) { Start-BackgroundTask "SetPower" $lvPower.SelectedItems[0].Tag } }))
                $pnlActions.Controls.Add((New-ActionButton "Refresh List" $colAccentBlue { Start-BackgroundTask "LoadPower" $null }))
                Start-BackgroundTask "LoadPower" $null 
            }
            "Tweaks" { 
                $lblHeaderTitle.Text = "System Tweaks"; $lblHeaderTitle.ForeColor = $colAccentPurple
                $lvTweaks.Visible = $true
                $pnlActions.Controls.Add((New-ActionButton "Create Restore Point" $colAccentCyan { Start-BackgroundTask "CreateRestorePoint" $null }))
                $pnlActions.Controls.Add((New-ActionButton "Apply Selected" $colAccentGreen { $selected = @(); foreach ($item in $lvTweaks.CheckedItems) { $selected += $item.Tag }; if ($selected.Count -gt 0) { Start-BackgroundTask "ApplyTweakBatch" $selected } }))
                $pnlActions.Controls.Add((New-ActionButton "Undo Selected" $colAccentRed { $selected = @(); foreach ($item in $lvTweaks.CheckedItems) { $selected += $item.Tag }; if ($selected.Count -gt 0) { Start-BackgroundTask "UndoTweakBatch" $selected } }))
                $pnlActions.Controls.Add((New-ActionButton "Refresh Status" $colAccentBlue { Start-BackgroundTask "CheckTweaks" $null }))
                Start-BackgroundTask "CheckTweaks" $null 
            }
            "Installs" { 
                $lblHeaderTitle.Text = "Software Installer"; $lblHeaderTitle.ForeColor = $colAccentGreen
                $lvInstalls.Visible = $true; $pnlPackageManager.Visible = $true; $pnlSearch.Visible = $true
                
                # Install Button
                $pnlActions.Controls.Add((New-ActionButton "Install Selected" $colAccentGreen { 
                    $selected = @(); foreach ($item in $lvInstalls.CheckedItems) { $selected += $item.Tag }
                    if ($selected.Count -gt 0) { $selected | ForEach-Object { Start-BackgroundTask "InstallApp" $_ } } 
                }))
                
                # Update Button
                $pnlActions.Controls.Add((New-ActionButton "Update Selected" $colAccentBlue { 
                    $selected = @(); foreach ($item in $lvInstalls.CheckedItems) { $selected += $item.Tag }
                    if ($selected.Count -gt 0) { $selected | ForEach-Object { Start-BackgroundTask "UpgradeApp" $_ } } 
                }))

                # Delete (Uninstall) Button
                $pnlActions.Controls.Add((New-ActionButton "Delete Selected" $colAccentRed { 
                    $selected = @(); foreach ($item in $lvInstalls.CheckedItems) { $selected += $item.Tag }
                    if ($selected.Count -gt 0) { $selected | ForEach-Object { Start-BackgroundTask "UninstallApp" $_ } } 
                }))
                
                # Refresh Button
                $pnlActions.Controls.Add((New-ActionButton "Refresh List" $colAccentBlue { 
                    Start-BackgroundTask "CheckInstalled" $null 
                }))

                # Export Button
                $pnlActions.Controls.Add((New-ActionButton "Export List" $colAccentGray { 
                    Start-BackgroundTask "ExportList" $null 
                }))

                # Perdanger's Preset Button
                $presetIds = @("Brave.Brave", "Anysphere.Cursor", "Discord.Discord", "AdrienAllard.FileConverter", "Git.Git", "DuongDieuPhap.ImageGlass", "Nvidia.NvidiaApp", "qBittorrent.qBittorrent", "RevoUninstaller.RevoUninstaller", "Spotify.Spotify", "Valve.Steam", "Telegram.TelegramDesktop", "VideoLAN.VLC", "RARLab.WinRAR", "AntibodySoftware.WizTree")
                $tooltipText = "Installs Perdanger's Preset:`n- Brave`n- Cursor IDE`n- Discord`n- File Converter`n- Git`n- ImageGlass`n- Nvidia App`n- qBittorrent`n- Revo Uninstaller`n- Spotify`n- Steam`n- Telegram`n- VLC`n- WinRAR`n- WizTree"
                $pnlActions.Controls.Add((New-ActionButtonWithTooltip "Perdanger's Preset" $colAccentPurple $tooltipText { 
                    $presetIds | ForEach-Object { Start-BackgroundTask "InstallApp" $_ } 
                }))

                if ($sync.InstalledApps.Count -eq 0) { Start-BackgroundTask "CheckInstalled" $null } else { Update-ListView "Installs" } 
            }
            "Maintenance" { 
                $lblHeaderTitle.Text = "System Maintenance"; $lblHeaderTitle.ForeColor = $colAccentOrange
                $flowMaint.Visible = $true 
            }
            "Tools" { 
                $lblHeaderTitle.Text = "Quick Tools"; $lblHeaderTitle.ForeColor = $colAccentGray
                $flowTools.Visible = $true 
            }
            "WindowsUpdate" { 
                $lblHeaderTitle.Text = "Windows Update Config"; $lblHeaderTitle.ForeColor = $colAccentBlue
                $pnlWU.Visible = $true
                Start-BackgroundTask "WU_GetStatus" $null
                $pnlActions.Controls.Add((New-ActionButton "Apply Settings" $colAccentGreen { $action = "WU_SetDefault"; if ($rbSec.Checked) { $action = "WU_SetSecurity" } elseif ($rbDis.Checked) { $action = "WU_SetDisabled" }; Start-BackgroundTask $action $null }))
                $pnlActions.Controls.Add((New-ActionButton "Refresh Status" $colAccentBlue { Start-BackgroundTask "WU_GetStatus" $null })) 
            }
        }
    }

    $form.Add_FormClosing({ Write-Host "[STOP] Shutting down System Manager..." -ForegroundColor Cyan; $sync.RunspacePool.Close(); $sync.RunspacePool.Dispose() })
    Set-View "Tweaks"
    [System.Windows.Forms.Application]::Run($form)
}

# ENHANCED FUNCTION: Create a detailed autounattend.xml file via GUI with regional settings and tooltips.
function Create-UnattendXml {
    if (-not $script:guiAvailable) {
        Write-LogAndHost "GUI is not available, cannot launch the Unattend XML Creator." -HostColor Red -LogPrefix "Create-UnattendXml"
        Start-Sleep -Seconds 2
        return
    }
    
    # Check if running with elevated privileges for certain operations
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
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
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$PSHOME\powershell.exe")

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
    $errorColor = [System.Drawing.Color]::FromArgb(200, 50, 50)

    # Helper functions for creating styled controls.
    function New-StyledLabel($Text, $Location, $Size = $null) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $Text
        $label.Location = $Location
        if ($Size) { $label.Size = $Size } else { $label.AutoSize = $true }
        $label.Font = $commonFont
        $label.ForeColor = $labelColor
        return $label
    }
    
    function New-StyledTextBox($Location, $Size, $Multiline = $false) {
        $textbox = New-Object System.Windows.Forms.TextBox
        $textbox.Location = $Location
        $textbox.Size = $Size
        $textbox.Font = $commonFont
        $textbox.BackColor = $controlBackColor
        $textbox.ForeColor = $controlForeColor
        $textbox.BorderStyle = "FixedSingle"
        $textbox.Multiline = $Multiline
        return $textbox
    }
    
    function New-StyledComboBox($Location, $Size) {
        $combobox = New-Object System.Windows.Forms.ComboBox
        $combobox.Location = $Location
        $combobox.Size = $Size
        $combobox.Font = $commonFont
        $combobox.BackColor = $controlBackColor
        $combobox.ForeColor = $controlForeColor
        $combobox.FlatStyle = "Flat"
        return $combobox
    }
    
    function New-StyledCheckBox($Text, $Location, $Checked) {
        $checkbox = New-Object System.Windows.Forms.CheckBox
        $checkbox.Text = $Text
        $checkbox.Location = $Location
        $checkbox.Font = $commonFont
        $checkbox.ForeColor = $labelColor
        $checkbox.AutoSize = $true
        $checkbox.Checked = $Checked
        return $checkbox
    }
    
    function New-StyledGroupBox($Text, $Location, $Size) {
        $groupbox = New-Object System.Windows.Forms.GroupBox
        $groupbox.Text = $Text
        $groupbox.Location = $Location
        $groupbox.Size = $Size
        $groupbox.Font = $commonFont
        $groupbox.ForeColor = $groupboxForeColor
        return $groupbox
    }
    
    # --- Tab 1: General Settings ---
    $tabGeneral = New-Object System.Windows.Forms.TabPage
    $tabGeneral.Text = "General"
    $tabGeneral.BackColor = $form.BackColor
    $tabControl.Controls.Add($tabGeneral) | Out-Null
    $yPos = 30
    
    # Computer Name
    $tabGeneral.Controls.Add((New-StyledLabel -Text "Computer Name:" -Location "20,$yPos")) | Out-Null
    $textComputerName = New-StyledTextBox -Location "180,$yPos" -Size "280,20"
    $textComputerName.Text = "DESKTOP-PC"
    $tabGeneral.Controls.Add($textComputerName) | Out-Null
    $labelComputerNameError = New-StyledLabel -Text "" -Location "470,$yPos" -Size "300,20"
    $labelComputerNameError.ForeColor = $errorColor
    $tabGeneral.Controls.Add($labelComputerNameError) | Out-Null
    $toolTip.SetToolTip($textComputerName, "Enter a name for the computer (15 characters max, no special characters except hyphen).")
    
    # Validate computer name
    $textComputerName.Add_TextChanged({
        $value = $textComputerName.Text
        $labelComputerNameError.Text = ""
        
        if ($value.Length -gt 15) {
            $labelComputerNameError.Text = "Too long (max 15 chars)"
        } elseif ($value -match '[^a-zA-Z0-9-]') {
            $labelComputerNameError.Text = "Invalid characters"
        } elseif ($value -match '^-|-$') {
            $labelComputerNameError.Text = "Cannot start/end with hyphen"
        }
    })
    $yPos += 40
    
    # Admin User Name
    $tabGeneral.Controls.Add((New-StyledLabel -Text "Admin User Name:" -Location "20,$yPos")) | Out-Null
    $textUserName = New-StyledTextBox -Location "180,$yPos" -Size "280,20"
    $textUserName.Text = "Admin"
    $tabGeneral.Controls.Add($textUserName) | Out-Null
    $labelUserNameError = New-StyledLabel -Text "" -Location "470,$yPos" -Size "300,20"
    $labelUserNameError.ForeColor = $errorColor
    $tabGeneral.Controls.Add($labelUserNameError) | Out-Null
    $toolTip.SetToolTip($textUserName, "Enter a username for the administrator account.")
    
    # Validate username
    $textUserName.Add_TextChanged({
        $value = $textUserName.Text
        $labelUserNameError.Text = ""
        
        if ($value.Length -eq 0) {
            $labelUserNameError.Text = "Username cannot be empty"
        } elseif ($value.Length -gt 20) {
            $labelUserNameError.Text = "Too long (max 20 chars)"
        } elseif ($value -match '[./\\[\]:;|=,+*?<>]') {
            $labelUserNameError.Text = "Invalid characters"
        } elseif ($value -eq "Administrator" -or $value -eq "Guest") {
            $labelUserNameError.Text = "Reserved username"
        }
    })
    $yPos += 40

    # Password
    $tabGeneral.Controls.Add((New-StyledLabel -Text "Password (optional):" -Location "20,$yPos")) | Out-Null
    $textPassword = New-StyledTextBox -Location "180,$yPos" -Size "280,20"
    $textPassword.UseSystemPasswordChar = $true
    $textPassword.MaxLength = 127
    $tabGeneral.Controls.Add($textPassword) | Out-Null
    $labelPasswordCounter = New-StyledLabel -Text "0/127" -Location "470,$yPos"
    $tabGeneral.Controls.Add($labelPasswordCounter) | Out-Null
    $labelPasswordError = New-StyledLabel -Text "" -Location "530,$yPos" -Size "240,20"
    $labelPasswordError.ForeColor = $errorColor
    $tabGeneral.Controls.Add($labelPasswordError) | Out-Null
    $toolTip.SetToolTip($textPassword, "Enter a password for the administrator account (optional). If left blank, no password will be set.")
    
    # Password strength indicator
    $passwordStrengthLabel = New-StyledLabel -Text "" -Location "180,$($yPos+25)" -Size "280,20"
    $tabGeneral.Controls.Add($passwordStrengthLabel) | Out-Null
    
    # Validate password
    $textPassword.Add_TextChanged({
        $length = $textPassword.Text.Length
        $labelPasswordCounter.Text = "$length/127"
        $labelPasswordError.Text = ""
        $passwordStrengthLabel.Text = ""
        
        if ($length -eq 127) {
            $labelPasswordCounter.ForeColor = [System.Drawing.Color]::Crimson
        } else {
            $labelPasswordCounter.ForeColor = $labelColor
        }
        
        # Password strength check (only if password is provided)
        if ($length -gt 0) {
            $strength = 0
            if ($length -ge 8) { $strength++ }
            if ($textPassword.Text -match '[A-Z]') { $strength++ }
            if ($textPassword.Text -match '[a-z]') { $strength++ }
            if ($textPassword.Text -match '[0-9]') { $strength++ }
            if ($textPassword.Text -match '[^a-zA-Z0-9]') { $strength++ }
            
            switch ($strength) {
                0 { $passwordStrengthLabel.Text = "Very Weak"; $passwordStrengthLabel.ForeColor = [System.Drawing.Color]::Crimson }
                1 { $passwordStrengthLabel.Text = "Weak"; $passwordStrengthLabel.ForeColor = [System.Drawing.Color]::OrangeRed }
                2 { $passwordStrengthLabel.Text = "Fair"; $passwordStrengthLabel.ForeColor = [System.Drawing.Color]::Orange }
                3 { $passwordStrengthLabel.Text = "Good"; $passwordStrengthLabel.ForeColor = [System.Drawing.Color]::YellowGreen }
                4 { $passwordStrengthLabel.Text = "Strong"; $passwordStrengthLabel.ForeColor = [System.Drawing.Color]::Green }
                5 { $passwordStrengthLabel.Text = "Very Strong"; $passwordStrengthLabel.ForeColor = [System.Drawing.Color]::DarkGreen }
            }
            
            if ($length -lt 8) {
                $labelPasswordError.Text = "Weak (min 8 chars recommended)"
            }
        } else {
            $passwordStrengthLabel.Text = "No password will be set"
            $passwordStrengthLabel.ForeColor = [System.Drawing.Color]::Gray
        }
    })
    $yPos += 50

    # --- Tab 2: Regional Settings ---
    $tabRegional = New-Object System.Windows.Forms.TabPage
    $tabRegional.Text = "Regional"
    $tabRegional.BackColor = $form.BackColor
    $tabRegional.Padding = New-Object System.Windows.Forms.Padding(10)
    $tabControl.Controls.Add($tabRegional) | Out-Null

    $groupLocale = New-StyledGroupBox "Language & Locale" "15,15" "750,150"
    $tabRegional.Controls.Add($groupLocale) | Out-Null
    $yPos = 30
    $commonLocales = @("ar-SA", "cs-CZ", "da-DK", "de-DE", "el-GR", "en-GB", "en-US", "es-ES", "es-MX", "fi-FI", "fr-CA", "fr-FR", "he-IL", "hu-HU", "it-IT", "ja-JP", "ko-KR", "nb-NO", "nl-NL", "pl-PL", "pt-BR", "pt-PT", "ro-RO", "ru-RU", "sk-SK", "sv-SE", "th-TH", "tr-TR", "zh-CN", "zh-TW")
    
    # UI Language
    $groupLocale.Controls.Add((New-StyledLabel -Text "UI Language:" -Location "15,$yPos")) | Out-Null
    $comboUiLanguage = New-StyledComboBox -Location "150,$yPos" -Size "250,20"
    $comboUiLanguage.Items.AddRange($commonLocales) | Out-Null
    $comboUiLanguage.Text = (Get-UICulture).Name
    $groupLocale.Controls.Add($comboUiLanguage) | Out-Null
    $groupLocale.Controls.Add((New-StyledLabel -Text "(e.g., en-US, de-DE)" -Location "410,$yPos")) | Out-Null
    $toolTip.SetToolTip($comboUiLanguage, "Select the language for the user interface.")
    $yPos += 40
    
    # System Locale
    $groupLocale.Controls.Add((New-StyledLabel -Text "System Locale:" -Location "15,$yPos")) | Out-Null
    $comboSystemLocale = New-StyledComboBox -Location "150,$yPos" -Size "250,20"
    $comboSystemLocale.Items.AddRange($commonLocales) | Out-Null
    $comboSystemLocale.Text = (Get-Culture).Name
    $groupLocale.Controls.Add($comboSystemLocale) | Out-Null
    $groupLocale.Controls.Add((New-StyledLabel -Text "(e.g., en-US, ja-JP)" -Location "410,$yPos")) | Out-Null
    $toolTip.SetToolTip($comboSystemLocale, "Select the system locale for non-Unicode programs.")
    $yPos += 40
    
    # User Locale
    $groupLocale.Controls.Add((New-StyledLabel -Text "User Locale:" -Location "15,$yPos")) | Out-Null
    $comboUserLocale = New-StyledComboBox -Location "150,$yPos" -Size "250,20"
    $comboUserLocale.Items.AddRange($commonLocales) | Out-Null
    $comboUserLocale.Text = (Get-Culture).Name
    $groupLocale.Controls.Add($comboUserLocale) | Out-Null
    $groupLocale.Controls.Add((New-StyledLabel -Text "(e.g., en-US, tr-TR)" -Location "410,$yPos")) | Out-Null
    $toolTip.SetToolTip($comboUserLocale, "Select the user locale for numbers, currency, date, and time formats.")

    # Time Zone Group
    $groupTimeZone = New-StyledGroupBox "Time Zone" "15,180" "750,220"
    $tabRegional.Controls.Add($groupTimeZone) | Out-Null
    $yPos = 30
    
    # Time Zone Search
    $groupTimeZone.Controls.Add((New-StyledLabel -Text "Search:" -Location "15,$yPos")) | Out-Null
    $textTimeZoneSearch = New-StyledTextBox -Location "85,$yPos" -Size "645,20"
    $groupTimeZone.Controls.Add($textTimeZoneSearch) | Out-Null
    $toolTip.SetToolTip($textTimeZoneSearch, "Type to filter time zones by name or offset.")
    $yPos += 35
    
    # Time Zone List
    $listTimeZone = New-Object System.Windows.Forms.ListBox
    $listTimeZone.Location = "15,$yPos"
    $listTimeZone.Size = "715,100"
    $listTimeZone.Font = $commonFont
    $listTimeZone.BackColor = $controlBackColor
    $listTimeZone.ForeColor = $controlForeColor
    $toolTip.SetToolTip($listTimeZone, "Select your time zone from the list.")
    
    # Windows 11 Time Zone IDs
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

    # Get all time zones with fallback
    $allTimeZonesInfo = try { 
        $windows11TimeZoneIds | ForEach-Object { [System.TimeZoneInfo]::FindSystemTimeZoneById($_) }
    } catch { 
        Write-LogAndHost "Could not find all static time zones. The list may be incomplete. Falling back to system's available time zones." -HostColor Yellow -LogPrefix "Create-UnattendXml"
        [System.TimeZoneInfo]::GetSystemTimeZones() 
    }
    
    # Format time zones for display
    $formattedTimeZones = foreach ($tz in $allTimeZonesInfo) {
        $offset = $tz.BaseUtcOffset
        $offsetSign = if ($offset.Ticks -ge 0) { "+" } else { "-" }
        $offsetString = "{0:hh\:mm}" -f $offset
        $displayString = "(UTC{0}{1}) {2}" -f $offsetSign, $offsetString, $tz.Id
        $timeZoneMap[$displayString] = $tz.Id
        $displayString
    }
    $sortedFormattedTimeZones = $formattedTimezones | Sort-Object
    
    # Populate time zone list
    if ($null -ne $sortedFormattedTimeZones) { $listTimeZone.Items.AddRange($sortedFormattedTimeZones) | Out-Null }
    
    # Set current time zone
    try {
        $currentTimeZoneId = (Get-TimeZone).Id
        $currentFormattedTz = $timeZoneMap.GetEnumerator() | Where-Object { $_.Value -eq $currentTimeZoneId } | Select-Object -First 1 -ExpandProperty Key
        if ($currentFormattedTz) { $listTimeZone.SelectedItem = $currentFormattedTz }
    } catch {}

    $groupTimeZone.Controls.Add($listTimeZone) | Out-Null; $yPos += $listTimeZone.Height + 10

    # Selected time zone display
    $groupTimeZone.Controls.Add((New-StyledLabel -Text "Current Selection:" -Location "15,$yPos")) | Out-Null
    $labelSelectedTimeZone = New-StyledLabel -Text "None" -Location "150,$yPos"
    $labelSelectedTimeZone.ForeColor = [System.Drawing.Color]::LightSteelBlue
    $labelSelectedTimeZone.AutoSize = $false
    $labelSelectedTimeZone.Size = '580,20'
    $groupTimeZone.Controls.Add($labelSelectedTimeZone) | Out-Null
    
    # Time zone selection event handlers
    $listTimeZone.Add_SelectedIndexChanged({
        if ($listTimeZone.SelectedItem) { 
            $labelSelectedTimeZone.Text = $listTimeZone.SelectedItem 
        } else { 
            $labelSelectedTimeZone.Text = "None" 
        }
    }) | Out-Null
    
    $textTimeZoneSearch.Add_TextChanged({
        $selected = $listTimeZone.SelectedItem
        $listTimeZone.BeginUpdate()
        $listTimeZone.Items.Clear()
        $searchText = $textTimeZoneSearch.Text
        $filteredTimeZones = $sortedFormattedTimeZones | Where-Object { $_ -match [regex]::Escape($searchText) }
        if ($null -ne $filteredTimeZones) { $listTimeZone.Items.AddRange($filteredTimeZones) | Out-Null }
        if ($selected -and $listTimeZone.Items.Contains($selected)) { 
            $listTimeZone.SelectedItem = $selected 
        } elseif ($listTimeZone.Items.Count -gt 0) { 
            $listTimeZone.SelectedIndex = 0 
        }
        $listTimeZone.EndUpdate()
    }) | Out-Null
    
    if ($listTimeZone.SelectedItem) { $labelSelectedTimeZone.Text = $listTimeZone.SelectedItem }

    # Keyboard Layouts Group
    $groupKeyboard = New-StyledGroupBox "Keyboard Layouts (select up to 5)" "15,415" "750,245"
    $tabRegional.Controls.Add($groupKeyboard) | Out-Null
    $yPos = 30
    
    # Keyboard Layout Search
    $groupKeyboard.Controls.Add((New-StyledLabel -Text "Search:" -Location "15,$yPos")) | Out-Null
    $textKeyboardSearch = New-StyledTextBox -Location "85,$yPos" -Size "645,20"
    $groupKeyboard.Controls.Add($textKeyboardSearch) | Out-Null
    $toolTip.SetToolTip($textKeyboardSearch, "Type to filter keyboard layouts by name.")
    $yPos += 35
    
    # Keyboard Layouts List
    $listKeyboardLayouts = New-Object System.Windows.Forms.CheckedListBox
    $listKeyboardLayouts.Location = "15,$yPos"
    $listKeyboardLayouts.Size = "715,110"
    $listKeyboardLayouts.Font = $commonFont
    $listKeyboardLayouts.BackColor = $controlBackColor
    $listKeyboardLayouts.ForeColor = $controlForeColor
    $listKeyboardLayouts.CheckOnClick = $true
    $toolTip.SetToolTip($listKeyboardLayouts, "Select up to 5 keyboard layouts. The first selected will be the default.")
    
    # Keyboard layout data
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

    # Selected keyboard layouts display
    $groupKeyboard.Controls.Add((New-StyledLabel -Text "Current Selection:" -Location "15,$yPos")) | Out-Null
    $labelSelectedKeyboards = New-StyledLabel -Text "None" -Location "150,$yPos"
    $labelSelectedKeyboards.ForeColor = [System.Drawing.Color]::LightSteelBlue
    $labelSelectedKeyboards.AutoSize = $false
    $labelSelectedKeyboards.Size = '580,50'
    $groupKeyboard.Controls.Add($labelSelectedKeyboards) | Out-Null

    # Update keyboard label function
    $updateKeyboardLabel = {
        $checkedItemsText = ($checkedKeyboardLayoutNames | Sort-Object) -join ', '
        if ([string]::IsNullOrWhiteSpace($checkedItemsText)) { 
            $labelSelectedKeyboards.Text = "None" 
        } else { 
            $labelSelectedKeyboards.Text = $checkedItemsText 
        }
    }

    # Keyboard layout selection event handlers
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
    
    # Set default keyboard layout
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
    $tabAutomation = New-Object System.Windows.Forms.TabPage
    $tabAutomation.Text = "Automation & Tweaks"
    $tabAutomation.BackColor = $form.BackColor
    $tabAutomation.Padding = New-Object System.Windows.Forms.Padding(10)
    $tabControl.Controls.Add($tabAutomation) | Out-Null

    # OOBE Skip Options Group
    $groupOobe = New-StyledGroupBox "OOBE Skip Options" "15,15" "750,220"
    $tabAutomation.Controls.Add($groupOobe) | Out-Null
    $yPos = 30
    
    # Hide EULA Page
    $checkHideEula = New-StyledCheckBox -Text "Hide EULA Page" -Location "20,$yPos" -Checked $true
    $groupOobe.Controls.Add($checkHideEula) | Out-Null
    $toolTip.SetToolTip($checkHideEula, "Automatically accepts the End User License Agreement (EULA) during setup.")
    $yPos += 40
    
    # Hide Local Account Screen
    $checkHideLocalAccount = New-StyledCheckBox -Text "Hide Local Account Screen" -Location "20,$yPos" -Checked $true
    $groupOobe.Controls.Add($checkHideLocalAccount) | Out-Null
    $toolTip.SetToolTip($checkHideLocalAccount, "Bypasses the screen that prompts to create a local user account.")
    $yPos += 40
    
    # Hide Online Account Screens
    $checkHideOnlineAccount = New-StyledCheckBox -Text "Hide Online Account Screens" -Location "20,$yPos" -Checked $true
    $groupOobe.Controls.Add($checkHideOnlineAccount) | Out-Null
    $toolTip.SetToolTip($checkHideOnlineAccount, "Bypasses the screens that prompt to sign in with or create a Microsoft Account.")
    $yPos += 40
    
    # Hide Wireless Setup
    $checkHideWireless = New-StyledCheckBox -Text "Hide Wireless Setup" -Location "20,$yPos" -Checked $true
    $groupOobe.Controls.Add($checkHideWireless) | Out-Null
    $toolTip.SetToolTip($checkHideWireless, "Skips the network and Wi-Fi connection screen during the Out-of-Box Experience (OOBE).")

    # First Logon System Tweaks Group
    $groupCustom = New-StyledGroupBox "First Logon System Tweaks" "15,250" "750,220"
    $tabAutomation.Controls.Add($groupCustom) | Out-Null
    $yPos = 30
    
    # Show File Extensions
    $checkShowFileExt = New-StyledCheckBox -Text "Show Known File Extensions" -Location "20,$yPos" -Checked $true
    $groupCustom.Controls.Add($checkShowFileExt) | Out-Null
    $toolTip.SetToolTip($checkShowFileExt, "Configures File Explorer to show file extensions like '.exe', '.txt', '.dll' by default.")
    $yPos += 40
    
    # Disable SmartScreen
    $checkDisableSmartScreen = New-StyledCheckBox -Text "Disable SmartScreen" -Location "20,$yPos" -Checked $true
    $groupCustom.Controls.Add($checkDisableSmartScreen) | Out-Null
    $toolTip.SetToolTip($checkDisableSmartScreen, "Turns off the Microsoft Defender SmartScreen filter, which checks for malicious files and websites.")
    $yPos += 40
    
    # Disable System Restore
    $checkDisableSysRestore = New-StyledCheckBox -Text "Disable System Restore" -Location "20,$yPos" -Checked $true
    $groupCustom.Controls.Add($checkDisableSysRestore) | Out-Null
    $toolTip.SetToolTip($checkDisableSysRestore, "Disables the automatic creation of restore points. This can save disk space but limits recovery options.")
    $yPos += 40
    
    # Disable App Suggestions
    $checkDisableSuggestions = New-StyledCheckBox -Text "Disable App Suggestions" -Location "20,$yPos" -Checked $true
    $groupCustom.Controls.Add($checkDisableSuggestions) | Out-Null
    $toolTip.SetToolTip($checkDisableSuggestions, "Prevents Windows from displaying app and content suggestions in the Start Menu and on the lock screen.")

    # Info label
    $automationInfoLabel = New-Object System.Windows.Forms.Label
    $automationInfoLabel.Text = "Hover over an option for a detailed description."
    $automationInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $automationInfoLabel.ForeColor = [System.Drawing.Color]::Gray
    $automationInfoLabel.AutoSize = $true
    $automationInfoLabel.Location = New-Object System.Drawing.Point(20, 485)
    $tabAutomation.Controls.Add($automationInfoLabel) | Out-Null

    # --- Tab 4: Bloatware Removal ---
    $tabBloatware = New-Object System.Windows.Forms.TabPage
    $tabBloatware.Text = "Bloatware"
    $tabBloatware.BackColor = $form.BackColor
    $tabControl.Controls.Add($tabBloatware) | Out-Null
    
    # Bloatware top panel
    $bloatTopPanel = New-Object System.Windows.Forms.Panel
    $bloatTopPanel.Dock = "Top"
    $bloatTopPanel.Height = 40
    $bloatTopPanel.BackColor = $form.BackColor
    $tabBloatware.Controls.Add($bloatTopPanel) | Out-Null
    
    # Bloatware table panel
    $bloatTablePanel = New-Object System.Windows.Forms.TableLayoutPanel
    $bloatTablePanel.Dock = "Fill"
    $bloatTablePanel.AutoScroll = $true
    $bloatTablePanel.BackColor = $form.BackColor
    $tabBloatware.Controls.Add($bloatTablePanel) | Out-Null
    $bloatTablePanel.BringToFront()
    
    # Bloatware bottom panel
    $bloatBottomPanel = New-Object System.Windows.Forms.Panel
    $bloatBottomPanel.Dock = "Bottom"
    $bloatBottomPanel.Height = 40
    $bloatBottomPanel.BackColor = $form.BackColor
    $tabBloatware.Controls.Add($bloatBottomPanel) | Out-Null
    
    # Bloatware checkboxes
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
    
    # Create bloatware table layout
    $bloatTablePanel.ColumnCount = 3
    $rowsNeeded = [math]::Ceiling($bloatwareList.Count / $bloatTablePanel.ColumnCount)
    $bloatTablePanel.RowCount = $rowsNeeded
    for ($i = 0; $i -lt $bloatTablePanel.ColumnCount; $i++) { 
        $bloatTablePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) | Out-Null 
    }
    for ($i = 0; $i -lt $bloatTablePanel.RowCount; $i++) { 
        $bloatTablePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null 
    }
    
    # Add bloatware checkboxes
    $col = 0; $row = 0
    foreach ($appName in $bloatwareList) {
        $checkbox = New-StyledCheckBox -Text $appName -Location "0,0" -Checked $false
        $checkbox.Margin = [System.Windows.Forms.Padding]::new(10, 5, 10, 5)
        $bloatTablePanel.Controls.Add($checkbox, $col, $row) | Out-Null
        $bloatwareCheckboxes += $checkbox
        $col++
        if ($col -ge $bloatTablePanel.ColumnCount) { $col = 0; $row++ }
    }
    
    # Select All button
    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Text = "Select All"
    $btnSelectAll.Size = "120,30"
    $btnSelectAll.Location = "10,5"
    $btnSelectAll.Font = $commonFont
    $btnSelectAll.ForeColor = [System.Drawing.Color]::White
    $btnSelectAll.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $btnSelectAll.FlatStyle = "Flat"
    $btnSelectAll.FlatAppearance.BorderSize = 0
    $btnSelectAll.add_Click({ 
        foreach($cb in $bloatwareCheckboxes) {$cb.Checked = $true} 
    }) | Out-Null
    $bloatTopPanel.Controls.Add($btnSelectAll) | Out-Null
    $toolTip.SetToolTip($btnSelectAll, "Select all bloatware items for removal.")
    
    # Deselect All button
    $btnDeselectAll = New-Object System.Windows.Forms.Button
    $btnDeselectAll.Text = "Deselect All"
    $btnDeselectAll.Size = "120,30"
    $btnDeselectAll.Location = "140,5"
    $btnDeselectAll.Font = $commonFont
    $btnDeselectAll.ForeColor = [System.Drawing.Color]::White
    $btnDeselectAll.BackColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $btnDeselectAll.FlatStyle = "Flat"
    $btnDeselectAll.FlatAppearance.BorderSize = 0
    $btnDeselectAll.add_Click({ 
        foreach($cb in $bloatwareCheckboxes) {$cb.Checked = $false} 
    }) | Out-Null
    $bloatTopPanel.Controls.Add($btnDeselectAll) | Out-Null
    $toolTip.SetToolTip($btnDeselectAll, "Deselect all bloatware items.")
    
    # Recommended button
    $btnRecommended = New-Object System.Windows.Forms.Button
    $btnRecommended.Text = "Recommended"
    $btnRecommended.Size = "120,30"
    $btnRecommended.Location = "270,5"
    $btnRecommended.Font = $commonFont
    $btnRecommended.ForeColor = [System.Drawing.Color]::White
    $btnRecommended.BackColor = [System.Drawing.Color]::FromArgb(60, 120, 60)
    $btnRecommended.FlatStyle = "Flat"
    $btnRecommended.FlatAppearance.BorderSize = 0
    $btnRecommended.add_Click({ 
        # Deselect all first
        foreach($cb in $bloatwareCheckboxes) {$cb.Checked = $false}
        
        # Select recommended items
        $recommendedItems = @(
            '3D Viewer', 'Clipchamp', 'Copilot', 'Cortana', 'Dev Home', 'Feedback Hub', 'Get Help', 
            'Mail and Calendar', 'Mixed Reality', 'Movies & TV', 'News', 'Notepad (modern)', 
            'Office 365', 'OneDrive', 'OneNote', 'Paint 3D', 'People', 'Power Automate', 
            'Skype', 'Solitaire Collection', 'Teams', 'Tips', 'To Do', 'Voice Recorder', 
            'Wallet', 'Weather', 'Windows Media Player (modern)', 'Xbox Apps', 'Your Phone / Phone Link'
        )
        
        foreach($cb in $bloatwareCheckboxes) {
            if ($recommendedItems -contains $cb.Text) {
                $cb.Checked = $true
            }
        }
    }) | Out-Null
    $bloatTopPanel.Controls.Add($btnRecommended) | Out-Null
    $toolTip.SetToolTip($btnRecommended, "Select commonly recommended bloatware items for removal.")
    
    # Info label
    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = "Bloatware removal works best with original Win 10 and 11 ISOs. Functionality on custom images is not guaranteed."
    $infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $infoLabel.ForeColor = [System.Drawing.Color]::Gray
    $infoLabel.Dock = "Fill"
    $infoLabel.TextAlign = "MiddleCenter"
    $bloatBottomPanel.Controls.Add($infoLabel) | Out-Null

    # --- Create and Cancel Buttons ---
    $buttonCreate = New-Object System.Windows.Forms.Button
    $buttonCreate.Text = "Create"
    $buttonCreate.Size = "120,30"
    $buttonCreate.Location = "265,10"
    $buttonCreate.Font = $commonFont
    $buttonCreate.ForeColor = [System.Drawing.Color]::White
    $buttonCreate.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $buttonCreate.FlatStyle = "Flat"
    $buttonCreate.FlatAppearance.BorderSize = 0
    $buttonCreate.add_Click({
        # Validate all inputs before proceeding
        $validationErrors = @()
        
        # Validate computer name
        if ([string]::IsNullOrWhiteSpace($textComputerName.Text)) {
            $validationErrors += "Computer name cannot be empty"
        } elseif ($textComputerName.Text.Length -gt 15) {
            $validationErrors += "Computer name too long (max 15 characters)"
        } elseif ($textComputerName.Text -match '[^a-zA-Z0-9-]') {
            $validationErrors += "Computer name contains invalid characters"
        } elseif ($textComputerName.Text -match '^-|-$') {
            $validationErrors += "Computer name cannot start or end with a hyphen"
        }
        
        # Validate username
        if ([string]::IsNullOrWhiteSpace($textUserName.Text)) {
            $validationErrors += "Username cannot be empty"
        } elseif ($textUserName.Text.Length -gt 20) {
            $validationErrors += "Username too long (max 20 characters)"
        } elseif ($textUserName.Text -match '[./\\[\]:;|=,+*?<>]') {
            $validationErrors += "Username contains invalid characters"
        } elseif ($textUserName.Text -eq "Administrator" -or $textUserName.Text -eq "Guest") {
            $validationErrors += "Cannot use reserved username"
        }
        
        # Password is now optional, but if provided, validate strength
        if ($textPassword.Text.Length -gt 0 -and $textPassword.Text.Length -lt 8) {
            $validationErrors += "Password too short (min 8 characters)"
        }
        
        # Validate regional settings
        if ([string]::IsNullOrWhiteSpace($comboUiLanguage.Text)) {
            $validationErrors += "UI Language must be selected"
        }
        
        if ([string]::IsNullOrWhiteSpace($comboSystemLocale.Text)) {
            $validationErrors += "System Locale must be selected"
        }
        
        if ([string]::IsNullOrWhiteSpace($comboUserLocale.Text)) {
            $validationErrors += "User Locale must be selected"
        }
        
        if ($null -eq $listTimeZone.SelectedItem) {
            $validationErrors += "Time Zone must be selected"
        }
        
        if ($checkedKeyboardLayoutNames.Count -eq 0) {
            $validationErrors += "At least one keyboard layout must be selected"
        }
        
        # Show validation errors if any
        if ($validationErrors.Count -gt 0) {
            $errorMessage = "Please correct the following errors:`n`n" + ($validationErrors -join "`n")
            [System.Windows.Forms.MessageBox]::Show($errorMessage, "Validation Failed", "OK", "Error") | Out-Null
            Write-LogAndHost "XML creation aborted due to validation errors." -HostColor Red -LogPrefix "Create-UnattendXml"
            return
        }
        
        # If all validations pass, close the form with OK result
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }) | Out-Null
    $buttonPanel.Controls.Add($buttonCreate) | Out-Null
    $toolTip.SetToolTip($buttonCreate, "Generate the autounattend.xml file with your selected settings.")
    
    $buttonCancel = New-Object System.Windows.Forms.Button
    $buttonCancel.Text = "Cancel"
    $buttonCancel.Size = "120,30"
    $buttonCancel.Location = "395,10"
    $buttonCancel.Font = $commonFont
    $buttonCancel.ForeColor = [System.Drawing.Color]::White
    $buttonCancel.BackColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $buttonCancel.FlatStyle = "Flat"
    $buttonCancel.FlatAppearance.BorderSize = 0
    $buttonCancel.add_Click({
        $form.Close()
    }) | Out-Null
    $buttonPanel.Controls.Add($buttonCancel) | Out-Null
    $toolTip.SetToolTip($buttonCancel, "Cancel and return to the main menu.")

    # Show the form and handle the result
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
        Write-LogAndHost "XML creation cancelled by user." -HostColor Yellow
        Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
        $null = Read-Host
        return
    }

    # Collect form data
    $selectedKeyboardLayouts = $checkedKeyboardLayoutNames | ForEach-Object { $keyboardLayoutData[$_] }
    $selectedTimeZoneId = if ($listTimeZone.SelectedItem) { $timeZoneMap[$listTimeZone.SelectedItem] } else { $null }

    $formData = @{
        ComputerName = $textComputerName.Text
        UserName = $textUserName.Text
        Password = $textPassword.Text
        UiLanguage = $comboUiLanguage.Text
        SystemLocale = $comboSystemLocale.Text
        UserLocale = $comboUserLocale.Text
        TimeZone = $selectedTimeZoneId
        KeyboardLayouts = $selectedKeyboardLayouts -join ';'
        HideEula = $checkHideEula.Checked
        HideLocalAccount = $checkHideLocalAccount.Checked
        HideOnlineAccount = $checkHideOnlineAccount.Checked
        HideWireless = $checkHideWireless.Checked
        ShowFileExt = $checkShowFileExt.Checked
        DisableSmartScreen = $checkDisableSmartScreen.Checked
        DisableSysRestore = $checkDisableSysRestore.Checked
        DisableSuggestions = $checkDisableSuggestions.Checked
        BloatwareToRemove = ($bloatwareCheckboxes | Where-Object { $_.Checked } | ForEach-Object { $_.Text })
    }

    # Final validation (redundant but safe)
    if ([string]::IsNullOrWhiteSpace($formData.ComputerName) -or [string]::IsNullOrWhiteSpace($formData.UserName) -or `
        [string]::IsNullOrWhiteSpace($formData.UiLanguage) -or [string]::IsNullOrWhiteSpace($formData.SystemLocale) -or `
        [string]::IsNullOrWhiteSpace($formData.UserLocale) -or [string]::IsNullOrWhiteSpace($formData.TimeZone) -or `
        $selectedKeyboardLayouts.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please fill in all general and regional settings, including at least one keyboard layout.", "Validation Failed", "OK", "Error") | Out-Null
        Write-LogAndHost "XML creation aborted due to missing required fields." -HostColor Red -LogPrefix "Create-UnattendXml"
        Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
        $null = Read-Host
        return
    }

    # Create the XML file
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $filePath = Join-Path -Path $desktopPath -ChildPath "autounattend.xml"
    Write-LogAndHost "Creating XML structure based on GUI selections..." -NoHost
        
    $xml = New-Object System.Xml.XmlDocument
    $xml.AppendChild($xml.CreateXmlDeclaration("1.0", "utf-8", $null)) | Out-Null
    $root = $xml.CreateElement("unattend")
    $root.SetAttribute("xmlns", "urn:schemas-microsoft-com:unattend")
    $xml.AppendChild($root) | Out-Null
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("d6p1", "http://schemas.microsoft.com/WMIConfig/2002/State")
    
    # Helper function to create components
    function New-Component($ParentNode, $Name, $Pass, $Token="31bf3856ad364e35", $Arch="amd64") {
        $settings = $ParentNode.SelectSingleNode("//unattend/settings[@pass='$Pass']")
        if (-not $settings) { 
            $settings = $ParentNode.OwnerDocument.CreateElement("settings")
            $settings.SetAttribute("pass", $Pass)
            $ParentNode.AppendChild($settings) | Out-Null
        }
        
        $component = $settings.SelectSingleNode("component[@name='$Name']")
        if (-not $component) {
            $component = $ParentNode.OwnerDocument.CreateElement("component")
            $component.SetAttribute("name", $Name)
            $component.SetAttribute("processorArchitecture", $Arch)
            $component.SetAttribute("publicKeyToken", $Token)
            $component.SetAttribute("language", "neutral")
            $component.SetAttribute("versionScope", "nonSxS")
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
    $localAccount.SetAttribute("action", $ns.LookupNamespace("d6p1"), "add")
    $localAccount.AppendChild($xml.CreateElement("Name")).InnerText = $formData.UserName
    $localAccount.AppendChild($xml.CreateElement("Group")).InnerText = "Administrators"
    $localAccount.AppendChild($xml.CreateElement("DisplayName")).InnerText = $formData.UserName
    
    # Only add password element if a password is provided
    if (-not [string]::IsNullOrWhiteSpace($formData.Password)) {
        $passwordNode = $localAccount.AppendChild($xml.CreateElement("Password"))
        $passwordNode.AppendChild($xml.CreateElement("Value")).InnerText = $formData.Password
        $passwordNode.AppendChild($xml.CreateElement("PlainText")).InnerText = "true"
    }
    
    # First Logon Commands
    $firstLogonCommands = $compShellOobe.AppendChild($xml.CreateElement("FirstLogonCommands"))
    $commandIndex = 1
    
    # Add system tweak commands
    if ($formData.ShowFileExt) { 
        $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand"))
        $syncCmd.SetAttribute("Order", $commandIndex++)
        $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = 'cmd /c reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v HideFileExt /t REG_DWORD /d 0 /f'
    }
    
    if ($formData.DisableSmartScreen) { 
        $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand"))
        $syncCmd.SetAttribute("Order", $commandIndex++)
        $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = 'cmd /c reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" /v SmartScreenEnabled /t REG_SZ /d "Off" /f'
    }
    
    if ($formData.DisableSysRestore) { 
        $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand"))
        $syncCmd.SetAttribute("Order", $commandIndex++)
        $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = 'powershell.exe -Command "Disable-ComputerRestore -Drive C:\"'
    }
    
    if ($formData.DisableSuggestions) { 
        $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand"))
        $syncCmd.SetAttribute("Order", $commandIndex++)
        $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = 'cmd /c reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v SubscribedContent-338389Enabled /t REG_DWORD /d 0 /f'
    }

    # Bloatware removal commands
    $bloatwareCommands = @{
        '3D Viewer' = 'Get-AppxPackage *Microsoft.Microsoft3DViewer* | Remove-AppxPackage -AllUsers'
        'Bing Search' = 'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Search" /v BingSearchEnabled /t REG_DWORD /d 0 /f'
        'Calculator' = 'Get-AppxPackage *Microsoft.WindowsCalculator* | Remove-AppxPackage -AllUsers'
        'Camera' = 'Get-AppxPackage *Microsoft.WindowsCamera* | Remove-AppxPackage -AllUsers'
        'Clipchamp' = 'Get-AppxPackage *Microsoft.Clipchamp* | Remove-AppxPackage -AllUsers'
        'Clock' = 'Get-AppxPackage *Microsoft.WindowsAlarms* | Remove-AppxPackage -AllUsers'
        'Copilot' = 'reg add "HKCU\Software\Policies\Microsoft\Windows\WindowsCopilot" /v TurnOffWindowsCopilot /t REG_DWORD /d 1 /f'
        'Cortana' = 'Get-AppxPackage *Microsoft.549981C3F5F10* | Remove-AppxPackage -AllUsers'
        'Dev Home' = 'Get-AppxPackage *Microsoft.DevHome* | Remove-AppxPackage -AllUsers'
        'Family' = 'Get-AppxPackage *Microsoft.Windows.Family* | Remove-AppxPackage -AllUsers'
        'Feedback Hub' = 'Get-AppxPackage *Microsoft.WindowsFeedbackHub* | Remove-AppxPackage -AllUsers'
        'Get Help' = 'Get-AppxPackage *Microsoft.GetHelp* | Remove-AppxPackage -AllUsers'
        'Handwriting (all languages)' = 'Get-WindowsCapability -Online | Where-Object { $_.Name -like "Language.Handwriting*" } | ForEach-Object { Remove-WindowsCapability -Online -Name $_.Name -NoRestart }'
        'Internet Explorer' = 'Disable-WindowsOptionalFeature -Online -FeatureName "Internet-Explorer-Optional-amd64" -NoRestart'
        'Mail and Calendar' = 'Get-AppxPackage *microsoft.windowscommunicationsapps* | Remove-AppxPackage -AllUsers'
        'Maps' = 'Get-AppxPackage *Microsoft.WindowsMaps* | Remove-AppxPackage -AllUsers'
        'Math Input Panel' = 'Remove-WindowsCapability -Online -Name "MathRecognizer~~~~0.0.1.0" -NoRestart'
        'Media Features' = 'Disable-WindowsOptionalFeature -Online -FeatureName "MediaPlayback" -NoRestart'
        'Mixed Reality' = 'Get-AppxPackage *Microsoft.MixedReality.Portal* | Remove-AppxPackage -AllUsers'
        'Movies & TV' = 'Get-AppxPackage *Microsoft.ZuneVideo* | Remove-AppxPackage -AllUsers'
        'News' = 'Get-AppxPackage *Microsoft.BingNews* | Remove-AppxPackage -AllUsers'
        'Notepad (modern)' = 'Get-AppxPackage *Microsoft.WindowsNotepad* | Remove-AppxPackage -AllUsers'
        'Office 365' = 'Get-AppxPackage *Microsoft.MicrosoftOfficeHub* | Remove-AppxPackage -AllUsers'
        'OneDrive' = '$process = Start-Process "$env:SystemRoot\SysWOW64\OneDriveSetup.exe" -ArgumentList "/uninstall" -PassThru -Wait; if ($process.ExitCode -ne 0) { Start-Process "$env:SystemRoot\System32\OneDriveSetup.exe" -ArgumentList "/uninstall" -PassThru -Wait }'
        'OneNote' = 'Get-AppxPackage *Microsoft.Office.OneNote* | Remove-AppxPackage -AllUsers'
        'OneSync' = '# Handled by Mail and Calendar'
        'OpenSSH Client' = 'Remove-WindowsCapability -Online -Name "OpenSSH.Client~~~~0.0.1.0" -NoRestart'
        'Outlook for Windows' = 'Get-AppxPackage *Microsoft.OutlookForWindows* | Remove-AppxPackage -AllUsers'
        'Paint' = 'Get-AppxPackage *Microsoft.Paint* | Remove-AppxPackage -AllUsers'
        'Paint 3D' = 'Get-AppxPackage *Microsoft.MSPaint* | Remove-AppxPackage -AllUsers'
        'People' = 'Get-AppxPackage *Microsoft.People* | Remove-AppxPackage -AllUsers'
        'Photos' = 'Get-AppxPackage *Microsoft.Windows.Photos* | Remove-AppxPackage -AllUsers'
        'Power Automate' = 'Get-AppxPackage *Microsoft.PowerAutomateDesktop* | Remove-AppxPackage -AllUsers'
        'PowerShell 2.0' = 'Disable-WindowsOptionalFeature -Online -FeatureName "MicrosoftWindowsPowerShellV2" -NoRestart'
        'PowerShell ISE' = 'Remove-WindowsCapability -Online -Name "PowerShell-ISE-v2~~~~0.0.1.0" -NoRestart'
        'Quick Assist' = 'Get-AppxPackage *Microsoft.QuickAssist* | Remove-AppxPackage -AllUsers'
        'Recall' = 'reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" /v DisableAllScreenshotCapture /t REG_DWORD /d 1 /f'
        'Remote Desktop Client' = '# Core component, removal not recommended.'
        'Skype' = 'Get-AppxPackage *Microsoft.SkypeApp* | Remove-AppxPackage -AllUsers'
        'Snipping Tool' = 'Get-AppxPackage *Microsoft.ScreenSketch* | Remove-AppxPackage -AllUsers'
        'Solitaire Collection' = 'Get-AppxPackage *Microsoft.MicrosoftSolitaireCollection* | Remove-AppxPackage -AllUsers'
        'Speech (all languages)' = 'Get-WindowsCapability -Online | Where-Object { $_.Name -like "Language.Speech*" } | ForEach-Object { Remove-WindowsCapability -Online -Name $_.Name -NoRestart }'
        'Steps Recorder' = 'Disable-WindowsOptionalFeature -Online -FeatureName "StepsRecorder" -NoRestart'
        'Sticky Notes' = 'Get-AppxPackage *Microsoft.MicrosoftStickyNotes* | Remove-AppxPackage -AllUsers'
        'Teams' = 'Get-AppxPackage *MicrosoftTeams* | Remove-AppxPackage -AllUsers'
        'Tips' = 'Get-AppxPackage *Microsoft.Getstarted* | Remove-AppxPackage -AllUsers'
        'To Do' = 'Get-AppxPackage *Microsoft.Todos* | Remove-AppxPackage -AllUsers'
        'Voice Recorder' = 'Get-AppxPackage *Microsoft.WindowsSoundRecorder* | Remove-AppxPackage -AllUsers'
        'Wallet' = 'Get-AppxPackage *Microsoft.Wallet* | Remove-AppxPackage -AllUsers'
        'Weather' = 'Get-AppxPackage *Microsoft.BingWeather* | Remove-AppxPackage -AllUsers'
        'Windows Fax and Scan' = 'Disable-WindowsOptionalFeature -Online -FeatureName "Windows-Fax-And-Scan" -NoRestart'
        'Windows Hello' = 'reg add "HKLM\SOFTWARE\Policies\Microsoft\Biometrics" /v Enabled /t REG_DWORD /d 0 /f; reg add "HKLM\SOFTWARE\Policies\Microsoft\Biometrics\CredentialProviders" /v Enabled /t REG_DWORD /d 0 /f'
        'Windows Media Player (classic)' = 'Disable-WindowsOptionalFeature -Online -FeatureName "WindowsMediaPlayer" -NoRestart'
        'Windows Media Player (modern)' = 'Get-AppxPackage *Microsoft.ZuneMusic* | Remove-AppxPackage -AllUsers'
        'Windows Terminal' = 'Get-AppxPackage *Microsoft.WindowsTerminal* | Remove-AppxPackage -AllUsers'
        'WordPad' = 'Remove-WindowsCapability -Online -Name "WordPad~~~~0.0.1.0" -NoRestart'
        'Xbox Apps' = 'Get-AppxPackage *Microsoft.Xbox* | Remove-AppxPackage -AllUsers; Get-AppxPackage *Microsoft.GamingApp* | Remove-AppxPackage -AllUsers'
        'Your Phone / Phone Link' = 'Get-AppxPackage *Microsoft.YourPhone* | Remove-AppxPackage -AllUsers'
    }

    # Add bloatware removal commands
    foreach ($bloat in $formData.BloatwareToRemove) {
        if ($bloatwareCommands.ContainsKey($bloat)) {
            $command = $bloatwareCommands[$bloat]
            if ($command.StartsWith("#")) { continue }
            
            if ($command -match 'Get-AppxPackage|Remove-AppxPackage|Get-WindowsCapability|Remove-WindowsCapability|Disable-WindowsOptionalFeature|Start-Process') {
                $bytes = [System.Text.Encoding]::Unicode.GetBytes($command)
                $encodedCommand = [Convert]::ToBase64String($bytes)
                $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand"))
                $syncCmd.SetAttribute("Order", $commandIndex++)
                $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = "powershell.exe -EncodedCommand $encodedCommand"
            } elseif ($command -match 'reg add|reg delete') {
                 $syncCmd = $firstLogonCommands.AppendChild($xml.CreateElement("SynchronousCommand"))
                 $syncCmd.SetAttribute("Order", $commandIndex++)
                 $syncCmd.AppendChild($xml.CreateElement("CommandLine")).InnerText = "cmd /c $command"
            }
        }
    }

    # Save the XML file
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
    $menuLines.Add(" Windows & Software Manager [PSS v1.7] ($(Get-Date -Format "dd.MM.yyyy HH:mm"))") 
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
    @{ Left = "[U] Uninstall Programs [GUI]";          Right = "[S] System Cleanup [GUI]" },
    @{ Left = "[C] Install Custom Program";            Right = "[F] Create Unattend.xml File [GUI]" },
    @{ Left = "[X] Activate Spotify";                  Right = "[I] Show System Information" },
    @{ Left = "[L] Import & Install from File";        Right = "[P] Perdanga System Manager" }                                     
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
# In PART 3: CONFIGURATION
$script:mainMenuLetters = @('a', 'c', 'e', 'f', 'g', 'i', 'n', 'p', 'r', 's', 't', 'u', 'w', 'x', 'l') 
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
            'l' { Clear-Host; Import-ProgramSelection }
            'r' { Clear-Host; Invoke-SystemRepair }
            'p' { Clear-Host; Invoke-PerdangaSystemManager }
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



