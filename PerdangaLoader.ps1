<#
================================================================================
================================================================================
                                 PART 1: FUNCTIONS
================================================================================
================================================================================
#>

<#
.SYNOPSIS
    Author: Roman Zhdanov
    Version: 1.4 
.DESCRIPTION
    This file contains all the helper functions used by the main script.
    It is dot-sourced to make these functions available in the main script's scope.
.NOTES
    This script should not be run directly.
#>

# Set log file name with timestamp and use script directory
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
# When run from `irm | iex`, $PSScriptRoot is not available. Default to a temp path.
$scriptDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($scriptDir)) {
    $scriptDir = $env:TEMP
}
$script:logFile = Join-Path -Path $scriptDir -ChildPath "install_log_$timestamp.txt"


# Function to write messages to console and log file
function Write-LogAndHost {
    param (
        [string]$Message,
        [string]$LogPrefix = "",
        [string]$HostColor = "White",
        [switch]$NoLog,
        [switch]$NoHost,
        [switch]$NoNewline
    )
    $fullLogMessage = "[$((Get-Date))] $LogPrefix$Message"
    if (-not $NoLog) {
        # This line will now work correctly from the beginning
        try {
            $fullLogMessage | Out-File -FilePath $script:logFile -Append -Encoding UTF8 -ErrorAction Stop
        }
        catch {
            Write-Host "FATAL: Could not write to log file at $($script:logFile). Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    if (-not $NoHost) {
        if ($NoNewline) {
            Write-Host $Message -ForegroundColor $HostColor -NoNewline
        } else {
            Write-Host $Message -ForegroundColor $HostColor
        }
    }
}

# Function to install Chocolatey
function Install-Chocolatey {
    try {
        # ENHANCEMENT: Log the user prompt for better traceability.
        Write-LogAndHost "Chocolatey is not installed. Would you like to install it? (Type y/n then press Enter)" -HostColor Yellow
        $confirmInput = Read-Host
        if ($confirmInput.Trim().ToLower() -eq 'y') {
            Write-LogAndHost "User chose to install Chocolatey." -NoHost
            try {
                Write-LogAndHost "Installing Chocolatey..." -NoLog
                Set-ExecutionPolicy Bypass -Scope Process -Force
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                $installOutput = Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1')) 2>&1
                $installOutput | Out-File -FilePath $script:logFile -Append -Encoding UTF8
                if ($LASTEXITCODE -eq 0) {
                    Write-LogAndHost "Chocolatey installed successfully." -HostColor Green
                    # Refresh environment variables to ensure choco is available
                    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
                    return $true
                } else {
                    Write-LogAndHost "ERROR: Failed to install Chocolatey. Exit code: $LASTEXITCODE. Details: $($installOutput | Out-String)" -HostColor Red -LogPrefix "Error installing Chocolatey. "
                    return $false
                }
            } catch {
                Write-LogAndHost "ERROR: Exception occurred while installing Chocolatey - $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during Chocolatey installation - "
                return $false
            }
        } else {
            Write-LogAndHost "Chocolatey installation cancelled by user." -HostColor Yellow
            return $false
        }
    } catch {
        Write-LogAndHost "ERROR: Could not read user input for Chocolatey installation. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Failed to read user input for Chocolatey installation - "
        return $false
    }
}

# Function to perform Windows activation
function Invoke-WindowsActivation {
    $script:activationAttempted = $true
    # ENHANCEMENT: Warnings should be logged for audit purposes.
    Write-LogAndHost "WARNING: Windows activation uses an external script from 'https://get.activated.win'. Ensure you trust the source before proceeding." -HostColor Yellow
    try {
        # ENHANCEMENT: Log the user prompt.
        Write-LogAndHost "Continue with Windows activation? (Type y/n then press Enter)" -HostColor Yellow
        $confirmActivation = Read-Host
        if ($confirmActivation.Trim().ToLower() -ne 'y') {
            Write-LogAndHost "Windows activation cancelled by user." -HostColor Yellow
            return
        }
    } catch {
        Write-LogAndHost "ERROR: Could not read user input for Windows activation. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Failed to read user input for Windows activation - "
        return
    }
    Write-LogAndHost "Attempting Windows activation..." -NoHost
    try {
        # Ensure TLS 1.2 for the activation script download
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        $activationScriptContent = Invoke-RestMethod -Uri "https://get.activated.win" -UseBasicParsing
        Invoke-Expression -Command $activationScriptContent 2>&1 | Tee-Object -FilePath $script:logFile -Append | ForEach-Object { Write-Host $_ -ForegroundColor Gray }
        
        if ($LASTEXITCODE -eq 0) { # This might not be reliable for iex from irm
            Write-LogAndHost "Windows activation script executed. Check console output above for status."
        }
        else {
            Write-LogAndHost "Windows activation script execution might have failed. Exit code: $LASTEXITCODE" -HostColor Yellow
        }
    }
    catch {
        Write-LogAndHost "ERROR: Exception during Windows activation - $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during Windows activation - "
    }
    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}

# Function to apply SpotX
function Invoke-SpotXActivation {
    Write-LogAndHost "Attempting Spotify Activation..." -HostColor Cyan
    Write-LogAndHost "INFO: This process modifies your Spotify client. It is recommended to close Spotify before proceeding." -HostColor Yellow
    Write-LogAndHost "WARNING: This script downloads and executes code from the internet (SpotX-Official GitHub). Ensure you trust the source." -HostColor Yellow

    try {
        # ENHANCEMENT: Log the user prompt.
        Write-LogAndHost "Continue with Spotify Activation? (Type y/n then press Enter)" -HostColor Yellow
        $confirmSpotX = Read-Host
        if ($confirmSpotX.Trim().ToLower() -ne 'y') {
            Write-LogAndHost "Spotify Activation cancelled by user." -HostColor Yellow
            return
        }
    } catch {
        Write-LogAndHost "ERROR: Could not read user input for Spotify Activation confirmation. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Failed to read user input for Spotify Activation confirmation - "
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
        Write-LogAndHost "Primary SpotX URL failed. Error: $($_.Exception.Message)" -HostColor DarkYellow -LogPrefix "SpotX Primary Download Failed: "
        Write-LogAndHost "Attempting fallback URL: $spotxUrlFallback" -NoHost
        try {
            $scriptContentToExecute = (Invoke-WebRequest -UseBasicParsing -Uri $spotxUrlFallback -ErrorAction Stop).Content
            Write-LogAndHost "Successfully downloaded from fallback URL." -NoHost
        } catch {
            Write-LogAndHost "ERROR: Fallback SpotX URL also failed. Error: $($_.Exception.Message)" -HostColor Red -LogPrefix "SpotX Fallback Download Failed: "
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
            Write-LogAndHost "ERROR: Exception occurred during SpotX script execution. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during SpotX execution - "
            $_ | Out-File -FilePath $script:logFile -Append -Encoding UTF8
        }
    } else {
        Write-LogAndHost "ERROR: Failed to obtain SpotX script content." -HostColor Red
    }

    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}

# Function to perform Windows update
function Invoke-WindowsUpdate {
    Write-LogAndHost "Checking for Windows updates..."
    Set-ExecutionPolicy Bypass -Scope Process -Force
    try {
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-LogAndHost "PSWindowsUpdate module not found. Installing..."
            Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction Stop
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
        Write-LogAndHost "ERROR: Failed to update Windows. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Error during Windows update: "
    }
    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}

# Function to disable Windows Telemetry
function Invoke-DisableTelemetry {
    Write-LogAndHost "Checking Windows Telemetry status..." -HostColor Cyan
    
    $telemetryService = Get-Service -Name "DiagTrack" -ErrorAction SilentlyContinue
    $telemetryRegValue = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -ErrorAction SilentlyContinue

    if ($telemetryService -and $telemetryService.StartType -eq 'Disabled' -and $telemetryRegValue -and $telemetryRegValue.AllowTelemetry -eq 0) {
        Write-LogAndHost "Windows Telemetry appears to be already disabled." -HostColor Green
        Write-LogAndHost "No changes were made." -HostColor Green
        Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
        $null = Read-Host
        return
    }
    
    try {
        # ENHANCEMENT: Log the user prompt.
        Write-LogAndHost "Telemetry is currently enabled. Continue with disabling? (Type y/n then press Enter)" -HostColor Yellow
        $confirmTelemetry = Read-Host
        if ($confirmTelemetry.Trim().ToLower() -ne 'y') {
            Write-LogAndHost "Telemetry disabling cancelled by user." -HostColor Yellow
            return
        }
    } catch {
        Write-LogAndHost "ERROR: Could not read user input for Telemetry confirmation. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Failed to read user input for Telemetry confirmation - "
        return
    }

    Write-LogAndHost "Applying telemetry settings..." -NoHost
    
    try {
        $servicesToDisable = @("DiagTrack", "dmwappushservice")
        foreach ($serviceName in $servicesToDisable) {
            $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
            if ($service) {
                try {
                    Write-LogAndHost "Stopping service: $serviceName..." -NoLog
                    Stop-Service -Name $serviceName -Force -ErrorAction Stop
                    Write-LogAndHost "Disabling service: $serviceName..." -NoLog
                    Set-Service -Name $serviceName -StartupType Disabled -ErrorAction Stop
                    Write-LogAndHost "$serviceName service stopped and disabled." -HostColor Green
                } catch {
                    Write-LogAndHost "ERROR: Could not stop or disable service '$serviceName'. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Telemetry Service Error: "
                }
            } else {
                Write-LogAndHost "Service '$serviceName' not found, skipping." -HostColor DarkGray -NoLog
            }
        }

        Write-LogAndHost "Configuring registry keys..." -NoLog
        
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
                if (-not (Test-Path $path)) {
                    Write-LogAndHost "Creating registry path: $path" -NoLog
                    New-Item -Path $path -Force -ErrorAction Stop | Out-Null
                }
                Set-ItemProperty -Path $path -Name $keyInfo.Name -Value $keyInfo.Value -Type $keyInfo.Type -Force -ErrorAction Stop
                Write-LogAndHost "Successfully set registry value '$($keyInfo.Name)' at '$path'." -NoLog
            } catch {
                Write-LogAndHost "ERROR: Failed to set registry key at '$path'. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Telemetry Registry Error: "
            }
        }
        
        Write-LogAndHost "Telemetry has been successfully disabled." -HostColor Green
    } catch {
        Write-LogAndHost "ERROR: An unexpected error occurred while disabling telemetry. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Telemetry Global Error: "
    }

    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}

# Function to install programs using Chocolatey
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
        Write-LogAndHost "Installing $program..." -LogPrefix "Installing $program (from list)..."
        try {
            $installOutput = & choco install $program -y --source=$Source --no-progress 2>&1
            $installOutput | Out-File -FilePath $script:logFile -Append -Encoding UTF8

            if ($LASTEXITCODE -eq 0) {
                # ENHANCEMENT: More robust check for already installed packages
                if ($installOutput -match "is already installed|already installed|Nothing to do") {
                    Write-LogAndHost "$program is already installed or up to date." -HostColor Green
                } else {
                    Write-LogAndHost "$program installed successfully." -HostColor White
                }
            } else {
                $allSuccess = $false
                Write-LogAndHost "ERROR: Failed to install $program. Exit code: $LASTEXITCODE. Details: $($installOutput | Out-String)" -HostColor Red -LogPrefix "Error installing $program. "
            }
        } catch {
            $allSuccess = $false
            Write-LogAndHost "ERROR: Exception occurred while installing $program. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during installation of $program - "
        }
        Write-Host ""
    }
    return $allSuccess
}

# Function to get installed Chocolatey packages
function Get-InstalledChocolateyPackages {
    $chocoLibPath = Join-Path -Path $env:ChocolateyInstall -ChildPath "lib"
    $installedPackages = @()
    if (Test-Path $chocoLibPath) {
        try {
            $installedPackages = Get-ChildItem -Path $chocoLibPath -Directory | Select-Object -ExpandProperty Name
            Write-LogAndHost "Found installed packages: $($installedPackages -join ', ')" -NoHost
        } catch {
            Write-LogAndHost "ERROR: Could not retrieve installed packages from $chocoLibPath. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Failed to get installed packages - "
        }
    } else {
        Write-LogAndHost "WARNING: Chocolatey lib directory not found at $chocoLibPath. Cannot list installed packages." -HostColor Yellow
    }
    return $installedPackages
}

# Function to uninstall programs using Chocolatey
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
        Write-LogAndHost "Uninstalling $program..." -LogPrefix "Uninstalling $program..."
        try {
            $uninstallOutput = & choco uninstall $program -y --no-progress 2>&1
            $uninstallOutput | Out-File -FilePath $script:logFile -Append -Encoding UTF8

            if ($LASTEXITCODE -eq 0) {
                Write-LogAndHost "$program uninstalled successfully." -HostColor White
            } else {
                $allSuccess = $false
                Write-LogAndHost "ERROR: Failed to uninstall $program. Exit code: $LASTEXITCODE. Details: $($uninstallOutput | Out-String)" -HostColor Red -LogPrefix "Error uninstalling $program. "
            }
        } catch {
            $allSuccess = $false
            Write-LogAndHost "ERROR: Exception occurred while uninstalling $program. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during uninstallation of $program - "
        }
        Write-Host ""
    }
    return $allSuccess
}

# ENHANCED FUNCTION: Create a detailed autounattend.xml file via GUI with regional settings and tooltips
function Create-UnattendXml {
    if (-not $script:guiAvailable) {
        Write-LogAndHost "ERROR: GUI is not available, cannot launch the Unattend XML Creator." -HostColor Red
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
    $toolTip.AutoPopDelay = 10000 # Keep tooltip visible for 10 seconds
    $toolTip.InitialDelay = 500   # Show after 0.5 seconds
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

    # Common Font and Color
    $commonFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $labelColor = [System.Drawing.Color]::White
    $controlBackColor = [System.Drawing.Color]::FromArgb(60, 60, 63)
    $controlForeColor = [System.Drawing.Color]::White
    $groupboxForeColor = [System.Drawing.Color]::Gainsboro

    # Helper functions for creating styled controls
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
    
    # --- ENHANCEMENT: Add password length limit and character counter ---
    $tabGeneral.Controls.Add((New-StyledLabel -Text "Password:" -Location "20,$yPos")) | Out-Null
    $textPassword = New-StyledTextBox -Location "180,$yPos" -Size "280,20"
    $textPassword.UseSystemPasswordChar = $true
    $textPassword.MaxLength = 127 # Enforce Windows unattend spec 127-char limit
    $tabGeneral.Controls.Add($textPassword) | Out-Null
    # Add a character counter label for user feedback
    $labelPasswordCounter = New-StyledLabel -Text "0/127" -Location "470,$yPos"
    $tabGeneral.Controls.Add($labelPasswordCounter) | Out-Null
    # Add event handler to update the counter in real-time
    $textPassword.Add_TextChanged({
        $length = $textPassword.Text.Length
        $labelPasswordCounter.Text = "$length/127"
        if ($length -eq 127) {
            $labelPasswordCounter.ForeColor = [System.Drawing.Color]::Crimson
        } else {
            $labelPasswordCounter.ForeColor = $labelColor # Revert to default color
        }
    })
    $yPos += 40

    # --- Tab 2: Regional Settings ---
    $tabRegional = New-Object System.Windows.Forms.TabPage; $tabRegional.Text = "Regional"; $tabRegional.BackColor = $form.BackColor; $tabRegional.Padding = New-Object System.Windows.Forms.Padding(10)
    $tabControl.Controls.Add($tabRegional) | Out-Null

    # GroupBox for Language and Locale
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

    # GroupBox for Time Zone
    $groupTimeZone = New-StyledGroupBox "Time Zone" "15,180" "750,220"
    $tabRegional.Controls.Add($groupTimeZone) | Out-Null
    $yPos = 30
    $groupTimeZone.Controls.Add((New-StyledLabel -Text "Search:" -Location "15,$yPos")) | Out-Null; $textTimeZoneSearch = New-StyledTextBox -Location "85,$yPos" -Size "645,20"; $groupTimeZone.Controls.Add($textTimeZoneSearch) | Out-Null; $yPos += 35
    $listTimeZone = New-Object System.Windows.Forms.ListBox; $listTimeZone.Location = "15,$yPos"; $listTimeZone.Size = "715,100"; $listTimeZone.Font = $commonFont; $listTimeZone.BackColor = $controlBackColor; $listTimeZone.ForeColor = $controlForeColor
    
    # --- ENHANCEMENT: Use a static list of Windows 11 Time Zone IDs ---
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
        Write-LogAndHost "WARNING: Could not find all static time zones. The list may be incomplete. Falling back to system's available time zones." -HostColor Yellow
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
    $sortedFormattedTimeZones = $formattedTimeZones | Sort-Object
    
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

    # --- ENHANCEMENT: Adjusted GroupBox title and height for new constraints ---
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

    # Label to show current Keyboard Layout selection(s)
    $groupKeyboard.Controls.Add((New-StyledLabel -Text "Current Selection:" -Location "15,$yPos")) | Out-Null
    # Increased label height to allow text wrapping for multiple selections
    $labelSelectedKeyboards = New-StyledLabel -Text "None" -Location "150,$yPos"; $labelSelectedKeyboards.ForeColor = [System.Drawing.Color]::LightSteelBlue; $labelSelectedKeyboards.AutoSize = $false; $labelSelectedKeyboards.Size = '580,50'
    $groupKeyboard.Controls.Add($labelSelectedKeyboards) | Out-Null

    # Script block to update the keyboard label based on the persistent list
    $updateKeyboardLabel = {
        $checkedItemsText = ($checkedKeyboardLayoutNames | Sort-Object) -join ', '
        if ([string]::IsNullOrWhiteSpace($checkedItemsText)) { $labelSelectedKeyboards.Text = "None" } else { $labelSelectedKeyboards.Text = $checkedItemsText }
    }

    # Event Handlers for Keyboard list
    $listKeyboardLayouts.Add_ItemCheck({
        param($sender, $e)
        $itemName = $sender.Items[$e.Index]
        if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
            # --- ENHANCEMENT: Limit selection to 5 keyboard layouts ---
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

    # --- Tab 3: Automation & Tweaks (COMBINED TAB) ---
    $tabAutomation = New-Object System.Windows.Forms.TabPage; $tabAutomation.Text = "Automation & Tweaks"; $tabAutomation.BackColor = $form.BackColor; $tabAutomation.Padding = New-Object System.Windows.Forms.Padding(10)
    $tabControl.Controls.Add($tabAutomation) | Out-Null

    # GroupBox for OOBE Automation
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

    # GroupBox for Customization/System Tweaks
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

    # --- NEW: Add instructional caption at the bottom of the tab ---
    $automationInfoLabel = New-Object System.Windows.Forms.Label
    $automationInfoLabel.Text = "Hover over an option for a detailed description ."
    $automationInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $automationInfoLabel.ForeColor = [System.Drawing.Color]::Gray
    $automationInfoLabel.AutoSize = $true
    # Position it below the group boxes. The second groupbox ends at y = 250 + 220 = 470.
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

    # Show the form and check the result
    try {
        $result = $form.ShowDialog()
    }
    catch {
        Write-LogAndHost "ERROR: An unexpected error occurred with the Unattend XML Creator GUI. Details: $($_.Exception.Message)" -HostColor Red
        $result = [System.Windows.Forms.DialogResult]::Cancel
    }
    finally {
        $form.Dispose()
    }
    
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-LogAndHost "XML creation cancelled by user." -HostColor Yellow; Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host; return
    }

    # Collect data from form
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
        Write-LogAndHost "XML creation aborted due to missing required fields." -HostColor Red
        Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host; return
    }

    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $filePath = Join-Path -Path $desktopPath -ChildPath "autounattend.xml"
    Write-LogAndHost "Creating XML structure based on GUI selections..." -NoHost
        
    # Create XML Document
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
    
    # --- Build XML from formData ---
    # Pass: specialize
    $compIntlSpec = New-Component -ParentNode $root -Name "Microsoft-Windows-International-Core" -Pass "specialize"
    $compIntlSpec.AppendChild($xml.CreateElement("InputLocale")).InnerText = $formData.KeyboardLayouts
    $compIntlSpec.AppendChild($xml.CreateElement("SystemLocale")).InnerText = $formData.SystemLocale
    $compIntlSpec.AppendChild($xml.CreateElement("UILanguage")).InnerText = $formData.UiLanguage
    $compIntlSpec.AppendChild($xml.CreateElement("UserLocale")).InnerText = $formData.UserLocale

    $compShellSpec = New-Component -ParentNode $root -Name "Microsoft-Windows-Shell-Setup" -Pass "specialize"
    $compShellSpec.AppendChild($xml.CreateElement("ComputerName")).InnerText = $formData.ComputerName
    $compShellSpec.AppendChild($xml.CreateElement("TimeZone")).InnerText = $formData.TimeZone
    
    # Pass: oobeSystem
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
    
    # FirstLogonCommands are executed after user logon
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
        Write-LogAndHost "ERROR: Failed to save the XML file to '$filePath'. Details: $($_.Exception.Message)" -HostColor Red
    }
    
    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}


# Function to test if a Chocolatey package exists
function Test-ChocolateyPackage {
    param (
        [string]$PackageName
    )
    Write-LogAndHost "Searching for package '$PackageName' in Chocolatey repository..." -NoLog
    try {
        $searchOutput = & choco search $PackageName --exact --limit-output --source="https://community.chocolatey.org/api/v2/" --no-progress 2>&1
        $searchOutput | Out-File -FilePath $script:logFile -Append -Encoding UTF8

        if ($LASTEXITCODE -ne 0) {
             Write-LogAndHost "Error during 'choco search' for '$PackageName'. Exit code: $LASTEXITCODE. Output: $($searchOutput | Out-String)" -HostColor Red
             return $false
        }
        
        if ($searchOutput -match "$([regex]::Escape($PackageName))\|.*" -or $searchOutput -match "1 packages found.") {
             Write-LogAndHost "Package '$PackageName' found in repository. Search output: $($searchOutput | Select-Object -First 1)" -HostColor Green
             return $true
        } else {
             Write-LogAndHost "Package '$PackageName' not found as an exact match in Chocolatey repository. Search output: $($searchOutput | Out-String)" -HostColor Yellow
             return $false
        }

    } catch {
        Write-LogAndHost "ERROR: Exception occurred while searching for package '$PackageName'. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during package search for '$PackageName' - "
        return $false
    }
}

# Function to display the selection menu
function Show-Menu {
    Clear-Host

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
        "                _\/\\\__________\///\\\\\/___\/\\\__________\//\\\\\\\\\\____\//\\\_____\//\\\\\\\\\\_ \/\\\___________________________",
        "                 _\///_____________\/////_____\///____________\//////////______\///______\//////////____\///____________________________"
    )

    foreach ($line in $asciiArt) {
        Write-Host $line -ForegroundColor Cyan
    }

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
    $menuLines.Add(" Chocolatey Package Manager [PSS v1.4] ($(Get-Date -Format "dd.MM.yyyy HH:mm"))") 
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
        @{ Left = "[C] Install Custom Program";            Right = "[F] Create Unattend.xml File [GUI]" },
        @{ Left = "[X] Activate Spotify";                  Right = "" }
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

    try {
        $consoleWidth = $Host.UI.RawUI.WindowSize.Width
    }
    catch {
        # Fallback if RawUI is not available
        $consoleWidth = 80
    }
    $blockPaddingValue = [math]::Floor(($consoleWidth - $fixedMenuWidth) / 2)
    if ($blockPaddingValue -lt 0) { $blockPaddingValue = 0 }
    $blockPaddingString = " " * $blockPaddingValue

    foreach ($lineEntry in $menuLines) {
        $trimmedEntry = $lineEntry.Trim()
        if ($trimmedEntry -eq $pssText -or
            $trimmedEntry -like ($pssUnderline.Trim()) -or
            $trimmedEntry -like ($dashedLine.Trim()) -or
            $trimmedEntry -eq $programHeader -or
            $trimmedEntry -eq $optionsHeader -or
            $trimmedEntry -like ($programUnderline.Trim()) -or
            $trimmedEntry -like ($optionsUnderline.Trim())) {
            Write-Host ($blockPaddingString + $lineEntry) -ForegroundColor Cyan
        } else {
            Write-Host ($blockPaddingString + $lineEntry) -ForegroundColor White
        }
    }
    Write-Host "" 

    $promptTextForOneLine = "Enter option, single number, or list of numbers:"
    $promptPaddingOneLine = [math]::Floor(($fixedMenuWidth - $promptTextForOneLine.Length) / 2)
    if ($promptPaddingOneLine -lt 0) { $promptPaddingOneLine = 0 }
    $centeredPromptOneLine = (" " * $promptPaddingOneLine) + $promptTextForOneLine
    Write-Host ($blockPaddingString + $centeredPromptOneLine) -NoNewline -ForegroundColor Yellow
}

# Log that the functions library has been loaded successfully
Write-LogAndHost "Functions library loaded." -NoHost

<#
================================================================================
================================================================================
                               PART 2: CONFIGURATION
================================================================================
================================================================================
#>

<#
.SYNOPSIS
    Configuration file for Perdanga Software Solutions.
.DESCRIPTION
    This file contains all the user-configurable variables for the main script,
    such as the list of programs to install.
.NOTES
    This script is dot-sourced by the main script and runs in its scope.
#>

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
    "obs-studio",
    "occt",
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
$script:mainMenuLetters = @('a', 'c', 'e', 'f', 'g', 'n', 't', 'u', 'w', 'x')
$script:mainMenuRegexPattern = "^[" + ($script:mainMenuLetters -join '') + "0-9\s,]+$"
$script:availableProgramNumbers = 1..($script:sortedPrograms.Count) | ForEach-Object { $_.ToString() }
$script:programToNumberMap = @{}
$script:numberToProgramMap = @{}

if ($script:sortedPrograms.Count -gt $script:availableProgramNumbers.Count) {
    Write-LogAndHost "WARNING: Not enough unique numbers available to assign to all programs." -ForegroundColor Yellow -LogPrefix "CRITICAL WARNING: "
}

for ($i = 0; $i -lt $script:sortedPrograms.Count; $i++) {
    if ($i -lt $script:availableProgramNumbers.Count) {
        $assignedNumber = $script:availableProgramNumbers[$i]
        $programName = $script:sortedPrograms[$i]
        $script:programToNumberMap[$programName] = $assignedNumber
        $script:numberToProgramMap[$assignedNumber] = $programName
    } else {
        Write-LogAndHost "WARNING: Ran out of assignable numbers. Program '$($script:sortedPrograms[$i])' will not be selectable by number." -LogPrefix "WARNING: "
        break
    }
}

Write-LogAndHost "Configuration loaded. $($script:sortedPrograms.Count) programs defined." -NoHost

<#
================================================================================
================================================================================
                                PART 3: MAIN SCRIPT
================================================================================
================================================================================
#>

<#
.SYNOPSIS
    Main script for Perdanga Software Solutions to install and manage programs using Chocolatey.
.DESCRIPTION
    This is the main entry point for the script.
.NOTES
    Author: Roman Zhdanov
    Version: 1.4
#>

# --- INITIAL SETUP ---
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
    Write-LogAndHost "WARNING: Could not set console buffer or window size. Error: $($_.Exception.Message)" -ForegroundColor Yellow
}

# --- SCRIPT-WIDE CHECKS ---
$script:activationAttempted = $false
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-LogAndHost "ERROR: This script must be run as Administrator." -ForegroundColor Red
    "[$((Get-Date))] Error: Script not run as Administrator." | Out-File -FilePath $script:logFile -Append -Encoding UTF8
    exit 1
}

$script:guiAvailable = $true
try { Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop }
catch { $script:guiAvailable = $false; "[$((Get-Date))] Warning: GUI is not available - $($_.Exception.Message)" | Out-File -FilePath $script:logFile -Append -Encoding UTF8 }


# --- CHOCOLATEY INITIALIZATION ---
Write-LogAndHost "Checking Chocolatey installation..."
try {
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        if (-not (Install-Chocolatey)) { Write-LogAndHost "ERROR: Chocolatey is required to proceed. Exiting script." -HostColor Red; exit 1 }
        if (-not (Get-Command choco -ErrorAction SilentlyContinue)) { Write-LogAndHost "ERROR: Chocolatey command not found after installation. Please install manually." -HostColor Red; exit 1 }
        if (-not $env:ChocolateyInstall) { $env:ChocolateyInstall = "$($env:ProgramData)\chocolatey"; Write-LogAndHost "ChocolateyInstall environment variable set to: $env:ChocolateyInstall" -NoHost }
    } else {
         if (-not $env:ChocolateyInstall) { $env:ChocolateyInstall = "$($env:ProgramData)\chocolatey"; Write-LogAndHost "ChocolateyInstall environment variable set to: $env:ChocolateyInstall" -NoHost }
    }
    $chocoVersion = & choco --version 2>&1
    if ($LASTEXITCODE -ne 0) { Write-LogAndHost "ERROR: Chocolatey is not functioning correctly. Exit code: $LASTEXITCODE" -HostColor Red; exit 1 }
    Write-LogAndHost "Found Chocolatey version: $($chocoVersion -join ' ')"
}
catch { Write-LogAndHost "ERROR: Exception occurred while checking Chocolatey. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during Chocolatey check - "; exit 1 }

Write-Host ""; Write-LogAndHost "Clearing Chocolatey cache..."
try {
    & choco cache remove --all 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8 
    if ($LASTEXITCODE -eq 0) { Write-LogAndHost "Cache cleared." } else { Write-LogAndHost "WARNING: Failed to clear Chocolatey cache. $($LASTEXITCODE)" -HostColor Yellow }
} catch { Write-LogAndHost "ERROR: Exception clearing Chocolatey cache. $($_.Exception.Message)" -HostColor Red }

Write-Host ""; Write-LogAndHost "Enabling automatic confirmation..."
try {
    & choco feature enable -n allowGlobalConfirmation 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8
    if ($LASTEXITCODE -eq 0) { Write-LogAndHost "Automatic confirmation enabled." } else { Write-LogAndHost "WARNING: Failed to enable automatic confirmation. $($LASTEXITCODE)" -HostColor Yellow }
} catch { Write-LogAndHost "ERROR: Exception enabling automatic confirmation. $($_.Exception.Message)" -HostColor Red }
Write-Host ""


# --- MAIN LOOP ---
do {
    Show-Menu
    try { $userInput = Read-Host } catch { Write-LogAndHost "ERROR: Could not read user input. $($_.Exception.Message)" -HostColor Red; Start-Sleep -Seconds 2; continue }
    $userInput = $userInput.Trim().ToLower()

    if ([string]::IsNullOrEmpty($userInput)) {
        Clear-Host; Write-LogAndHost "No input detected. Please enter an option." -HostColor Yellow
        Write-LogAndHost "Press any key to return to the menu..." -HostColor DarkGray -NoLog; $null = Read-Host; continue 
    }

    if ($userInput -and $userInput -notmatch $script:mainMenuRegexPattern) { 
        Clear-Host; $validOptions = ($script:mainMenuLetters | Sort-Object | ForEach-Object { $_.ToUpper() }) -join ','
        $errorMessage = "Invalid input: '$userInput'. Use options [$validOptions] or program numbers."
        Write-LogAndHost $errorMessage -HostColor Red -LogPrefix "Invalid user input: '$userInput'."; Start-Sleep -Seconds 2; continue
    }

    if ($script:mainMenuLetters -contains $userInput) {
        switch ($userInput) {
            'e' {
                Write-LogAndHost "Exiting script..."
                try { & choco feature disable -n allowGlobalConfirmation 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8 } catch { Write-LogAndHost "ERROR: Exception disabling auto-confirm - $($_.Exception.Message)" -HostColor Red }
                # The --local-only argument is deprecated and removed in recent Chocolatey versions. This command correctly lists locally installed packages.
                try { & choco list 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8 } catch { Write-LogAndHost "ERROR: Exception listing installed programs - $($_.Exception.Message)" -HostColor Red }
                exit 0
            }
            'a' {
                Clear-Host
                try { 
                    # ENHANCEMENT: Log the user prompt.
                    Write-LogAndHost "Are you sure you want to install all programs? (y/n)" -HostColor Yellow
                    $confirmInput = Read-Host 
                } catch { Write-LogAndHost "ERROR: Could not read user input." -HostColor Red; Start-Sleep -Seconds 2; continue }
                if ($confirmInput.Trim().ToLower() -eq 'y') {
                    Write-LogAndHost "User chose to install all programs." -NoHost; Clear-Host
                    if (Install-Programs -ProgramsToInstall $script:sortedPrograms) { Write-LogAndHost "All programs installation process completed." } else { Write-LogAndHost "Some programs may not have installed correctly. Check log." -HostColor Yellow }
                    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host
                } else { Write-LogAndHost "Installation of all programs cancelled." }
            }
            'g' { 
                if ($script:guiAvailable) {
                    Write-LogAndHost "User chose GUI-based installation." -NoHost
                    $form = New-Object System.Windows.Forms.Form; $form.Text = "Perdanga GUI - Install Programs"; $form.Size = New-Object System.Drawing.Size(400, 450); $form.StartPosition = "CenterScreen"; $form.FormBorderStyle = "FixedDialog"; $form.MaximizeBox = $false; $form.BackColor = [System.Drawing.Color]::FromArgb(0, 30, 60)
                    $panel = New-Object System.Windows.Forms.Panel; $panel.Size = New-Object System.Drawing.Size(360, 350); $panel.Location = New-Object System.Drawing.Point(10, 10); $panel.AutoScroll = $true; $panel.BackColor = [System.Drawing.Color]::FromArgb(0, 30, 60); $form.Controls.Add($panel)
                    $checkboxes = @(); $yPos = 10
                    for ($i = 0; $i -lt $script:sortedPrograms.Length; $i++) {
                        $progName = $script:sortedPrograms[$i]; $dispNumber = $script:programToNumberMap[$progName]; $displayText = "$($dispNumber). $progName".PadRight(30) 
                        $checkbox = New-Object System.Windows.Forms.CheckBox; $checkbox.Text = $displayText; $checkbox.Location = New-Object System.Drawing.Point(10, $yPos); $checkbox.Size = New-Object System.Drawing.Size(330, 24); $checkbox.Font = New-Object System.Drawing.Font("Arial", 12); $checkbox.ForeColor = [System.Drawing.Color]::White; $checkbox.BackColor = [System.Drawing.Color]::FromArgb(0, 30, 60); $panel.Controls.Add($checkbox); $checkboxes += $checkbox; $yPos += 28
                    }
                    $okButton = New-Object System.Windows.Forms.Button; $okButton.Text = "Install Selected"; $okButton.Location = New-Object System.Drawing.Point(140, 370); $okButton.Size = New-Object System.Drawing.Size(120, 30); $okButton.Font = New-Object System.Drawing.Font("Arial", 10); $okButton.ForeColor = [System.Drawing.Color]::White; $okButton.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180); $okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $okButton.FlatAppearance.BorderSize = 0; $okButton.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() }); $form.Controls.Add($okButton)
                    Clear-Host
                    $result = $form.ShowDialog()
                    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                        $selectedProgramsFromGui = @(); for ($i = 0; $i -lt $checkboxes.Length; $i++) { if ($checkboxes[$i].Checked) { $selectedProgramsFromGui += $script:sortedPrograms[$i] } }
                        if ($selectedProgramsFromGui.Count -eq 0) { Write-LogAndHost "No programs selected via GUI for installation." }
                        else { Clear-Host; if (Install-Programs -ProgramsToInstall $selectedProgramsFromGui) { Write-LogAndHost "Selected programs installation process completed. Perdanga Forever." } else { Write-LogAndHost "Some GUI selected programs failed to install." -HostColor Yellow } }
                        Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host
                    } else { Write-LogAndHost "GUI installation cancelled by user." }
                } else { Clear-Host; Write-LogAndHost "ERROR: GUI selection (g) is not available." -HostColor Red; Start-Sleep -Seconds 2; }
            }
            'u' { 
                if ($script:guiAvailable) {
                    Write-LogAndHost "User chose GUI-based uninstallation." -NoHost
                    $installedChocoPackages = Get-InstalledChocolateyPackages
                    if ($installedChocoPackages.Count -eq 0) { Clear-Host; Write-LogAndHost "No Chocolatey packages found to uninstall." -HostColor Yellow; Start-Sleep -Seconds 2; continue }
                    $form = New-Object System.Windows.Forms.Form; $form.Text = "Perdanga GUI - Uninstall Programs"; $form.Size = New-Object System.Drawing.Size(400, 450); $form.StartPosition = "CenterScreen"; $form.FormBorderStyle = "FixedDialog"; $form.MaximizeBox = $false; $form.BackColor = [System.Drawing.Color]::FromArgb(60, 0, 0)
                    $panel = New-Object System.Windows.Forms.Panel; $panel.Size = New-Object System.Drawing.Size(360, 350); $panel.Location = New-Object System.Drawing.Point(10, 10); $panel.AutoScroll = $true; $panel.BackColor = [System.Drawing.Color]::FromArgb(60, 0, 0); $form.Controls.Add($panel)
                    $checkboxes = @(); $yPos = 10
                    foreach ($packageName in ($installedChocoPackages | Sort-Object)) {
                        $checkbox = New-Object System.Windows.Forms.CheckBox; $checkbox.Text = $packageName; $checkbox.Location = New-Object System.Drawing.Point(10, $yPos); $checkbox.Size = New-Object System.Drawing.Size(330, 24); $checkbox.Font = New-Object System.Drawing.Font("Arial", 12); $checkbox.ForeColor = [System.Drawing.Color]::White; $checkbox.BackColor = [System.Drawing.Color]::FromArgb(60, 0, 0); $panel.Controls.Add($checkbox); $checkboxes += $checkbox; $yPos += 28
                    }
                    $okButton = New-Object System.Windows.Forms.Button; $okButton.Text = "Uninstall Selected"; $okButton.Location = New-Object System.Drawing.Point(130, 370); $okButton.Size = New-Object System.Drawing.Size(140, 30); $okButton.Font = New-Object System.Drawing.Font("Arial", 10); $okButton.ForeColor = [System.Drawing.Color]::White; $okButton.BackColor = [System.Drawing.Color]::FromArgb(180, 70, 70); $okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $okButton.FlatAppearance.BorderSize = 0; $okButton.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() }); $form.Controls.Add($okButton)
                    Clear-Host
                    $result = $form.ShowDialog()
                    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                        $selectedProgramsToUninstall = @(); foreach ($cb in $checkboxes) { if ($cb.Checked) { $selectedProgramsToUninstall += $cb.Text } }
                        if ($selectedProgramsToUninstall.Count -eq 0) { Write-LogAndHost "No programs selected via GUI for uninstallation." }
                        else { Clear-Host; if (Uninstall-Programs -ProgramsToUninstall $selectedProgramsToUninstall) { Write-LogAndHost "Selected programs uninstallation process completed. Perdanga Forever." } else { Write-LogAndHost "Some GUI selected programs failed to uninstall." -HostColor Yellow } }
                        Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host
                    } else { Write-LogAndHost "GUI uninstallation cancelled by user." }
                } else { Clear-Host; Write-LogAndHost "ERROR: GUI uninstallation (u) is not available." -HostColor Red; Start-Sleep -Seconds 2; }
            }
            'c' {
                Clear-Host
                $customPackageName = ""
                try {
                    $customPackageName = Read-Host "Enter the exact Chocolatey package ID (e.g., 'notepadplusplus.install', 'git')"
                    $customPackageName = $customPackageName.Trim()
                } catch {
                    Write-LogAndHost "ERROR: Could not read user input for custom package name. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Failed to read custom package name input - "
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
                    # ENHANCEMENT: Log the user prompt.
                    Write-LogAndHost "Package '$customPackageName' found. Proceed with installation?" -HostColor Yellow
                    try {
                        $confirmInstallCustom = Read-Host "(Type y/n then press Enter)"
                        if ($confirmInstallCustom.Trim().ToLower() -eq 'y') {
                            Write-LogAndHost "User confirmed installation of custom package '$customPackageName'." -NoHost
                            Clear-Host
                            if (Install-Programs -ProgramsToInstall @($customPackageName)) {
                                Write-LogAndHost "Custom package '$customPackageName' installation process completed. Perdanga Forever."
                            } else {
                                Write-LogAndHost "Failed to install custom package '$customPackageName'. Check log: $(Split-Path -Leaf $script:logFile)" -HostColor Red
                            }
                        } else {
                            Write-LogAndHost "Installation of custom package '$customPackageName' cancelled by user." -HostColor Yellow
                        }
                    } catch {
                         Write-LogAndHost "ERROR: Could not read user input for custom package installation confirmation. $($_.Exception.Message)" -HostColor Red
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
                    # ENHANCEMENT: Log the user prompt.
                    Write-LogAndHost "This will install Windows Updates. Proceed? (y/n)" -HostColor Yellow
                    $confirmInput = Read-Host 
                } catch { Write-LogAndHost "ERROR: Could not read user input." -HostColor Red; Start-Sleep -Seconds 2; continue }
                if ($confirmInput.Trim().ToLower() -eq 'y') { Clear-Host; Invoke-WindowsUpdate; Write-LogAndHost "Windows Update process finished." } else { Write-LogAndHost "Update process cancelled." }
            }
            'f' { Clear-Host; Create-UnattendXml }
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
                # ENHANCEMENT: Log the user prompt.
                Write-LogAndHost "Install these $($validProgramNamesToInstall.Count) program(s)? (Type y/n then press Enter)" -HostColor Yellow
                $confirmMultiInput = Read-Host
            } catch { Write-LogAndHost "ERROR: Could not read user input. $($_.Exception.Message)" -HostColor Red; Start-Sleep -Seconds 2; continue }
            
            if ($confirmMultiInput.Trim().ToLower() -eq 'y') {
                Clear-Host
                if (Install-Programs -ProgramsToInstall $validProgramNamesToInstall) { Write-LogAndHost "Selected programs installation process completed. Perdanga Forever." }
                else { Write-LogAndHost "Some selected programs failed. Check log: $(Split-Path -Leaf $script:logFile)" -HostColor Yellow }
                Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host
            } else { Write-LogAndHost "Installation of selected programs cancelled." }
        }
    }
    elseif ($script:numberToProgramMap.ContainsKey($userInput)) { 
        $programToInstall = $script:numberToProgramMap[$userInput]
        Clear-Host
        try {
            # ENHANCEMENT: Log the user prompt.
            Write-LogAndHost "Install '$($programToInstall)' (program #$($userInput))? (Type y/n then press Enter)" -HostColor Yellow
            $confirmSingleInput = Read-Host
        } catch { Write-LogAndHost "ERROR: Could not read user input. $($_.Exception.Message)" -HostColor Red; Start-Sleep -Seconds 2; continue }
        
        if ($confirmSingleInput.Trim().ToLower() -eq 'y') {
            Clear-Host
            if (Install-Programs -ProgramsToInstall @($programToInstall)) { Write-LogAndHost "$($programToInstall) installation process completed. Perdanga Forever." }
            else { Write-LogAndHost "Failed to install $($programToInstall). Check log: $(Split-Path -Leaf $script:logFile)" -HostColor Red }
            Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host
        } else { Write-LogAndHost "Installation of '$($programToInstall)' cancelled." }
    }
    else {
        Clear-Host
        Write-LogAndHost "Invalid selection: '$($userInput)'. Use options [A,G,U,C,T,X,W,N,F,E] or program numbers." -HostColor Red
        Start-Sleep -Seconds 2
    }
} while ($true)
