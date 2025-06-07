<#
.SYNOPSIS
    Functions library for Perdanga Software Solutions.

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
    $fullLogMessage = "[$(Get-Date)] $LogPrefix$Message"
    if (-not $NoLog) {
        # This line will now work correctly from the very beginning
        $fullLogMessage | Out-File -FilePath $script:logFile -Append -Encoding UTF8
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
        Write-LogAndHost "Chocolatey is not installed. Would you like to install it? (Type y/n then press Enter)" -HostColor Yellow -NoLog
        $confirmInput = Read-Host
        if ($confirmInput.Trim().ToLower() -eq 'y') {
            Write-LogAndHost "User chose to install Chocolatey." -NoHost
            try {
                Write-LogAndHost "Installing Chocolatey..." -NoLog
                Set-ExecutionPolicy Bypass -Scope Process -Force
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072 # Tls12
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
    Write-LogAndHost "WARNING: Windows activation uses an external script from 'https://get.activated.win'. Ensure you trust the source before proceeding." -HostColor Yellow -NoLog
    try {
        Write-LogAndHost "Continue with Windows activation? (Type y/n then press Enter)" -HostColor Yellow -NoLog
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
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-RestMethod -Uri "https://get.activated.win" | Invoke-Expression 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8
        # irm https://get.activated.win | iex (original command)
        if ($LASTEXITCODE -eq 0) { # This might not be reliable for iex from irm
            Write-LogAndHost "Windows activation script executed. Check console output for status."
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
        Write-LogAndHost "Continue with Spotify Activation? (Type y/n then press Enter)" -HostColor Yellow -NoLog
        $confirmSpotX = Read-Host
        if ($confirmSpotX.Trim().ToLower() -ne 'y') {
            Write-LogAndHost "Spotify Activation cancelled by user." -HostColor Yellow
            return
        }
    } catch {
        Write-LogAndHost "ERROR: Could not read user input for Spotify Activation confirmation. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Failed to read user input for Spotify Activation confirmation - "
        return
    }

    $spotxParams = "-new_theme" # SpotX parameters from original bat file
    $spotxUrlPrimary = 'https://raw.githubusercontent.com/SpotX-Official/spotx-official.github.io/main/run.ps1'
    $spotxUrlFallback = 'https://spotx-official.github.io/run.ps1'

    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

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
            $_ | Out-File -FilePath $script:logFile -Append -Encoding UTF8 # Log the full error for fallback
            Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
            $null = Read-Host
            return
        }
    }

    if ($scriptContentToExecute) {
        Write-LogAndHost "Executing SpotX script with parameters: '$effectiveParams'"
        $fullScriptToRun = "$scriptContentToExecute $effectiveParams"
        try {
            Invoke-Expression -Command $fullScriptToRun 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8
            Write-LogAndHost "SpotX script execution attempt finished. Check console output from SpotX above for status." -HostColor Green
        } catch {
            Write-LogAndHost "ERROR: Exception occurred during SpotX script execution. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during SpotX execution - "
            $_ | Out-File -FilePath $script:logFile -Append -Encoding UTF8
        }
    } else {
        # This case should be covered by the try-catch blocks above, but as a safeguard:
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
            Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction Stop # Install for current user to avoid admin issues with Install-Module
            Write-LogAndHost "PSWindowsUpdate module installed successfully for current user."
        }
        Import-Module PSWindowsUpdate -ErrorAction Stop
        Write-LogAndHost "Checking for available updates..."
        $updates = Get-WindowsUpdate -ErrorAction Stop
        if ($updates.Count -gt 0) {
            Write-LogAndHost "Found $($updates.Count) updates. Installing..."
            Install-WindowsUpdate -AcceptAll -AutoReboot -ErrorAction Stop # Added AutoReboot, user should be aware
            Write-LogAndHost "Windows updates installed successfully. A reboot might be required."
        } else {
            Write-LogAndHost "No updates available."
        }
    } catch {
        Write-LogAndHost "ERROR: Failed to update Windows. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Error during Windows update: "
    }
    Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
    $null = Read-Host
}

# UPDATED FUNCTION: Disable Windows Telemetry (with check)
function Invoke-DisableTelemetry {
    Write-LogAndHost "Checking Windows Telemetry status..." -HostColor Cyan
    
    # Check if telemetry is already disabled
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
        Write-LogAndHost "Telemetry is currently enabled. Continue with disabling? (Type y/n then press Enter)" -HostColor Yellow -NoLog
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
        # Disable and Stop Telemetry Services
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

        # Set Registry Keys to disable Telemetry
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
                Write-LogAndHost "Setting registry value '$($keyInfo.Name)' at '$path'" -NoLog
                Set-ItemProperty -Path $path -Name $keyInfo.Name -Value $keyInfo.Value -Type $keyInfo.Type -Force -ErrorAction Stop
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
    Write-Host "" # Add a newline for better readability

    foreach ($program in $ProgramsToInstall) {
        Write-LogAndHost "Installing $program..." -LogPrefix "Installing $program (from list)..."
        try {
            $installOutput = & choco install $program -y --source=$Source --no-progress 2>&1
            $installOutput | Out-File -FilePath $script:logFile -Append -Encoding UTF8 # Always log raw output

            if ($LASTEXITCODE -eq 0) {
                if ($installOutput -match "already installed|Nothing to do") {
                    Write-LogAndHost "$program is already installed or up to date." -HostColor Green
                } else {
                    Write-LogAndHost "$program installed successfully." -HostColor White
                }
            } else {
                $allSuccess = $false
                Write-LogAndHost "ERROR: Failed to install $program. Exit code: $LASTEXITCODE. Details: $($installOutput | Out-String)" -HostColor Red -LogPrefix "Error installing $program. "
            }
        } catch {
            $allSuccess = false
            Write-LogAndHost "ERROR: Exception occurred while installing $program. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during installation of $program - "
        }
        Write-Host ""
    }
    return $allSuccess
}

# Function to get installed Chocolatey packages
function Get-InstalledChocolateyPackages {
    $chocoLibPath = Join-Path -Path $env:ChocolateyInstall -ChildPath "lib" # Use env variable for choco path
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
            $allSuccess = false
            Write-LogAndHost "ERROR: Exception occurred while uninstalling $program. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during uninstallation of $program - "
        }
        Write-Host ""
    }
    return $allSuccess
}

# Function to test if a Chocolatey package exists
function Test-ChocolateyPackage {
    param (
        [string]$PackageName
    )
    Write-LogAndHost "Searching for package '$PackageName' in Chocolatey repository..." -NoLog
    try {
        # Suppress progress for search command
        $searchOutput = & choco search $PackageName --exact --limit-output --source="https://community.chocolatey.org/api/v2/" --no-progress 2>&1
        $searchOutput | Out-File -FilePath $script:logFile -Append -Encoding UTF8

        if ($LASTEXITCODE -ne 0) {
             Write-LogAndHost "Error during 'choco search' for '$PackageName'. Exit code: $LASTEXITCODE. Output: $($searchOutput | Out-String)" -HostColor Red
             return $false
        }
        
        # Check if the output contains the package name (choco search output format can vary)
        # A simple check: if the package name is found in the output lines.
        # More robust: choco search <pkg> --exact returns "<pkg>|<version>" if found, or "0 packages found."
        if ($searchOutput -match "$($PackageName)\|.*" -or $searchOutput -match "1 packages found.") { # Check for exact match pattern or count
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

    # ASCII Art for the header
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

    # Print ASCII art in Cyan
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
    $menuLines.Add(" Chocolatey Package Manager [PSS v1.3] ($(Get-Date -Format 'dd.MM.yyyy HH:mm'))") 
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

    # Format programs into three columns
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
    
    # Define options for two-column layout
    $optionPairs = @(
        @{ Left = "[A] Install All Programs";              Right = "[W] Activate Windows" },
        @{ Left = "[G] Select Specific Programs via GUI";  Right = "[N] Update Windows" },
        @{ Left = "[U] Uninstall Programs via GUI";        Right = "[T] Disable Windows Telemetry" },
        @{ Left = "[C] Install Custom Program";            Right = "[X] Activate Spotify" }
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

    $consoleWidth = $Host.UI.RawUI.WindowSize.Width
    $blockPaddingValue = [math]::Floor(($consoleWidth - $fixedMenuWidth) / 2)
    if ($blockPaddingValue -lt 0) { $blockPaddingValue = 0 }
    $blockPaddingString = " " * $blockPaddingValue

    # Print menu with specified colors
    foreach ($lineEntry in $menuLines) {
        $trimmedEntry = $lineEntry.Trim()
        # These specific lines will be Cyan
        if ($trimmedEntry -eq $pssText -or
            $trimmedEntry -like ($pssUnderline.Trim()) -or
            $trimmedEntry -like ($dashedLine.Trim()) -or
            $trimmedEntry -eq $programHeader -or
            $trimmedEntry -eq $optionsHeader -or
            $trimmedEntry -like ($programUnderline.Trim()) -or
            $trimmedEntry -like ($optionsUnderline.Trim())) {
            Write-Host ($blockPaddingString + $lineEntry) -ForegroundColor Cyan
        } else {
            # All other lines will be White
            Write-Host ($blockPaddingString + $lineEntry) -ForegroundColor White
        }
    }
    Write-Host "" 

    $promptTextForOneLine = "Enter option, single number, or list of numbers:"
    $promptPaddingOneLine = [math]::Floor(($fixedMenuWidth - $promptTextForOneLine.Length) / 2)
    if ($promptPaddingOneLine -lt 0) { $promptPaddingOneLine = 0 }
    $centeredPromptOneLine = (" " * $promptPaddingOneLine) + $promptTextForOneLine
    
    # Changed prompt color to White to match screenshot
    Write-Host ($blockPaddingString + $centeredPromptOneLine) -NoNewline -ForegroundColor White
}

# Log that the functions library has been loaded successfully
Write-LogAndHost "Functions library loaded." -NoHost

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
    "discord",
    "file-converter",
    "git",
    "googlechrome",
    "gpu-z",
    "hwmonitor",
    "imageglass",
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
# The following variables are set in the 'script' scope to be accessible
# throughout the main script and all functions.

# Sort the program list for consistent display
$script:sortedPrograms = $programs | Sort-Object

# Define the letters for main menu commands
$script:mainMenuLetters = @('a', 'e', 'g', 'n', 'w', 'u', 'c', 'x', 't')

# Generate numbers for program selection
$script:availableProgramNumbers = 1..($script:sortedPrograms.Count) | ForEach-Object { $_.ToString() }

# Create maps for program-to-number and number-to-program lookups
$script:programToNumberMap = @{} # Maps program name to its assigned number
$script:numberToProgramMap = @{} # Maps assigned number to program name

# Check if there are enough numbers for all programs
if ($script:sortedPrograms.Count -gt $script:availableProgramNumbers.Count) {
    Write-LogAndHost "WARNING: Not enough unique numbers available to assign to all programs." -ForegroundColor Yellow -LogPrefix "CRITICAL WARNING: "
}

# Populate the lookup maps
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

# Log that the configuration has been loaded successfully
Write-LogAndHost "Configuration loaded. $($script:sortedPrograms.Count) programs defined." -NoHost

<#
.SYNOPSIS
    Main script for Perdanga Software Solutions to install and manage programs using Chocolatey.

.DESCRIPTION
    This is the main entry point for the script. It loads configuration and functions from separate
    files, performs initial checks, initializes Chocolatey, and then enters the main user interaction loop.

.NOTES
    Author: Roman Zhdanov
    Version: 1.3
#>

# --- INITIAL SETUP ---
# Verify running in PowerShell
if ($PSVersionTable.PSVersion -eq $null) {
    Write-Host "ERROR: This script must be run in PowerShell, not in Command Prompt." -ForegroundColor Red
    exit 1
}

# Set console encoding to UTF-8 for proper ASCII art rendering
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Set console buffer size, window size, and background color
try {
    $Host.UI.RawUI.BackgroundColor = 'DarkCyan'
    $Host.UI.RawUI.BufferSize = New-Object Management.Automation.Host.Size(150, 3000)
    $Host.UI.RawUI.WindowSize = New-Object Management.Automation.Host.Size(150, 50)
    Clear-Host # Clear the screen to apply the new background color immediately
} catch {
    Write-Host "WARNING: Could not set console UI properties. Error: $($_.Exception.Message)" -ForegroundColor Yellow
}

# --- SCRIPT-WIDE CHECKS ---
$script:activationAttempted = $false

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-LogAndHost "ERROR: This script must be run as Administrator." -ForegroundColor Red
    "[$((Get-Date))] Error: Script not run as Administrator." | Out-File -FilePath $script:logFile -Append -Encoding UTF8
    exit 1
}

# Check GUI availability early
$script:guiAvailable = $true
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
}
catch {
    $script:guiAvailable = $false
    "[$((Get-Date))] Warning: GUI is not available - $($_.Exception.Message)" | Out-File -FilePath $script:logFile -Append -Encoding UTF8
}


# --- CHOCOLATEY INITIALIZATION ---
Write-LogAndHost "Checking Chocolatey installation..."
try {
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        if (-not (Install-Chocolatey)) {
            Write-LogAndHost "ERROR: Chocolatey is required to proceed. Exiting script." -HostColor Red
            exit 1
        }
        if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
            Write-LogAndHost "ERROR: Chocolatey command (choco) still not found after installation attempt. Please install Chocolatey manually and re-run." -HostColor Red
            exit 1
        }
        if (-not $env:ChocolateyInstall) {
            $env:ChocolateyInstall = "$($env:ProgramData)\chocolatey"
            Write-LogAndHost "ChocolateyInstall environment variable set to: $env:ChocolateyInstall" -NoHost
        }
    } else {
         if (-not $env:ChocolateyInstall) {
            $env:ChocolateyInstall = "$($env:ProgramData)\chocolatey"
            Write-LogAndHost "ChocolateyInstall environment variable set to: $env:ChocolateyInstall" -NoHost
        }
    }
    $chocoVersion = & choco --version 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-LogAndHost "ERROR: Chocolatey is not installed or not functioning correctly. Exit code: $LASTEXITCODE" -HostColor Red
        exit 1
    }
    Write-LogAndHost "Found Chocolatey version: $($chocoVersion -join ' ')"
}
catch {
    Write-LogAndHost "ERROR: Exception occurred while checking Chocolatey. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during Chocolatey check - "
    exit 1
}
Write-Host ""
Write-LogAndHost "Clearing Chocolatey cache..."
try {
    & choco cache remove --all 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8 
    if ($LASTEXITCODE -eq 0) { Write-LogAndHost "Cache cleared." }
    else { Write-LogAndHost "WARNING: Failed to clear Chocolatey cache. $($LASTEXITCODE)" -HostColor Yellow }
}
catch { Write-LogAndHost "ERROR: Exception occurred while clearing Chocolatey cache. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during cache clearing - " }
Write-Host ""
Write-LogAndHost "Enabling automatic confirmation..."
try {
    & choco feature enable -n allowGlobalConfirmation 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8
    if ($LASTEXITCODE -eq 0) { Write-LogAndHost "Automatic confirmation enabled." }
    else { Write-LogAndHost "WARNING: Failed to enable automatic confirmation. $($LASTEXITCODE)" -HostColor Yellow }
}
catch { Write-LogAndHost "ERROR: Exception occurred while enabling automatic confirmation. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during enabling automatic confirmation - " }
Write-Host ""


# --- MAIN LOOP ---
do {
    Show-Menu
    try {
        $userInput = Read-Host
    } catch {
        Write-LogAndHost "ERROR: Could not read user input. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Failed to read user input - "
        Start-Sleep -Seconds 2
        continue
    }
    $userInput = $userInput.Trim().ToLower()

    if ([string]::IsNullOrEmpty($userInput)) {
        Clear-Host
        Write-LogAndHost "No input detected. Please enter an option or program number(s)." -HostColor Yellow
        Write-LogAndHost "Press any key to return to the menu..." -HostColor DarkGray -NoLog
        $null = Read-Host 
        continue 
    }

    if ($userInput -and $userInput -notmatch "^[aegnuwcxt0-9\s,]+$") { 
        Clear-Host
        Write-LogAndHost "Invalid input: '$userInput'. Use options [A,G,U,C,T,X,W,N,E] or program numbers." -HostColor Red -LogPrefix "Invalid user input: '$userInput' - contains invalid characters."
        Start-Sleep -Seconds 2
        continue
    }

    # Handle main menu letter options
    if ($script:mainMenuLetters -contains $userInput) {
        switch ($userInput) {
            'e' {
                if (-not $script:activationAttempted) { Write-LogAndHost "Exiting script. Windows activation not attempted by user choice on this run." -NoHost }
                Write-LogAndHost "Exiting script..."
                Write-LogAndHost "Disabling automatic confirmation..."
                try {
                    & choco feature disable -n allowGlobalConfirmation 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8
                    if ($LASTEXITCODE -eq 0) { Write-LogAndHost "Automatic confirmation disabled." }
                } catch { Write-LogAndHost "ERROR: Exception during disabling automatic confirmation - $($_.Exception.Message)" -HostColor Red }
                Write-Host ""
                Write-LogAndHost "Checking installed programs..."
                try {
                    & choco list --localonly 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8
                    if ($LASTEXITCODE -eq 0) { Write-LogAndHost "Installed programs list (at exit) saved to $($script:logFile)" }
                } catch { Write-LogAndHost "ERROR: Exception during listing installed programs - $($_.Exception.Message)" -HostColor Red }
                Write-Host ""
                exit 0
            }
            'a' {
                Clear-Host
                try {
                    Write-LogAndHost "Are you sure you want to install all programs? (Type y/n then press Enter)" -HostColor Yellow -NoLog
                    $confirmInput = Read-Host
                } catch {
                    Write-LogAndHost "ERROR: Could not read user input. $($_.Exception.Message)" -HostColor Red
                    Start-Sleep -Seconds 2
                    continue
                }
                if ($confirmInput.Trim().ToLower() -eq 'y') {
                    Write-LogAndHost "User chose to install all programs." -NoHost
                    Clear-Host
                    if (Install-Programs -ProgramsToInstall $script:sortedPrograms) { Write-LogAndHost "All programs installation process completed. Perdanga Forever." }
                    else { Write-LogAndHost "Some programs may not have installed correctly. Check log: $(Split-Path -Leaf $script:logFile)" -HostColor Yellow }
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
            'c' { # Install Custom Package
                Clear-Host
                Write-LogAndHost "User chose to install a custom package." -NoHost
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
                    Write-LogAndHost "Package '$customPackageName' found. Proceed with installation?" -HostColor Yellow -NoLog
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
            't' {
                Clear-Host
                Write-LogAndHost "User chose to disable Telemetry." -NoHost
                Invoke-DisableTelemetry
            }
            'x' {
                Clear-Host
                Write-LogAndHost "User chose Spotify Activation." -NoHost
                Invoke-SpotXActivation
            }
            'w' { 
                Clear-Host
                Write-LogAndHost "User chose to activate Windows." -NoHost 
                Invoke-WindowsActivation
            }
            'n' {
                Clear-Host
                try {
                    Write-LogAndHost "This option will check for and install Windows Updates." -HostColor Yellow
                    Write-LogAndHost "Are you sure you want to proceed? (Type y/n then press Enter)" -HostColor Yellow -NoLog
                    $confirmInput = Read-Host
                } catch { Write-LogAndHost "ERROR: Could not read user input. $($_.Exception.Message)" -HostColor Red; Start-Sleep -Seconds 2; continue }
                
                if ($confirmInput.Trim().ToLower() -eq 'y') {
                    Write-LogAndHost "User chose to update Windows." -NoHost
                    Clear-Host
                    Invoke-WindowsUpdate
                    Write-LogAndHost "Windows Update process finished."
                }
                else { Write-LogAndHost "Update process cancelled." }
            }
        }
    }
    # Handle multiple program installations via comma-separated numbers
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
                Write-LogAndHost "Install these $($validProgramNamesToInstall.Count) program(s)? (Type y/n then press Enter)" -HostColor Yellow -NoLog
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
    # Handle single program installation via number
    elseif ($script:numberToProgramMap.ContainsKey($userInput)) { 
        $programToInstall = $script:numberToProgramMap[$userInput]
        Clear-Host
        try {
            Write-LogAndHost "Install '$($programToInstall)' (program #$($userInput))? (Type y/n then press Enter)" -HostColor Yellow -NoLog
            $confirmSingleInput = Read-Host
        } catch { Write-LogAndHost "ERROR: Could not read user input. $($_.Exception.Message)" -HostColor Red; Start-Sleep -Seconds 2; continue }
        
        if ($confirmSingleInput.Trim().ToLower() -eq 'y') {
            Clear-Host
            if (Install-Programs -ProgramsToInstall @($programToInstall)) { Write-LogAndHost "$($programToInstall) installation process completed. Perdanga Forever." }
            else { Write-LogAndHost "Failed to install $($programToInstall). Check log: $(Split-Path -Leaf $script:logFile)" -HostColor Red }
            Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray; $null = Read-Host
        } else { Write-LogAndHost "Installation of '$($programToInstall)' cancelled." }
    }
    # Handle invalid input
    else {
        Clear-Host
        Write-LogAndHost "Invalid selection: '$($userInput)'. Use options [A,G,U,C,T,X,W,N,E] or program numbers." -HostColor Red
        Start-Sleep -Seconds 2
    }
} while ($true)
