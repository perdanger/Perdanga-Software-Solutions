<#
.SYNOPSIS
    Functions library for Perdanga Software Solutions.

.DESCRIPTION
    This file contains all the helper functions used by the main script.
    It is dot-sourced to make these functions available in the main script's scope.

.NOTES
    This script should not be run directly.
#>

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
    $menuLines.Add(" Chocolatey Package Manager [PSS v1.3] ($(Get-Date -Format "dd.MM.yyyy HH:mm"))") 
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
    
    # --- START OF FIX ---
    # The old method used hardcoded spaces, causing misalignment.
    # The new method defines options in pairs and calculates spacing dynamically.

    # Define options in a structured way for proper alignment.
    $optionPairs = @(
        @{ Left = "[A] Install All Programs";              Right = "[W] Activate Windows" },
        @{ Left = "[G] Select Specific Programs via GUI";  Right = "[N] Update Windows" },
        @{ Left = "[U] Uninstall Programs via GUI";        Right = "[T] Disable Windows Telemetry" },
        @{ Left = "[C] Install Custom Program";            Right = "" },
        @{ Left = "[X] Activate Spotify";                  Right = "" }
    )

    # Calculate the width for the first column to ensure all options align vertically.
    # This is based on the longest string in the left column, plus some padding.
    $column1Width = ($optionPairs.Left | Measure-Object -Property Length -Maximum).Maximum + 5

    # Build and add each option line to the menu array.
    foreach ($pair in $optionPairs) {
        # Pad the left item to the calculated width to align the second column.
        $leftColumn = $pair.Left.PadRight($column1Width)
        # Combine the columns into one line.
        $fullLine = "{0}{1}" -f $leftColumn, $pair.Right
        # Add the line to the menu with no extra indent to align with the start of the underline.
        $menuLines.Add($fullLine.TrimEnd())
    }

    # Add a blank line for visual separation.
    $menuLines.Add("")

    # Center the 'Exit' option relative to the overall menu width for a clean look.
    $exitLine = "[E] Exit Script"
    $exitPadding = [math]::Floor(($fixedMenuWidth - $exitLine.Length) / 2)
    if ($exitPadding -lt 0) { $exitPadding = 0 }
    $centeredExitLine = (" " * $exitPadding) + $exitLine
    $menuLines.Add($centeredExitLine)

    # --- END OF FIX ---
    
    $menuLines.Add($optionsUnderline)

    $consoleWidth = $Host.UI.RawUI.WindowSize.Width
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
