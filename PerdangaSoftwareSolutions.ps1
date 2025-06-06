<#
.SYNOPSIS
    Main script for Perdanga Software Solutions to install and manage programs using Chocolatey.

.DESCRIPTION
    This is the main entry point for the script. It loads configuration and functions from separate
    files, performs initial checks, initializes Chocolatey, and then enters the main user interaction loop.

.NOTES
    Author: Roman Zhdanov
    Version: 1.3 (Telemetry Control)
#>

# --- INITIAL SETUP ---
# Verify running in PowerShell
if ($PSVersionTable.PSVersion -eq $null) {
    Write-Host "ERROR: This script must be run in PowerShell, not in Command Prompt." -ForegroundColor Red
    Write-Host "Reason: The script uses PowerShell-specific cmdlets and features." -ForegroundColor Red
    Write-Host "Solution: Run the script using RunPerdangaSoftwareSolutions.bat or directly in PowerShell." -ForegroundColor Yellow
    exit 1
}

# Set console encoding to UTF-8 for proper ASCII art rendering
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Set console buffer size and window size
try {
    $Host.UI.RawUI.BufferSize = New-Object Management.Automation.Host.Size(150, 3000)
    $Host.UI.RawUI.WindowSize = New-Object Management.Automation.Host.Size(150, 50)
} catch {
    Write-Host "WARNING: Could not set console buffer or window size. This may happen in some environments (e.g., VS Code integrated terminal)." -ForegroundColor Yellow
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor DarkYellow
}

# Set log file name with timestamp and use script directory
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$script:logFile = Join-Path -Path $PSScriptRoot -ChildPath "install_log_$timestamp.txt"

# --- LOAD MODULES ---
# Dot-source the configuration and function files.
# The scripts will run in the current scope, making variables and functions available.
# It's crucial to load functions first, as config might use Write-LogAndHost.
try {
    . (Join-Path -Path $PSScriptRoot -ChildPath "functions.ps1")
    . (Join-Path -Path $PSScriptRoot -ChildPath "config.ps1")
} catch {
    Write-Host "ERROR: Failed to load 'config.ps1' or 'functions.ps1'." -ForegroundColor Red
    Write-Host "Please ensure both files are in the same directory as this main script." -ForegroundColor Red
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor DarkYellow
    # Try to log to the file even if function loading failed partially
    "[$((Get-Date))] CRITICAL: Failed to load script modules. Error: $($_.Exception.Message)" | Out-File -FilePath $script:logFile -Append -Encoding UTF8
    exit 1
}


# --- SCRIPT-WIDE CHECKS ---
# Define a variable to track whether activation was attempted
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
        # Re-check after install attempt
        if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
            Write-LogAndHost "ERROR: Chocolatey command (choco) still not found after installation attempt. Please install Chocolatey manually and re-run." -HostColor Red
            exit 1
        }
         # Set ChocolateyInstall environment variable if not set (common after fresh install)
        if (-not $env:ChocolateyInstall) {
            $env:ChocolateyInstall = "$($env:ProgramData)\chocolatey"
            Write-LogAndHost "ChocolateyInstall environment variable set to: $env:ChocolateyInstall" -NoHost
        }
    } else {
         if (-not $env:ChocolateyInstall) {
            $env:ChocolateyInstall = "$($env:ProgramData)\chocolatey" # Or try to get it from choco path
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

    # Validate user input against allowed characters (main menu letters, numbers, spaces, commas)
    # ADDED 't' to the regex for Telemetry option
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
                    & choco list --localonly 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8 # Use --localonly for installed
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
                         Start-Sleep -Seconds 2 # Give user time to see error
                    }
                } else {
                    # Test-ChocolateyPackage already logs details of why it failed (not found, error, etc.)
                    Write-LogAndHost "Custom package '$customPackageName' could not be installed (either not found or validation failed)." -HostColor Red
                }
                Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
                $null = Read-Host
            }
            't' { # ADDED: Disable Telemetry
                Clear-Host
                Write-LogAndHost "User chose to disable Telemetry." -NoHost
                Invoke-DisableTelemetry
                # Invoke-DisableTelemetry already has a Read-Host at the end
            }
            'x' { # Spotify Activation (formerly SpotX Enhancement)
                Clear-Host
                Write-LogAndHost "User chose Spotify Activation." -NoHost
                Invoke-SpotXActivation
                # Invoke-SpotXActivation already has a Read-Host at the end
            }
            'w' { 
                Clear-Host; 
                Write-LogAndHost "User chose to activate Windows." -NoHost 
                Invoke-WindowsActivation; 
                # Invoke-WindowsActivation already has a Read-Host at the end
            }
            'n' { # Update Windows (formerly Combined Update Windows & Update All Programs)
                Clear-Host
                try {
                    Write-LogAndHost "This option will check for and install Windows Updates." -HostColor Yellow # Updated text
                    Write-LogAndHost "Are you sure you want to proceed? (Type y/n then press Enter)" -HostColor Yellow -NoLog
                    $confirmInput = Read-Host
                } catch { Write-LogAndHost "ERROR: Could not read user input. $($_.Exception.Message)" -HostColor Red; Start-Sleep -Seconds 2; continue }
                
                if ($confirmInput.Trim().ToLower() -eq 'y') {
                    Write-LogAndHost "User chose to update Windows." -NoHost # Updated text
                    Clear-Host
                    Invoke-WindowsUpdate # This function has its own Read-Host
                    Write-LogAndHost "Windows Update process finished." # Updated text
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
        Write-LogAndHost "Invalid selection: '$($userInput)'. Use options [A,G,U,C,T,X,W,N,E] or program numbers." -HostColor Red # Updated message
        Start-Sleep -Seconds 2
    }
} while ($true)
