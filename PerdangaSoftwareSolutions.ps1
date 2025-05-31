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

# Define programs to install
$programs = @("7zip.install", "brave", "file-converter", "googlechrome", "gpu-z", "hwmonitor", "imageglass", "nvidia-app", "obs-studio", "occt", "qbittorrent", "revo-uninstaller", "spotify", "steam", "telegram", "vcredist-all", "vlc", "winrar", "wiztree")
$script:sortedPrograms = $programs | Sort-Object

# Assign numbers to programs, excluding main menu command letters
$script:mainMenuLetters = @('a', 'e', 'g', 'n', 'w') # Lowercase main command letters
$script:availableProgramNumbers = 1..($script:sortedPrograms.Count) | ForEach-Object { $_.ToString() }

$script:programToNumberMap = @{} # Maps program name to its assigned number for display
$script:numberToProgramMap = @{} # Maps assigned number to program name for installation

if ($script:sortedPrograms.Count -gt $script:availableProgramNumbers.Count) {
    Write-Host "WARNING: Not enough unique numbers available to assign to all programs." -ForegroundColor Yellow
    "[$((Get-Date))] CRITICAL WARNING: Not enough unique numbers for programs. Program count: $($script:sortedPrograms.Count), Available numbers: $($script:availableProgramNumbers.Count)" | Out-File -FilePath $script:logFile -Append -Encoding UTF8
}

for ($i = 0; $i -lt $script:sortedPrograms.Count; $i++) {
    if ($i -lt $script:availableProgramNumbers.Count) {
        $assignedNumber = $script:availableProgramNumbers[$i]
        $programName = $script:sortedPrograms[$i]
        $script:programToNumberMap[$programName] = $assignedNumber
        $script:numberToProgramMap[$assignedNumber] = $programName
    } else {
        "[$((Get-Date))] WARNING: Ran out of assignable numbers. Program '$($script:sortedPrograms[$i])' will not be selectable by number." | Out-File -FilePath $script:logFile -Append -Encoding UTF8
        break # Stop assigning if we run out of numbers
    }
}

# Define a variable to track whether activation was attempted
$script:activationAttempted = $false

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator." -ForegroundColor Red
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

#endregion

#region Helper Functions

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
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
                $installOutput = Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1')) 2>&1
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
        irm https://get.activated.win | iex 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8
        if ($LASTEXITCODE -eq 0) {
            Write-LogAndHost "Windows activation completed successfully."
        }
        else {
            Write-LogAndHost "ERROR: Windows activation failed. Exit code: $LASTEXITCODE" -HostColor Red
        }
    }
    catch {
        Write-LogAndHost "ERROR: Exception during Windows activation - $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during Windows activation - "
    }
}

# Function to update all installed programs
function Invoke-UpdateAllPrograms {
    Write-LogAndHost "Updating all installed programs..."
    try {
        $updateOutput = & choco upgrade all -y --source="https://community.chocolatey.org/api/v2/" --no-progress 2>&1
        $updateOutput | Out-File -FilePath $script:logFile -Append -Encoding UTF8
        if ($LASTEXITCODE -eq 0) {
            Write-LogAndHost "All programs updated successfully."
        }
        else {
            Write-LogAndHost "ERROR: Failed to update programs." -HostColor Red
        }
    }
    catch {
        Write-LogAndHost "ERROR: Exception occurred while updating programs." -HostColor Red -LogPrefix "Error: Exception during program update - "
    }
    Write-Host ""
}

# Function to perform Windows update
function Invoke-WindowsUpdate {
    Write-LogAndHost "Checking for Windows updates..."
    try {
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-LogAndHost "PSWindowsUpdate module not found. Installing..."
            Install-Module -Name PSWindowsUpdate -Force -ErrorAction Stop
            Write-LogAndHost "PSWindowsUpdate module installed successfully."
        }
        Import-Module PSWindowsUpdate -ErrorAction Stop
        Write-LogAndHost "Checking for available updates..."
        $updates = Get-WindowsUpdate -ErrorAction Stop
        if ($updates.Count -gt 0) {
            Write-LogAndHost "Found $($updates.Count) updates. Installing..."
            Install-WindowsUpdate -AcceptAll -ErrorAction Stop
            Write-LogAndHost "Windows updates installed successfully."
        } else {
            Write-LogAndHost "No updates available."
        }
    } catch {
        Write-LogAndHost "ERROR: Failed to update Windows." -HostColor Red -LogPrefix "Error during Windows update: "
    }
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
            $allSuccess = $false
            Write-LogAndHost "ERROR: Exception occurred while installing $program. Details: $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during installation of $program - "
        }
        Write-Host ""
    }
    return $allSuccess
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

    # Fixed width of the main menu block
    $fixedMenuWidth = 67

    # Header section
    $pssText = "Perdanga Software Solutions"
    $pssUnderline = "==================================================================="
    $dashedLine = "-------------------------------------------------------------------"

    $fixedHeaderPadding = 22
    $centeredPssTextLine = (" " * $fixedHeaderPadding) + $pssText

    $menuLines.Add($pssUnderline)
    $menuLines.Add($centeredPssTextLine)
    $menuLines.Add($dashedLine)
    $menuLines.Add(" Installing programs via Chocolatey ($(Get-Date -Format "MM/dd/yyyy hh:mm tt"))")
    $menuLines.Add(" Log saved to: install_log_$timestamp.txt")
    $menuLines.Add(" $($script:sortedPrograms.Count) programs available for installation")
    $menuLines.Add($pssUnderline)
    $menuLines.Add("")

    # Available Programs section
    $programHeader = "Available Programs:"
    $centeredProgramHeader = (" " * $fixedHeaderPadding) + $programHeader
    $programUnderline = "-------------------------------------------------------------------"

    $menuLines.Add($centeredProgramHeader)
    $menuLines.Add($programUnderline)

    $formattedPrograms = @()
    $sortedDisplayNumbers = $script:numberToProgramMap.Keys | Sort-Object { [int]$_ }

    foreach ($dispNumber in $sortedDisplayNumbers) {
        $programName = $script:numberToProgramMap[$dispNumber]
        $formattedPrograms += "$($dispNumber). $($programName)"
    }

    $programColumns = @{
        0 = @()
        1 = @()
        2 = @()
    }

    $programsPerColumn = [math]::Ceiling($formattedPrograms.Count / 3.0) # Distribute programs evenly

    for ($i = 0; $i -lt $formattedPrograms.Count; $i++) {
        if ($i -lt $programsPerColumn) {
            $programColumns[0] += $formattedPrograms[$i]
        } elseif ($i -lt ($programsPerColumn * 2)) {
            $programColumns[1] += $formattedPrograms[$i]
        } else {
            $programColumns[2] += $formattedPrograms[$i]
        }
    }

    $col1MaxLength = ($programColumns[0] | Measure-Object -Property Length -Maximum).Maximum
    $col2MaxLength = ($programColumns[1] | Measure-Object -Property Length -Maximum).Maximum
    $col3MaxLength = ($programColumns[2] | Measure-Object -Property Length -Maximum).Maximum
    if ($col1MaxLength -eq $null) {$col1MaxLength = 0}
    if ($col2MaxLength -eq $null) {$col2MaxLength = 0}
    if ($col3MaxLength -eq $null) {$col3MaxLength = 0}

    $maxRows = [math]::Max($programColumns[0].Count, [math]::Max($programColumns[1].Count, $programColumns[2].Count))

    for ($i = 0; $i -lt $maxRows; $i++) {
        $line = "  "

        if ($i -lt $programColumns[0].Count) {
            $line += $programColumns[0][$i].PadRight($col1MaxLength)
        } else {
            $line += "".PadRight($col1MaxLength)
        }
        $line += "    "

        if ($i -lt $programColumns[1].Count) {
            $line += $programColumns[1][$i].PadRight($col2MaxLength)
        } else {
            $line += "".PadRight($col2MaxLength)
        }
        $line += "    "

        if ($i -lt $programColumns[2].Count) {
            $line += $programColumns[2][$i].PadRight($col3MaxLength)
        }

        $menuLines.Add($line.TrimEnd())
    }

    $menuLines.Add("")

    # Select an Option section
    $optionsHeader = "Select an Option:"
    $optionsUnderline = "-------------------------------------------------------------------"

    $centeredOptionsHeader = (" " * $fixedHeaderPadding) + $optionsHeader

    $menuLines.Add($centeredOptionsHeader)
    $menuLines.Add($optionsUnderline)

    [string[]]$optionLines = @(
        "                  [A] Install All Programs                   ",
        "                  [G] Select Specific Programs via GUI       ",
        "                  [W] Activate Windows                       ",
        "                  [N] Update Windows                         ",
        "                  [E] Exit Script                            "
    )

    $menuLines.AddRange($optionLines)

    # Calculate overall menu block centering in console
    $consoleWidth = $Host.UI.RawUI.WindowSize.Width
    $blockPaddingValue = [math]::Floor(($consoleWidth - $fixedMenuWidth) / 2)
    if ($blockPaddingValue -lt 0) { $blockPaddingValue = 0 }
    $blockPaddingString = " " * $blockPaddingValue

    # Print all menu lines
    foreach ($lineEntry in $menuLines) {
        if ($lineEntry.TrimStart() -eq $pssText -or
            $lineEntry.TrimStart() -like ("=" * $lineEntry.TrimStart().Length) -or
            $lineEntry.TrimStart() -eq $programHeader -or
            $lineEntry.TrimStart() -eq $optionsHeader -or
            ($lineEntry.Replace("-", "").Trim().Length -eq 0 -and $lineEntry.Trim().Length -gt 0)) {
            Write-Host ($blockPaddingString + $lineEntry) -ForegroundColor Cyan
        } elseif ($lineEntry.TrimStart() -match "^\[(a|g|w|n|e)\]") {
            $fullLineWithoutBlockPadding = $lineEntry
            $leadingSpacesMatch = $fullLineWithoutBlockPadding | Select-String -Pattern "\S"
            $leadingSpaces = 0
            if ($leadingSpacesMatch) { $leadingSpaces = $leadingSpacesMatch.Matches[0].Index }

            $trimmedLineForColor = $fullLineWithoutBlockPadding.Substring($leadingSpaces)

            Write-Host ($blockPaddingString + (" " * $leadingSpaces) + "[") -NoNewline

            $letter = $Matches[1]
            $restOfLine = $trimmedLineForColor.Substring(3)

            switch ($letter) {
                "a" { Write-Host $letter -ForegroundColor Green -NoNewline }
                "g" { Write-Host $letter -ForegroundColor Yellow -NoNewline }
                "w" { Write-Host $letter -ForegroundColor Magenta -NoNewline }
                "n" { Write-Host $letter -ForegroundColor DarkCyan -NoNewline }
                "e" { Write-Host $letter -ForegroundColor Red -NoNewline }
                default { Write-Host $letter -ForegroundColor White -NoNewline }
            }
            Write-Host "]" -NoNewline
            Write-Host $restOfLine -ForegroundColor White
        } else {
            Write-Host ($blockPaddingString + $lineEntry) -ForegroundColor White
        }
    }

    Write-Host "" # Consolidate newlines after menu content

    # Define the default single-line dynamic prompt
    $dynamicPromptTextSingleLine = "Enter option [A-E], program number(s) [1-$($script:sortedPrograms.Count)] (e.g. 1 5 17), or single #"

    # Check if the dynamic prompt is too long for the fixed width
    if ($dynamicPromptTextSingleLine.Length -gt ($fixedMenuWidth - 4)) {
        $promptLine1Text = "Enter an option [A-E], a single number [1],"
        $promptLine2Text = "or a list of numbers [1 5 17]:"

        $promptPadding1 = [math]::Floor(($fixedMenuWidth - $promptLine1Text.Length) / 2)
        if ($promptPadding1 -lt 0) { $promptPadding1 = 0 }
        $centeredPromptLine1 = (" " * $promptPadding1) + $promptLine1Text

        $promptPadding2 = [math]::Floor(($fixedMenuWidth - $promptLine2Text.Length) / 2)
        if ($promptPadding2 -lt 0) { $promptPadding2 = 0 }
        $centeredPromptLine2 = (" " * $promptPadding2) + $promptLine2Text

        Write-Host ($blockPaddingString + $centeredPromptLine1) -ForegroundColor Yellow
        Write-Host ($blockPaddingString + $centeredPromptLine2) -NoNewline -ForegroundColor Yellow
    } else {
        $promptTextForOneLine = $dynamicPromptTextSingleLine

        $promptPaddingOneLine = [math]::Floor(($fixedMenuWidth - $promptTextForOneLine.Length) / 2)
        if ($promptPaddingOneLine -lt 0) { $promptPaddingOneLine = 0 }
        $centeredPromptOneLine = (" " * $promptPaddingOneLine) + $promptTextForOneLine

        Write-Host ($blockPaddingString + $centeredPromptOneLine) -NoNewline -ForegroundColor Yellow
    }
}

#endregion

#region Chocolatey Initialization

# Check Chocolatey installation
Write-LogAndHost "Checking Chocolatey installation..."
try {
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        if (-not (Install-Chocolatey)) {
            Write-LogAndHost "ERROR: Chocolatey is required to proceed. Exiting script." -HostColor Red
            exit 1
        }
        # Re-check if choco is now available
        if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
            Write-LogAndHost "ERROR: Chocolatey command still not found after installation attempt." -HostColor Red
            exit 1
        }
    }
    $chocoVersion = & choco --version 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-LogAndHost "ERROR: Chocolatey is not installed or not functioning correctly. Exit code: $LASTEXITCODE" -HostColor Red
        exit 1
    }
    Write-LogAndHost "Found Chocolatey version: $chocoVersion"
}
catch {
    Write-LogAndHost "ERROR: Exception occurred while checking Chocolatey. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during Chocolatey check - "
    exit 1
}
Write-Host ""

# Clear Chocolatey cache
Write-LogAndHost "Clearing Chocolatey cache..."
try {
    & choco cache remove 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8
    if ($LASTEXITCODE -eq 0) {
        Write-LogAndHost "Cache cleared."
    } else {
        Write-LogAndHost "ERROR: Failed to clear Chocolatey cache." -HostColor Red
    }
}
catch {
    Write-LogAndHost "ERROR: Exception occurred while clearing Chocolatey cache. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during cache clearing - "
}
Write-Host ""

# Enable automatic confirmation
Write-LogAndHost "Enabling automatic confirmation..."
try {
    & choco feature enable -n allowGlobalConfirmation 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8
    if ($LASTEXITCODE -eq 0) {
        Write-LogAndHost "Automatic confirmation enabled."
    } else {
        Write-LogAndHost "ERROR: Failed to enable automatic confirmation." -HostColor Red
    }
}
catch {
    Write-LogAndHost "ERROR: Exception occurred while enabling automatic confirmation. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during enabling automatic confirmation - "
}
Write-Host ""

#endregion

#region Main Loop

# Main loop for user selection
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

    # Check for empty input
    if ([string]::IsNullOrEmpty($userInput)) {
        Clear-Host
        Write-LogAndHost "No input detected. Please enter an option or program number(s)." -HostColor Yellow
        Write-LogAndHost "Press any key to return to the menu..." -HostColor DarkGray -NoLog
        $null = Read-Host # Wait for any input
        continue # Return to displaying the menu
    }

    # Validate input format
    if ($userInput -and $userInput -notmatch '^[aegnw0-9\s,]+$') {
        Clear-Host
        Write-LogAndHost "Invalid input: '$userInput'. Use only letters [A-E], numbers [1-$($script:sortedPrograms.Count)], spaces, or commas." -HostColor Red -LogPrefix "Invalid user input: '$userInput' - contains invalid characters."
        Start-Sleep -Seconds 2
        continue
    }

    if ($script:mainMenuLetters -contains $userInput) {
        switch ($userInput) {
            'e' {
                if (-not $script:activationAttempted) {
                    Write-LogAndHost "Exiting script. Windows activation not attempted by user choice on this run." -NoHost
                }
                Write-LogAndHost "Exiting script..."

                Write-LogAndHost "Disabling automatic confirmation..."
                try {
                    & choco feature disable -n allowGlobalConfirmation 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8
                    if ($LASTEXITCODE -eq 0) { Write-LogAndHost "Automatic confirmation disabled." }
                } catch {
                    Write-LogAndHost "ERROR: Exception during disabling automatic confirmation - $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during disabling automatic confirmation - "
                }
                Write-Host ""

                Write-LogAndHost "Checking installed programs..."
                try {
                    & choco list 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8
                    if ($LASTEXITCODE -eq 0) { Write-LogAndHost "Installed programs list (at exit) saved to $($script:logFile)" }
                } catch {
                    Write-LogAndHost "ERROR: Exception during listing installed programs - $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during listing installed programs - "
                }
                Write-Host ""
                exit 0
            }
            'a' {
                Clear-Host
                try {
                    Write-LogAndHost "Are you sure you want to install all programs? (Type y/n then press Enter)" -HostColor Yellow -NoLog
                    $confirmInput = Read-Host
                } catch {
                    Write-LogAndHost "ERROR: Could not read user input. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Failed to read user input for 'Install All Programs' - "
                    Start-Sleep -Seconds 2
                    continue
                }
                if ($confirmInput.Trim().ToLower() -eq 'y') {
                    Write-LogAndHost "User chose to install all programs." -NoHost
                    Clear-Host
                    if (Install-Programs -ProgramsToInstall $script:sortedPrograms) {
                        Write-LogAndHost "All programs installed successfully. Perdanga Forever."
                    } else {
                        Write-LogAndHost "Some programs may not have installed correctly. Check log: $($script:logFile)" -HostColor Yellow
                    }
                    Write-Host ""
                } else {
                    Write-LogAndHost "Installation of all programs cancelled."
                }
            }
            'g' {
                if ($script:guiAvailable) {
                    Write-LogAndHost "User chose GUI-based selection." -NoHost
                    $form = New-Object System.Windows.Forms.Form
                    $form.Text = "Perdanga GUI"
                    $form.Size = New-Object System.Drawing.Size(400, 450)
                    $form.StartPosition = "CenterScreen"
                    $form.FormBorderStyle = "FixedDialog"
                    $form.MaximizeBox = $false
                    $form.BackColor = [System.Drawing.Color]::FromArgb(0, 30, 60)
                    $panel = New-Object System.Windows.Forms.Panel
                    $panel.Size = New-Object System.Drawing.Size(360, 350)
                    $panel.Location = New-Object System.Drawing.Point(10, 10)
                    $panel.AutoScroll = $true
                    $panel.BackColor = [System.Drawing.Color]::FromArgb(0, 30, 60)
                    $form.Controls.Add($panel)
                    $checkboxes = @()
                    $yPos = 10
                    for ($i = 0; $i -lt $script:sortedPrograms.Length; $i++) {
                        $progName = $script:sortedPrograms[$i]
                        $dispNumber = $script:programToNumberMap[$progName]
                        $displayText = "$($dispNumber). $progName".PadRight(30) # Pad to align checkboxes
                        $checkbox = New-Object System.Windows.Forms.CheckBox
                        $checkbox.Text = $displayText
                        $checkbox.Location = New-Object System.Drawing.Point(10, $yPos)
                        $checkbox.Size = New-Object System.Drawing.Size(330, 24)
                        $checkbox.Font = New-Object System.Drawing.Font("Arial", 12)
                        $checkbox.ForeColor = [System.Drawing.Color]::White
                        $checkbox.BackColor = [System.Drawing.Color]::FromArgb(0, 30, 60)
                        $panel.Controls.Add($checkbox)
                        $checkboxes += $checkbox
                        $yPos += 28
                    }
                    $okButton = New-Object System.Windows.Forms.Button
                    $okButton.Text = "Install Selected"
                    $okButton.Location = New-Object System.Drawing.Point(140, 370)
                    $okButton.Size = New-Object System.Drawing.Size(120, 30)
                    $okButton.Font = New-Object System.Drawing.Font("Arial", 10)
                    $okButton.ForeColor = [System.Drawing.Color]::White
                    $okButton.BackColor = [System.Drawing.Color]::Gray
                    $okButton.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
                    $form.Controls.Add($okButton)
                    Clear-Host
                    $result = $form.ShowDialog()
                    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                        $selectedProgramsFromGui = @()
                        for ($i = 0; $i -lt $checkboxes.Length; $i++) {
                            if ($checkboxes[$i].Checked) { $selectedProgramsFromGui += $script:sortedPrograms[$i] }
                        }
                        if ($selectedProgramsFromGui.Count -eq 0) {
                            Write-LogAndHost "No programs selected via GUI."
                        }
                        else {
                            Clear-Host
                            if (Install-Programs -ProgramsToInstall $selectedProgramsFromGui) {
                                Write-LogAndHost "Selected programs installed. Perdanga Forever."
                            } else {
                                Write-LogAndHost "Some GUI selected programs failed." -HostColor Yellow
                            }
                        }
                    }
                } else {
                    Clear-Host
                    Write-LogAndHost "ERROR: GUI selection (g) is not available." -HostColor Red
                }
            }
            'w' {
                Clear-Host
                Invoke-WindowsActivation
                Write-LogAndHost "User chose to activate Windows." -NoHost
            }
            'n' {
                Clear-Host
                try {
                    Write-LogAndHost "Are you sure you want to update Windows? (Type y/n then press Enter)" -HostColor Yellow -NoLog
                    $confirmInput = Read-Host
                } catch {
                    Write-LogAndHost "ERROR: Could not read user input. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Failed to read user input for 'Update Windows' - "
                    Start-Sleep -Seconds 2
                    continue
                }
                if ($confirmInput.Trim().ToLower() -eq 'y') {
                    Clear-Host
                    Invoke-WindowsUpdate
                    Write-LogAndHost "User chose to update Windows." -NoHost
                } else {
                    Write-LogAndHost "Windows update cancelled."
                }
            }
        }
    }
    elseif ($userInput -match '[, ]+') {
        Clear-Host
        $selectedIndividualInputs = $userInput -split '[, ]+' | ForEach-Object { $_.Trim() }
        $validProgramNamesToInstall = New-Object System.Collections.Generic.List[string]
        $invalidNumbersInList = New-Object System.Collections.Generic.List[string]

        foreach ($inputNumStr in $selectedIndividualInputs) {
            if ($inputNumStr -match '^\d+$' -and $script:numberToProgramMap.ContainsKey($inputNumStr)) {
                $programName = $script:numberToProgramMap[$inputNumStr]
                if (-not $validProgramNamesToInstall.Contains($programName)) {
                    $validProgramNamesToInstall.Add($programName)
                }
            } elseif ($inputNumStr -ne "") {
                $invalidNumbersInList.Add($inputNumStr)
            }
        }

        if ($validProgramNamesToInstall.Count -eq 0) {
            Write-LogAndHost "No valid program numbers found in your input: '$userInput'." -HostColor Red
            if ($invalidNumbersInList.Count -gt 0) {
                Write-LogAndHost "The following inputs were unrecognized or invalid: $($invalidNumbersInList -join ', ')" -HostColor Red -NoLog
            }
            Write-LogAndHost "User entered list '$userInput', but no valid programs found. Invalid inputs: $($invalidNumbersInList -join ', ')" -NoHost -LogPrefix ""
            Start-Sleep -Seconds 2
        } else {
            Write-LogAndHost "Based on your input '$userInput', the following program(s) are selected for installation:" -HostColor Cyan -NoLog
            $validProgramNamesToInstall | ForEach-Object { Write-LogAndHost "- $_" -NoLog }
            Write-Host ""

            if ($invalidNumbersInList.Count -gt 0) {
                Write-LogAndHost "Note: The following inputs were invalid and will be skipped: $($invalidNumbersInList -join ', ')" -HostColor Yellow -NoLog
                Write-LogAndHost "User input for multiple programs: '$userInput'. Valid programs: $($validProgramNamesToInstall -join ', '). Invalid inputs: $($invalidNumbersInList -join ', ')" -NoHost -LogPrefix ""
            } else {
                Write-LogAndHost "User input for multiple programs: '$userInput'. Valid programs: $($validProgramNamesToInstall -join ', ')" -NoHost -LogPrefix ""
            }
            try {
                Write-LogAndHost "Do you want to proceed with installing these $($validProgramNamesToInstall.Count) program(s)? (Type y/n then press Enter)" -HostColor Yellow -NoLog
                $confirmMultiInput = Read-Host
            } catch {
                Write-LogAndHost "ERROR: Could not read user input. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Failed to read user input for multiple programs - "
                Start-Sleep -Seconds 2
                continue
            }
            if ($confirmMultiInput.Trim().ToLower() -eq 'y') {
                Write-LogAndHost "User confirmed installation of $($validProgramNamesToInstall.Count) programs from list." -NoHost
                Clear-Host
                if (Install-Programs -ProgramsToInstall $validProgramNamesToInstall) {
                    Write-LogAndHost "All selected programs from your list installed successfully. Perdanga Forever."
                } else {
                    Write-LogAndHost "Some programs from your list may not have installed correctly. Check log: $($script:logFile)" -HostColor Yellow
                }
                Write-Host ""
            } else {
                Write-LogAndHost "Installation of selected programs from list cancelled."
            }
        }
    }
    elseif ($script:numberToProgramMap.ContainsKey($userInput)) {
        $programToInstall = $script:numberToProgramMap[$userInput]
        Clear-Host
        try {
            Write-LogAndHost "Are you sure you want to install '$($programToInstall)' (program #$($userInput))? (Type y/n then press Enter)" -HostColor Yellow -NoLog
            $confirmSingleInput = Read-Host
        } catch {
            Write-LogAndHost "ERROR: Could not read user input. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Failed to read user input for single program - "
            Start-Sleep -Seconds 2
            continue
        }
        if ($confirmSingleInput.Trim().ToLower() -eq 'y') {
            Clear-Host
            Write-LogAndHost "User selected program #'$($userInput)' ($($programToInstall)) for installation and confirmed." -NoHost
            if (Install-Programs -ProgramsToInstall @($programToInstall)) {
                Write-LogAndHost "$($programToInstall) installed successfully. Perdanga Forever."
            } else {
                Write-LogAndHost "Failed to install $($programToInstall). Check log: $($script:logFile)" -HostColor Red
            }
            Write-Host ""
        } else {
            Write-LogAndHost "Installation of '$($programToInstall)' cancelled."
        }
    }
    else {
        Clear-Host
        Write-LogAndHost "Invalid selection: '$($userInput)'. Please enter a valid option [A-E], program number [1-$($script:sortedPrograms.Count)], or list of numbers." -HostColor Red -LogPrefix "Invalid user input: '$($userInput)'"
        Start-Sleep -Seconds 2
    }

} while ($true)

#endregion
