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
    Write-Host "WARNING: Could not set console buffer or window size. This may happen in some environments (e.g., VS Code integrated terminal).\" -ForegroundColor Yellow
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor DarkYellow
}

# Set log file name with timestamp and use script directory
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$script:logFile = Join-Path -Path $PSScriptRoot -ChildPath "install_log_$timestamp.txt"

# Define programs to install
$programs = @("7zip.install", "brave", "file-converter", "googlechrome", "gpu-z", "hwmonitor", "imageglass", "nvidia-app", "obs-studio", "occt", "qbittorrent", "revo-uninstaller", "spotify", "steam", "telegram", "vcredist-all", "vlc", "winrar", "wiztree", "discord", "git")
$script:sortedPrograms = $programs | Sort-Object

# Assign numbers to programs, excluding main menu command letters
# Added 'u' for Uninstall GUI, 'c' for Custom Package
$script:mainMenuLetters = @('a', 'e', 'g', 'n', 'w', 'u', 'c') # Lowercase main command letters
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

# Function to get installed Chocolatey packages
function Get-InstalledChocolateyPackages {
    $chocoLibPath = "C:\ProgramData\chocolatey\lib"
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

# Function to test if a Chocolatey package exists
function Test-ChocolateyPackage {
    param (
        [string]$PackageName
    )
    Write-LogAndHost "Searching for package '$PackageName' in Chocolatey repository..." -NoLog
    try {
        $tempStdoutFile = Join-Path -Path $env:TEMP -ChildPath "choco_search_stdout_$(New-Guid).txt"
        $tempStderrFile = Join-Path -Path $env:TEMP -ChildPath "choco_search_stderr_$(New-Guid).txt"

        $searchArguments = @(
            "search",
            $PackageName,
            "--exact",
            "--limit-output",
            "--source=https://community.chocolatey.org/api/v2/",
            "--no-progress" # Suppress progress bar
        )
        
        $chocoCmd = Get-Command choco -ErrorAction SilentlyContinue
        if (-not $chocoCmd) {
            Write-LogAndHost "ERROR: Chocolatey command (choco) not found. Cannot search for package." -HostColor Red
            return $false
        }

        $process = Start-Process -FilePath $chocoCmd.Source -ArgumentList $searchArguments -Wait -NoNewWindow -PassThru -RedirectStandardOutput $tempStdoutFile -RedirectStandardError $tempStderrFile
        $searchExitCode = $process.ExitCode
        
        $stdout = Get-Content $tempStdoutFile -ErrorAction SilentlyContinue
        $stderr = Get-Content $tempStderrFile -ErrorAction SilentlyContinue
        
        Remove-Item $tempStdoutFile -ErrorAction SilentlyContinue
        Remove-Item $tempStderrFile -ErrorAction SilentlyContinue

        Write-LogAndHost "choco search for '$PackageName' exited with code $searchExitCode." -NoHost
        if ($stdout) { Write-LogAndHost "stdout: $($stdout | Out-String)" -NoHost }
        if ($stderr) { Write-LogAndHost "stderr: $($stderr | Out-String)" -NoHost }

        if ($searchExitCode -ne 0) {
            Write-LogAndHost "Error during 'choco search' for '$PackageName'. Exit code: $searchExitCode. Stderr: $($stderr | Out-String)" -HostColor Red
            return $false
        }

        $trimmedStdout = if ($stdout) { ($stdout | ForEach-Object {$_.Trim()}) -join "`n" } else { "" }

        # MODIFIED CONDITION: Check if stdout is exactly the package name OR starts with packagename|
        if (-not [string]::IsNullOrWhiteSpace($trimmedStdout) -and 
            ( ($trimmedStdout -ieq $PackageName) -or ($trimmedStdout -ilike "$($PackageName)|*") )
           ) {
            Write-LogAndHost "Package '$PackageName' found in repository (Search result: '$trimmedStdout')." -HostColor Green
            return $true
        } else {
            Write-LogAndHost "Package '$PackageName' not found as an exact match or prefix in Chocolatey repository. Expected '$PackageName' or '$($PackageName)|*', got '$trimmedStdout'." -HostColor Yellow
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
    $fixedMenuWidth = 67
    $pssText = "Perdanga Software Solutions"
    $pssUnderline = "==================================================================="
    $dashedLine = "-------------------------------------------------------------------"
    $fixedHeaderPadding = 22 
    $centeredPssTextLine = (" " * $fixedHeaderPadding) + $pssText

    $menuLines.Add($pssUnderline)
    $menuLines.Add($centeredPssTextLine)
    $menuLines.Add($dashedLine)
    $menuLines.Add(" Chocolatey Package Manager [PSS v1.1] ($(Get-Date -Format "dd.MM.yyyy HH:mm"))")
    $menuLines.Add(" Log saved to: install_log_$timestamp.txt")
    $menuLines.Add(" $($script:sortedPrograms.Count) programs available for installation")
    $menuLines.Add($pssUnderline)
    $menuLines.Add("")

    $programHeader = "Available Programs for Installation:"
    $centeredProgramHeader = (" " * (($fixedMenuWidth - $programHeader.Length) / 2)) + $programHeader
    $programUnderline = "-------------------------------------------------------------------"
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
    $col3MaxLength = ($programColumns[2] | Measure-Object -Property Length -Maximum).Maximum
    if ($col1MaxLength -eq $null) {$col1MaxLength = 0}
    if ($col2MaxLength -eq $null) {$col2MaxLength = 0}
    if ($col3MaxLength -eq $null) {$col3MaxLength = 0}

    $maxRows = [math]::Max($programColumns[0].Count, [math]::Max($programColumns[1].Count, $programColumns[2].Count))

    for ($i = 0; $i -lt $maxRows; $i++) {
        $line = "  "
        if ($i -lt $programColumns[0].Count) { $line += $programColumns[0][$i].PadRight($col1MaxLength) } 
        else { $line += "".PadRight($col1MaxLength) }
        $line += "    " 
        if ($i -lt $programColumns[1].Count) { $line += $programColumns[1][$i].PadRight($col2MaxLength) }
        else { $line += "".PadRight($col2MaxLength) }
        $line += "    " 
        if ($i -lt $programColumns[2].Count) { $line += $programColumns[2][$i].PadRight($col3MaxLength) }
        $menuLines.Add($line.TrimEnd())
    }
    $menuLines.Add("")

    $optionsHeader = "Select an Option:"
    $optionsUnderline = "-------------------------------------------------------------------"
    $centeredOptionsHeader = (" " * (($fixedMenuWidth - $optionsHeader.Length) / 2)) + $optionsHeader
    $menuLines.Add($centeredOptionsHeader)
    $menuLines.Add($optionsUnderline)

    [string[]]$optionLines = @( 
        "                  [A] Install All Programs                   ",
        "                  [G] Select Specific Programs via GUI       ",
        "                  [U] Uninstall Programs via GUI             ", 
        "                  [C] Install Custom Package                 ", # New Option
        "                  [W] Activate Windows                       ",
        "                  [N] Update Windows                         ",
        "                  [E] Exit Script                            "
    )
    $menuLines.AddRange($optionLines)

    $consoleWidth = $Host.UI.RawUI.WindowSize.Width
    $blockPaddingValue = [math]::Floor(($consoleWidth - $fixedMenuWidth) / 2)
    if ($blockPaddingValue -lt 0) { $blockPaddingValue = 0 }
    $blockPaddingString = " " * $blockPaddingValue

    foreach ($lineEntry in $menuLines) {
        if ($lineEntry.TrimStart() -eq $pssText -or
            $lineEntry.TrimStart() -like ("=" * $lineEntry.TrimStart().Length) -or
            $lineEntry.TrimStart() -eq $programHeader -or
            $lineEntry.TrimStart() -eq $optionsHeader -or
            ($lineEntry.Replace("-", "").Trim().Length -eq 0 -and $lineEntry.Trim().Length -gt 0)) {
            Write-Host ($blockPaddingString + $lineEntry) -ForegroundColor Cyan
        } elseif ($lineEntry.TrimStart() -match "^\[(a|g|w|n|e|u|c)\]") { # Added 'c'
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
                "u" { Write-Host $letter -ForegroundColor DarkRed -NoNewline } 
                "c" { Write-Host $letter -ForegroundColor Blue -NoNewline } # MODIFIED Color for Custom
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
    Write-Host "" 

    $dynamicPromptTextSingleLine = "Enter option 'A-E', single number '1', or list of numbers '1 5 17':" # Updated prompt
    if ($dynamicPromptTextSingleLine.Length -gt ($fixedMenuWidth - 4)) {
        $promptLine1Text = "Enter option 'A-E', single number '1'," # Updated
        $promptLine2Text = "or list of numbers '1 5 17':"
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
Write-LogAndHost "Checking Chocolatey installation..."
try {
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        if (-not (Install-Chocolatey)) {
            Write-LogAndHost "ERROR: Chocolatey is required to proceed. Exiting script." -HostColor Red
            exit 1
        }
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
Write-LogAndHost "Clearing Chocolatey cache..."
try {
    & choco cache remove --all 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8 # Added --all to clear all types of cache
    if ($LASTEXITCODE -eq 0) { Write-LogAndHost "Cache cleared." }
    else { Write-LogAndHost "ERROR: Failed to clear Chocolatey cache." -HostColor Red }
}
catch { Write-LogAndHost "ERROR: Exception occurred while clearing Chocolatey cache. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during cache clearing - " }
Write-Host ""
Write-LogAndHost "Enabling automatic confirmation..."
try {
    & choco feature enable -n allowGlobalConfirmation 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8
    if ($LASTEXITCODE -eq 0) { Write-LogAndHost "Automatic confirmation enabled." }
    else { Write-LogAndHost "ERROR: Failed to enable automatic confirmation." -HostColor Red }
}
catch { Write-LogAndHost "ERROR: Exception occurred while enabling automatic confirmation. $($_.Exception.Message)" -HostColor Red -LogPrefix "Error: Exception during enabling automatic confirmation - " }
Write-Host ""
#endregion

#region Main Loop
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

    if ($userInput -and $userInput -notmatch '^[aegnuwc0-9\s,]+$') { # Added 'c' to regex
        Clear-Host
        Write-LogAndHost "Invalid input: '$userInput'. Use only letters [A,G,U,C,W,N,E], numbers [1-$($script:sortedPrograms.Count)], spaces, or commas." -HostColor Red -LogPrefix "Invalid user input: '$userInput' - contains invalid characters."
        Start-Sleep -Seconds 2
        continue
    }

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
                    & choco list 2>&1 | Out-File -FilePath $script:logFile -Append -Encoding UTF8
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
                    if (Install-Programs -ProgramsToInstall $script:sortedPrograms) { Write-LogAndHost "All programs installed successfully. Perdanga Forever." }
                    else { Write-LogAndHost "Some programs may not have installed correctly. Check log: $($script:logFile)" -HostColor Yellow }
                    Write-Host ""
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
                        else { Clear-Host; if (Install-Programs -ProgramsToInstall $selectedProgramsFromGui) { Write-LogAndHost "Selected programs installed. Perdanga Forever." } else { Write-LogAndHost "Some GUI selected programs failed to install." -HostColor Yellow } }
                    } else { Write-LogAndHost "GUI installation cancelled by user." }
                } else { Clear-Host; Write-LogAndHost "ERROR: GUI selection (g) is not available." -HostColor Red }
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
                        else { Clear-Host; if (Uninstall-Programs -ProgramsToUninstall $selectedProgramsToUninstall) { Write-LogAndHost "Selected programs uninstalled. Perdanga Forever." } else { Write-LogAndHost "Some GUI selected programs failed to uninstall." -HostColor Yellow } }
                    } else { Write-LogAndHost "GUI uninstallation cancelled by user." }
                } else { Clear-Host; Write-LogAndHost "ERROR: GUI uninstallation (u) is not available." -HostColor Red }
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
                                Write-LogAndHost "Custom package '$customPackageName' installed successfully. Perdanga Forever."
                            } else {
                                Write-LogAndHost "Failed to install custom package '$customPackageName'. Check log: $($script:logFile)" -HostColor Red
                            }
                        } else {
                            Write-LogAndHost "Installation of custom package '$customPackageName' cancelled by user." -HostColor Yellow
                        }
                    } catch {
                         Write-LogAndHost "ERROR: Could not read user input for custom package installation confirmation. $($_.Exception.Message)" -HostColor Red
                         Start-Sleep -Seconds 2
                    }
                } else {
                    # Test-ChocolateyPackage already logs details of why it failed (not found, error, etc.)
                    Write-LogAndHost "Custom package '$customPackageName' could not be installed (either not found or validation failed)." -HostColor Red
                }
                Write-LogAndHost "Press any key to return to the menu..." -NoLog -HostColor DarkGray
                $null = Read-Host
            }
            'w' { Clear-Host; Invoke-WindowsActivation; Write-LogAndHost "User chose to activate Windows." -NoHost }
            'n' {
                Clear-Host
                try {
                    Write-LogAndHost "Are you sure you want to update Windows? (Type y/n then press Enter)" -HostColor Yellow -NoLog
                    $confirmInput = Read-Host
                } catch { Write-LogAndHost "ERROR: Could not read user input. $($_.Exception.Message)" -HostColor Red; Start-Sleep -Seconds 2; continue }
                if ($confirmInput.Trim().ToLower() -eq 'y') { Clear-Host; Invoke-WindowsUpdate; Write-LogAndHost "User chose to update Windows." -NoHost }
                else { Write-LogAndHost "Windows update cancelled." }
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
                if (Install-Programs -ProgramsToInstall $validProgramNamesToInstall) { Write-LogAndHost "Selected programs installed. Perdanga Forever." }
                else { Write-LogAndHost "Some selected programs failed. Check log: $($script:logFile)" -HostColor Yellow }
            } else { Write-LogAndHost "Installation of selected programs cancelled." }
        }
    }
    elseif ($script:numberToProgramMap.ContainsKey($userInput)) { 
        $programToInstall = $script:numberToProgramMap[$userInput]
        Clear-Host
        try {
            Write-LogAndHost "Install '$($programToInstall)' (program #$($userInput))? (Type y/n then press Enter)" -HostColor Yellow -NoLog
            $confirmSingleInput = Read-Host
        } catch { Write-LogAndHost "ERROR: Could not read user input. $($_.Exception.Message)" -HostColor Red; Start-Sleep -Seconds 2; continue }
        if ($confirmSingleInput.Trim().ToLower() -eq 'y') {
            Clear-Host
            if (Install-Programs -ProgramsToInstall @($programToInstall)) { Write-LogAndHost "$($programToInstall) installed. Perdanga Forever." }
            else { Write-LogAndHost "Failed to install $($programToInstall). Check log: $($script:logFile)" -HostColor Red }
        } else { Write-LogAndHost "Installation of '$($programToInstall)' cancelled." }
    }
    else {
        Clear-Host
        Write-LogAndHost "Invalid selection: '$($userInput)'. Use options [A,G,U,C,W,N,E] or program numbers." -HostColor Red # Updated message
        Start-Sleep -Seconds 2
    }
} while ($true)
#endregion
