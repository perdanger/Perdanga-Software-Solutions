<#
.SYNOPSIS
    Configuration file for Perdanga Software Solutions.
#>

# --- PROGRAM DEFINITIONS ---
# Edit this list to add or remove programs.
# Use the exact Chocolatey package ID.
$programs = @(
    "7zip.install",
    "brave",
    "file-converter",
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
    "wiztree",
    "discord",
    "git"
)

# --- SCRIPT-WIDE VARIABLES ---
# The following variables are set in the 'script' scope to be accessible
# throughout the main script and all functions.

# Sort the program list for consistent display
$script:sortedPrograms = $programs | Sort-Object

# Define the letters for main menu commands
# ADDED 't' for the new Telemetry option
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
