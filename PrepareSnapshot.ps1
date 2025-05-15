<#
.SYNOPSIS
    This script backs up specific files to an external drive, creating a timestamped folder for each backup.
    It checks if Outlook is running and prevents the backup if it is.
.DESCRIPTION
    The script performs the following actions:
    1. Checks if Microsoft Outlook is running. If it is, the script displays an error message and exits.
    2. Defines the source file paths and the destination path (on the external drive).
    3. Creates a new directory on the external drive named "OutlookBackups" (if it doesn't exist).
    4. Creates a new subdirectory within "OutlookBackups" using the current date and time as the name.
    5. Copies the specified files to the newly created timestamped directory.
    6. Displays a success message with the destination path.
.PARAMETER ExternalDrive
    The drive letter of the external drive to copy the files to.  If not provided, the script will attempt to use the first removable drive it finds.
.NOTES
    * The script assumes that the user has the necessary permissions to access the source files and write to the destination drive.
    * The script uses the current date and time to create the timestamped directory.
    * Error handling is included to catch potential issues during file copying and directory creation.
    * The script is designed to be run on a Windows system.
#>
param (
    [Parameter(Mandatory = $false, HelpMessage = "The drive letter of the external drive.")]
    [string]$ExternalDrive
)

#region Script Constants
# Define the files to be backed up.  These are hardcoded, as requested.
$FilesToBackup = @(
    "$env:USERPROFILE\AppData\Local\Microsoft\Outlook\outlook.pst",
    "$env:USERPROFILE\Documents\MyFile1.docx",
    "$env:USERPROFILE\Documents\MyFile2.xlsx",
    "$env:USERPROFILE\AppData\Roaming\Microsoft\Signatures" # Added a folder.
)

# Base name for the backup directory on the external drive
$BackupBaseDir = "OutlookBackups"
#endregion

#region Helper Functions

function Test-OutlookRunning {
    <#
    .SYNOPSIS
        Checks if Microsoft Outlook is running.
    .DESCRIPTION
        This function checks if the Outlook.exe process is running.
    .RETURNS
        $true if Outlook is running, $false otherwise.
    #>
    try {
        # Use Get-Process to check for Outlook.  This is more robust.
        $OutlookProcess = Get-Process -Name "outlook" -ErrorAction Stop
        return $true # Outlook process was found.
    }
    catch {
        # Get-Process threw an error, meaning Outlook is not running.
        return $false
    }
}

function Get-ExternalDrive {
    <#
    .SYNOPSIS
        Gets the drive letter of the external drive.
    .DESCRIPTION
        This function retrieves the drive letter of the first removable drive connected to the system.
    .RETURNS
        The drive letter of the external drive as a string, or $null if no external drive is found.
    #>
    if ($ExternalDrive) {
        #If the external drive is provided as a parameter, use it.
        if (Test-Path -Path "$ExternalDrive:") {
            return $ExternalDrive
        }
        else {
            Write-Warning "The specified external drive '$ExternalDrive' was not found. Attempting to find a removable drive automatically..."
        }
    }
    try {
        # Get all removable drives.  This should work even if no drive letter is assigned.
        $Drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 2 }

        if ($Drives) {
            #Return the first removable drive.
            return $Drives[0].DeviceID.Substring(0, 2) #Extracts the drive letter (e.g., "E:")
        }
        else {
            Write-Warning "No removable drive found."
            return $null
        }
    }
    catch {
        Write-Warning "Error getting external drive: $($_.Exception.Message)"
        return $null
    }
}

function Copy-FilesToBackup {
    <#
    .SYNOPSIS
        Copies the specified files to the destination directory.
    .DESCRIPTION
        This function copies the files defined in $FilesToBackup to the specified destination directory.
    .PARAMETER SourceFiles
        An array of file paths to copy.
    .PARAMETER DestinationDir
        The directory to copy the files to.
    .PARAMETER ShouldContinue
        A boolean that controls if the copy should proceed
    .RETURNS
        $true if the files were copied successfully, $false otherwise.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$SourceFiles,
        [Parameter(Mandatory = $true)]
        [string]$DestinationDir,
        [Parameter(Mandatory = $true)]
        [bool]$ShouldContinue
    )
    if (-not $ShouldContinue)
    {
        return $false
    }

    try {

        # Create the destination directory if it doesn't exist.
        if (-not (Test-Path -Path $DestinationDir -PathType 'Container')) {
            Write-Verbose "Creating directory: $DestinationDir"
            New-Item -Path $DestinationDir -ItemType 'Directory' -Force | Out-Null # Use -Force to suppress errors if it already exists
        }

        # Copy each file.
        foreach ($SourceFile in $SourceFiles) {
            if (Test-Path -Path $SourceFile) {
                $FileName = Split-Path -Path $SourceFile -Leaf
                $DestinationPath = Join-Path -Path $DestinationDir -ChildPath $FileName
                Write-Verbose "Copying '$SourceFile' to '$DestinationPath'"
                Copy-Item -Path $SourceFile -Destination $DestinationPath -Force -ErrorAction Stop
                Write-Verbose "Successfully copied '$SourceFile' to '$DestinationPath'"
            }
            else {
                Write-Warning "Source file '$SourceFile' not found. Skipping."
            }
        }
        return $true
    }
    catch {
        Write-Error "Error copying files: $($_.Exception.Message)"
        return $false
    }
}

#endregion

#region Main Script Logic
# Check if Outlook is running.
if (Test-OutlookRunning) {
    Write-Error "Outlook is running. Please close Outlook before running this script."
    exit  # Stop the script execution.
}

# Get the external drive letter.
$ExternalDriveLetter = Get-ExternalDrive

if (-not $ExternalDriveLetter) {
    Write-Error "No external drive found. Please connect an external drive and run the script again."
    exit # Stop the script if no external drive is found.
}

# Construct the full path to the backup directory.
$BackupDir = Join-Path -Path "$ExternalDriveLetter:" -ChildPath $BackupBaseDir

# Create a timestamped subdirectory.
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$DestinationDir = Join-Path -Path $BackupDir -ChildPath $Timestamp

# Copy the files.
$ShouldContinue = $true;
$CopyResult = Copy-FilesToBackup -SourceFiles $FilesToBackup -DestinationDir $DestinationDir -ShouldContinue $ShouldContinue

if ($CopyResult) {
    Write-Host "Files backed up successfully to: $DestinationDir" -ForegroundColor Green
}
else
{
    Write-Error "File backup failed."
}

#endregion
