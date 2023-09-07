#
# Script.ps1
#
# Author: ORLANDO R HERNANDEZ
# Script Version: 1.0.1

# Must do:
#	The installer executable must contain the year of revit it is for at the end of the name. 
#    Example. "BIMPro 2020.exe" 

#### What this script does ####
# 1. Checks if PowerShell is running in 32-bit mode on a 64-bit machine and restarts it in 64-bit mode if necessary.
# 2. Defines functions for various tasks, like finding the current BIMPro installer, extracting the Revit year from the installer's filename, and getting the version of the BIMPro installer.
# 3. Sets up logging for the script's activities.
# 4. Uninstalls the previous version of BIMPro if it exists.
# 5. Installs the new version of BIMPro.
# 6. Copies the script's log file to a specific folder.

#############################################################################
#If Powershell is running the 32-bit version on a 64-bit machine, we 
#need to force powershell to run in 64-bit mode .
#############################################################################
If ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    Try {
		write-Output "Swithing to 64 Bit"
        write-Output "$PSCOMMANDPATH"
		Start-Sleep -Seconds 2
		
        & "$ENV:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -NoProfile -File $PSCOMMANDPATH
    }
    Catch {
        Throw "Failed to start $PSCOMMANDPATH"
    }
}
#############################################################################

function GetCurrenBIMProInstaller
{ 
    # Return the path if file is found else exit script with code 777
    $BIMProExe = (Get-Item "$PSScriptRoot\BIMPro*.exe" -ErrorAction SilentlyContinue -Force).FullName
    if(Test-Path $BIMProExe -ErrorAction SilentlyContinue)
    {return $BIMProExe}
    else
    {
        Write-Output "Could not fine the BIMPro Installer file. `n Exiting..."
        Start-Sleep -Seconds 5
        EXIT 777
    }
}

function GetTheRevitYearFromTheInstallerFileName
{
    Param($FilePath)
    if($FilePath)
    {
        # Split the file path by backslashes to get an array of components
    $PathComponents = $FilePath -split '\\'
    
    # Get the last component (the file name)
    $FileName = $PathComponents[-1]
    
    # Use regex to match the last set of four consecutive digits before a period
    if ($FileName -match '(\d{4})\.(?=[^.]*$)')
    {
        # Extract the matched four-digit number
        $Year = $Matches[1]
    }
    
    # Output the extracted year
    return $Year
    }
}

function GetTheVersionOfTheBIMProInstaller
{
    Param($FilePath)
    $InstallerVersion = (Get-Item $FilePath).VersionInfo.ProductVersion
    return $InstallerVersion
}

$TimeString = (get-date).ToString("mmddyyyy_HHmmss")
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path -Parent -Path $script:MyInvocation.MyCommand.Path

$Log = "$env:windir\Temp\Log_$ScriptName`_$TimeString.txt"
Start-Transcript $Log
$InstallerPath = GetCurrenBIMProInstaller   # The funtion 'GetCurrenBIMProInstaller' gets the full path for the installer
$RevitYear = $(GetTheRevitYearFromTheInstallerFileName $InstallerPath)    # runs a function to get the year
$BIMProInstallerVersion = GetTheVersionOfTheBIMProInstaller $InstallerPath # Gets the version from the installer exe "$InstallerPath"

$BIMProInstaller = "`"$InstallerPath`" /SILENT /NORESTART /LOG"    
$BIMProUninstaller = "C:\ProgramData\Autodesk\Revit\Addins\$RevitYear\unins000.exe"  # The year is the variable $RevitYear

# Uninstall previous version if it exists 
if (Test-Path $BIMProUninstaller -ErrorAction SilentlyContinue)
{
    Write-Output "Uninstalling Previous version"
    & cmd /c $BIMProUninstaller /SILENT
}

Start-Sleep -Seconds 3  # Wait 3 seconds before installing

#install BIMPro
Write-Output "Installing BIMPro $RevitYear v$BIMProInstallerVersion"
& cmd /c $BIMProInstaller

Write-Output "Done!"

Stop-Transcript

## Copy logs to Contech Logs folder
$logsFolder = "C:\ConTech\LogFiles"
if(-not(Test-Path $logsFolder))
{New-Item -Path $logsFolder -ItemType Directory}
Start-Sleep -Seconds 2
Copy-Item $Log "$logsFolder\"