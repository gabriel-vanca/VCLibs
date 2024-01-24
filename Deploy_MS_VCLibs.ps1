<#
.SYNOPSIS
    Installs and updates Microsoft VCLibs
.DESCRIPTION
	Installs Microsoft VCLibs if no VCLibs.14 is detected at all,
    or updates it if the subversion detected is too old.
    # This can only be installed in a user context (*-AppxPackage).
    # You cannot use *-AppxProvisionedPackage as it produced 'Element not found'.
    
    Deployment tested on:
        - Windows 10
        - Windows 11
        - Windows Sandbox
        - Windows Server 2019
        - Windows Server 2022
        - Windows Server 2022 vNext (Windows Server 2025)

    If necessary, Microsoft.UI.Xaml will be installed/updated as well.
.PARAMETER ForceReinstall
    (Optional)
    Forces reinstall.
    Typically you shouldn't need to activate this, unless you have problems with running
        an app due to the installed version being too old.
    You can't install an older version of an installed appx package, you'd need to remove it first.
    We also don't want to remove later versions as they will have been installed outside of the
        package ecosystem and may have been installed for very good reasons by another application.
    We also don't want to just remove the same version again either for similar reasons.
    ForceReinstall therefore forces uninstallation and (re)installation.
.EXAMPLE
    To use the default Chocolatey Community Repository, run this:
	    PS> ./Chocolatey_Deploy
    To use a local repository, run either of these:
        PS> ./Chocolatey_Deploy "http://10.10.10.1:8624/nuget/Thoth/" "THOTH"
.LINK
	https://github.com/gabrielvanca/VCLibs
.NOTES
	Author: Gabriel Vanca
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $False)] [Switch]$ForceReinstall = $False
)

[String]$VersionToLookFor = "14.0.30704.0"

#Requires -RunAsAdministrator

# Force use of TLS 1.2 for all downloads.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$windowsVersion = [Environment]::OSVersion.Version
if ($windowsVersion.Major -lt "10") {
    throw "This package requires a minimum of Windows 10 / Server 2019."
}

# .appx is only  supported on Windows 10 version 1709 and later - https://learn.microsoft.com/en-us/windows/msix/supported-platforms
# See https://en.wikipedia.org/wiki/Windows_10_version_history for build numbers
if ($windowsVersion.Build -lt "16299") {
    throw "This package requires a minimum of Windows 10 / Server 2019 version 1709 / OS build 16299."
}

if($ForceReinstall -eq $False) {
    $vclibsList = Get-AppxPackage Microsoft.VCLibs.140.00.UWPDesktop | Where-Object version -ge $VersionToLookFor
    if([string]::IsNullorEmpty($vclibsList)) {
        Write-Host "Microsoft.VCLibs.140.00.UWPDesktop missing" -ForegroundColor DarkRed
        Write-Warning "VCLibs package required for WinGet and Terminal installation."
    } else {
        Write-Host "The installed version of Microsoft.VCLibs.140.00.UWPDesktop is the same or newer as the version we are looking for." -ForegroundColor DarkGreen
        Start-Sleep -Seconds 1
        Return
    }
} else {
    Write-Host "FORCED (RE)INSTALL ENABLED." -ForegroundColor DarkYellow
    Start-Sleep -Seconds 7
    Write-Host "Skipping version check." -ForegroundColor DarkYellow
    Write-Host "Initialising uninstall." -ForegroundColor DarkYellow
    Write-Warning "If anything relies on VCLibs, even if not turned on, the uninstall will fail."
    Get-AppxPackage "Microsoft.VCLibs.140.00.UWPDesktop" | remove-AppxPackage -allusers
}


Write-Host "Installing VCLibs package" -ForegroundColor DarkYellow

# Downloading necessary graphical component, usually for Windows Server or Sandbox deployments
# (https://docs.microsoft.com/en-us/troubleshoot/developer/visualstudio/cpp/libraries/c-runtime-packages-desktop-bridge)
$WebClient = New-Object System.Net.WebClient
$fileURL = "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx"
$fileDownloadLocalPath = "$env:Temp\Microsoft.VCLibs.x64.14.00.Desktop.appx"
$WebClient.DownloadFile($fileURL, $fileDownloadLocalPath)

# Invoke-WebRequest -Uri https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx `
#                   -outfile $env:Temp\Microsoft.VCLibs.x64.14.00.Desktop.appx

# This will not work if a Windows Terminal window is already open
Write-Host "Any software that uses VCLibs should be closed." -ForegroundColor DarkMagenta
Write-Host "Make sure this script is not running in Terminal" -ForegroundColor DarkMagenta

Write-Warning "Terminating any open Terminal in 10 seconds so as to proceed with VCLibs installation."
Start-Sleep -Seconds 10

# We're running into a powershell.exe -command as otherwise, if there is no Terminal on, the script just terminates
powershell.exe -command {
    $scriptPath = "https://raw.githubusercontent.com/gabriel-vanca/PowerShell_Library/main/Scripts/Windows/Software/close_windows-terminal.ps1"
    $close_terminal = Invoke-RestMethod $scriptPath
    Invoke-Expression $close_terminal
}

# Installing the component
Add-AppxPackage $fileDownloadLocalPath 
# Removing installation file
Remove-Item $fileDownloadLocalPath 

$vclibsList = Get-AppxPackage Microsoft.VCLibs.140.00.UWPDesktop | Where-Object version -ge $VersionToLookFor
if([string]::IsNullorEmpty($vclibsList)) {
    Write-Host "Microsoft.VCLibs.140.00.UWPDesktop installation failure" -ForegroundColor DarkRed
    Write-Host "This script will terminate in 7 seconds" -ForegroundColor DarkRed
    Start-Sleep -Seconds 7
    throw "Microsoft.VCLibs.140.00.UWPDesktop installation failure"
} else {
    Write-Host "Microsoft.VCLibs.140.00.UWPDesktop sucessfully installed" -ForegroundColor DarkGreen
}

# This is necessary to check when running in Windows Sandbox and on Windows Server
# Winget and Windows Terminal will not run without Microsoft.UI.Xaml
# Both versions 2.7 and 2.8 are necessary
Write-Host "Checking Microsoft UI XAML status" -ForegroundColor DarkYellow

[String]$UIXAML_VersionToLookFor = "7.2208.15002.0"
$UIXAML_List = Get-AppxPackage Microsoft.UI.Xaml.2.7 | Where-Object version -ge $UIXAML_VersionToLookFor
if([string]::IsNullorEmpty($UIXAML_List)) {
    Write-Host "Microsoft.UI.Xaml version missing" -ForegroundColor DarkMagenta
    Write-Host "Initialising install of Microsoft.UI.Xaml" -ForegroundColor DarkYellow

    $WebClient = New-Object System.Net.WebClient
    # https://github.com/microsoft/microsoft-ui-xaml/releases?q=xaml&expanded=true
    $fileURL = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.7.3/Microsoft.UI.Xaml.2.7.x64.appx"
    $fileDownloadLocalPath = "$env:Temp\Microsoft.UI.Xaml.2.7.x64.appx"
    $WebClient.DownloadFile($fileURL, $fileDownloadLocalPath)

    # Installing the component
    Add-AppxPackage $fileDownloadLocalPath 
    # Removing installation file
    Remove-Item $fileDownloadLocalPath 

    $UIXAML_List = Get-AppxPackage Microsoft.UI.Xaml.2.7 | Where-Object version -ge $UIXAML_VersionToLookFor
    if([string]::IsNullorEmpty($UIXAML_List)) {
        Write-Error "Microsoft UI Xaml installation failed."
    } else {
        Write-Host "Microsoft UI Xaml installed successfully" -ForegroundColor DarkGreen
    }
} else {
    Write-Host "Microsoft.UI.Xaml.2.7 present" -ForegroundColor DarkGreen
}

$UIXAML_VersionToLookFor = "8.2306.22001.0"
$UIXAML_List = Get-AppxPackage Microsoft.UI.Xaml.2.8 | Where-Object version -ge $UIXAML_VersionToLookFor
if([string]::IsNullorEmpty($UIXAML_List)) {
    Write-Host "Microsoft.UI.Xaml version missing" -ForegroundColor DarkMagenta
    Write-Host "Initialising install of Microsoft.UI.Xaml" -ForegroundColor DarkYellow

    $WebClient = New-Object System.Net.WebClient
    # https://github.com/microsoft/microsoft-ui-xaml/releases?q=xaml&expanded=true
    $fileURL = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.5/Microsoft.UI.Xaml.2.8.x64.appx"
    $fileDownloadLocalPath = "$env:Temp\Microsoft.UI.Xaml.2.8.x64.appx"
    $WebClient.DownloadFile($fileURL, $fileDownloadLocalPath)

    # Installing the component
    Add-AppxPackage $fileDownloadLocalPath 
    # Removing installation file
    Remove-Item $fileDownloadLocalPath 

    $UIXAML_List = Get-AppxPackage Microsoft.UI.Xaml.2.8 | Where-Object version -ge $UIXAML_VersionToLookFor
    if([string]::IsNullorEmpty($UIXAML_List)) {
        Write-Error "Microsoft UI Xaml installation failed."
    } else {
        Write-Host "Microsoft UI Xaml installed successfully" -ForegroundColor DarkGreen
    }
} else {
    Write-Host "Microsoft.UI.Xaml.2.8 present" -ForegroundColor DarkGreen
}
