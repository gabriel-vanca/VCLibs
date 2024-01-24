<#
.SYNOPSIS
    Installs and updates Microsoft VCLibs
.DESCRIPTION
	Installs Microsoft VCLibs if no VCLibs.14 is detected at all,
    or updates it if the subversion detected is too old.
    ⚠️This is normally only necessary on Windows Server and Windows Sandbox,
    but ocasionally it is required on consumer deployments too.
    Read the GitHub repo description for a more in-depth investigation.
    ⚠️ Winget and Windows Terminal will not run without Microsoft VCLibs
    
    🪟Deployment tested on:
        - ✅Windows 10
        - ✅Windows 11
        - ✅Windows Sandbox
        - ✅Windows Server 2019
        - ✅Windows Server 2022
        - ✅Windows Server 2022 vNext (Windows Server 2025)
.PARAMETER BypassChocolatey
    (Optional)
    By default disabled. If Chocolatey is installed on the system,
    Chocolatey will be used to install/update VCLibs.
    If Chocolatey is not present, this parameter has no effect.
    If Chocolatey is installed but Chocolatey install fails for
    whatever reason, the manual install will occur as if the bypass
    was activated.
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
    PS> ./Deploy_MS_VCLibs
.LINK
	https://github.com/gabrielvanca/VCLibs
.NOTES
	Author: Gabriel Vanca
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $False)] [Switch]$BypassChocolatey = $False,
    [Parameter(Mandatory = $False)] [Switch]$ForceReinstall = $False
)

[String]$VersionToLookFor = "14.0.30704.0"
[Switch]$ChocolateyInstalled = $False
[Switch]$MustUseChocolatey = $False
[Switch]$MustUninstall = $False


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

# Expected path of the choco.exe file.
$chocoInstallPath = "$Env:ProgramData/chocolatey/choco.exe"
if (Test-Path "$chocoInstallPath") {
    $ChocolateyInstalled = $True
    if($BypassChocolatey) {
        Write-Error "Chocolatey is present, but bypass is enabled" -ForegroundColor DarkYellow
        Write-Error "Proceeding with manual install." -ForegroundColor DarkYellow
        $MustUseChocolatey = $False
        Start-Sleep -Seconds 5
    } else {
        Write-Error "Chocolatey is present. Proceeding with Chocolatey install." -ForegroundColor DarkGreen
        $MustUseChocolatey = $True
    }
} else {
    $ChocolateyInstalled = $False
    $MustUseChocolatey = $False
    if(!($BypassChocolatey)) {
        Write-Host "Chocolatey is not present." -ForegroundColor DarkMagenta
        Start-Sleep -Seconds 5
    }
    Write-Error "Proceeding with manual install." -ForegroundColor DarkYellow
}


# TODO BELOW



if($ForceReinstall) {
    Write-Host "FORCED (RE)INSTALL ENABLED." -ForegroundColor DarkYellow
    Write-Host "Skipping version check." -ForegroundColor DarkYellow
    $MustUninstall = $True
} else {
    $vclibsList = Get-AppxPackage Microsoft.VCLibs.140.00.UWPDesktop | Where-Object version -ge $VersionToLookFor
    if([string]::IsNullorEmpty($vclibsList)) {
        Write-Host "Microsoft.VCLibs.140.00.UWPDesktop missing or too old" -ForegroundColor DarkRed
        Write-Warning "VCLibs package required for WinGet and Terminal installation."
        $MustUninstall = $True
    } else {
        Write-Host "The installed version of Microsoft.VCLibs.140.00.UWPDesktop is the same or newer as the version we are looking for." -ForegroundColor DarkGreen
        $MustUninstall = $False
        Return
    }
}

if($MustUninstall) {
    Write-Host "Initialising uninstall." -ForegroundColor DarkYellow
    Write-Warning "If anything relies on VCLibs, even if not turned on, the uninstall will fail."
    Start-Sleep -Seconds 7
    if($ChocolateyInstalled) {
        try {
            choco uninstall microsoft-vclibs
        } catch {
            # Expected to fail if vclibs not installed via Chocolatey
            # No need to handle the error
        }
    }
    Get-AppxPackage "Microsoft.VCLibs.140.00.UWPDesktop" | remove-AppxPackage -allusers
}


Write-Host "Installing VCLibs package" -ForegroundColor DarkYellow

if($MustUseChocolatey) {
    try {
        choco install microsoft-vclibs -y --ignore-checksums
        Write-Host "Microsoft.VCLibs.140.00.UWPDesktop sucessfully installed" -ForegroundColor DarkGreen
        Return
    } catch {
        $MustUseChocolatey = $False
        # No need to handle the error in catch{}. Error handled bellow automatically through manual install
    }
}

# Downloading necessary graphical component, usually for Windows Server or Sandbox deployments
# (https://docs.microsoft.com/en-us/troubleshoot/developer/visualstudio/cpp/libraries/c-runtime-packages-desktop-bridge)
$WebClient = New-Object System.Net.WebClient
$fileURL = "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx"
$fileDownloadLocalPath = "$env:Temp\Microsoft.VCLibs.x64.14.00.Desktop.appx"
$WebClient.DownloadFile($fileURL, $fileDownloadLocalPath)

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

Write-Host "Proceeding with validation" -ForegroundColor DarkYellow
$vclibsList = Get-AppxPackage Microsoft.VCLibs.140.00.UWPDesktop | Where-Object version -ge $VersionToLookFor
if([string]::IsNullorEmpty($vclibsList)) {
    Write-Error "Microsoft.VCLibs.140.00.UWPDesktop installation failure"
    Start-Sleep -Seconds 3
    throw "Microsoft.VCLibs.140.00.UWPDesktop installation failure"
} else {
    Write-Host "Microsoft.VCLibs.140.00.UWPDesktop sucessfully installed" -ForegroundColor DarkGreen
}
