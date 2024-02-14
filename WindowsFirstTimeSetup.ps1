<#PSScriptInfo
.VERSION 3
.GUID d10b5afa-9179-4b97-8b73-31a3f1ac045b
.AUTHOR brady.greenwood@outlook.com
.COPYRIGHT 2024 Brady Greenwood, GNU GPLv3
#>
<#
.SYNOPSIS
Prepares a new Windows 11 installation by removing unnecessary applications,
installing desired applications via Winget, applying group policy changes 
via registry, and activates Windows/Office using Massgrave.

.EXAMPLE
./WindowsFirstTimeSetup.ps1

.EXAMPLE
./WindowsFirstTimeSetup.ps1 -Force

.EXAMPLE
./WindowsFirstTimeSetup.ps1 -Activate

.EXAMPLE
./WindowsFirstTimeSetup.ps1 -Skip 'Install', 'Office', 'Registry'
#>

<# Roadmap/To-do
    TODO Move application lists to string[] parameters for customization.
    TODO Provide parameters to include/exclude extra apps to remove 
        (-IncludeAppsToRemove, -ExcludeAppsToRemove)
    TODO Provide parameters to include/exclude extra apps to install 
        (-IncludeAppsToInstall, -ExcludeAppsToInstall)
    TODO Provide parameter to point to a custom XML file for ODT 
        (-ODTConfigPath)
    TODO Remove Microsoft Edge, if possible
    TODO Disable unnecessary services
    TODO Disable telemetry
    TODO Provide proper -Confirm parameter support; would prompt for 
         confirmation for individual application installs, removals, reg hacks,
         etc. if -Confirm is passed or if $ConfirmPreference is Medium/Low
    TODO Separate Office Deployment Tool from Winget installs.
    TODO Automate ODT installation instead of prompting user where to install. Leverage $TEMP.
    TODO Make Microsoft 365 version of Office default; keep 2021 VL as an option.
    TODO Provide method to prevent installing Visio or Project.
    TODO Error-handle if Winget is not yet updated.
#>

#Requires -RunAsAdministrator

param(
    # Allows skipping various steps in the script.
    # If Install is skipped, Office will be skipped.
    [string[]]
    [ValidateSetAttribute('Registry','Install','Office','Uninstall')]
    $Skip,

    # Runs Massgrave script at end
    [switch]
    $Activate,

    # Forces the uninstallation of all default applications. 
    # Has no effect if $Skip contains 'Uninstall'.
    [switch]
    $Force
)

if ($IsLinux -or $IsMacOS) {
    throw "This script is only intended to run on Windows 11."
}

$COMPUTER_MODEL = $(Get-WmiObject win32_computersystem).model
$IS_VIRTUAL = $COMPUTER_MODEL -eq "VirtualBox" -or 
              $COMPUTER_MODEL -eq "VMware Virtual Platform" -or 
              $COMPUTER_MODEL -eq "Virtual Machine" # TODO: Hyper-V
$APPS_TO_REMOVE_VM = @(
	"Microsoft.OutlookForWindows"
    "Microsoft.Paint"
    "Microsoft.XboxGameOverlay"
    "MicrosoftCorporationII.QuickAssist"
    "Microsoft.Xbox.TCUI"
    "Microsoft.XboxSpeechToTextOverlay"
    "Microsoft.XboxIdentityProvider"
    "Microsoft.WindowsCalculator"
    "Microsoft.WindowsSoundRecorder"
    "Microsoft.WindowsAlarms"
    "Microsoft.WindowsCamera"
    "Microsoft.PowerAutomateDesktop"
    "Microsoft.ScreenSketch"
    "Microsoft.XboxGamingOverlay"
    "Microsoft.GamingApp"
    "Microsoft.XboxGameCallableUI"
)
$APPS_TO_REMOVE = @(
	"Microsoft.OneDriveSync"
	"Microsoft.BingNews"
    "Microsoft.People"
    "Microsoft.BingWeather"
    "Microsoft.WindowsMaps"
    "Microsoft.GetHelp"
    "Microsoft.MicrosoftStickyNotes"
    "Microsoft.Getstarted"
    "Microsoft.MicrosoftOfficeHub"
    "Microsoft.549981C3F5F10" # Cortana
    "Microsoft.Todos"
    "Microsoft.ZuneMusic"
    "Microsoft.ZuneVideo"
    "Clipchamp.Clipchamp"
    "microsoft.windowscommunicationsapps"
    "Microsoft.MicrosoftSolitaireCollection"
    "Microsoft.YourPhone"
    "MicrosoftWindows.Client.WebExperience" # Widgets
)
# Use 'winget search <appname>' to find names for apps
## WinGet must be updated prior to first use. Update via Windows Store & Windows Update
$WINGET_TO_INSTALL = @(
	"Mozilla.Firefox",
    "Microsoft.PowerShell"    
)
$WINGET_INTERACTIVE_INSTALL = @(
	"Microsoft.VisualStudioCode"
    "Microsoft.VisualStudio.2022.Community"
    "Microsoft.OfficeDeploymentTool"
)
# Use Office Customization Tool (https://config.office.com/deploymentsettings) to generate own XML
$ODT_XML_FILE = @"
<Configuration ID="f1b7d029-c64e-4d7d-8fb5-59512a7b9e39">
  <Add OfficeClientEdition="64" Channel="PerpetualVL2021" MigrateArch="TRUE">
    <Product ID="Standard2021Volume" PIDKEY="KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3">
      <Language ID="en-us" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Outlook" />
    </Product>
    <Product ID="VisioPro2021Volume" PIDKEY="KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4">
      <Language ID="en-us" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Outlook" />
    </Product>
    <Product ID="ProjectPro2021Volume" PIDKEY="FTNWT-C6WBT-8HMGF-K9PRX-QV9H8">
      <Language ID="en-us" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Outlook" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
  </AppSettings>
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@
$ODT_FILE_REMOVES = @(
	"$env:USERPROFILE\configuration-Office365-x64.xml"
    "$env:USERPROFILE\configuration-Office365-x86.xml"
    "$env:USERPROFILE\configuration-Office2019Enterprise.xml"
    "$env:USERPROFILE\configuration-Office2021Enterprise.xml"
    "$env:USERPROFILE\ODT.xml"
    "$env:USERPROFILE\setup.exe"
)

# Registry hacks
if ($Skip -notcontains 'Registry') {
    Write-Host 'Applying group policy and registry settings...'
    ## Disable Windows Update auto-restart
    New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
    New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' `
                     -Name 'NoAutoRebootWithLoggedOnUsers' `
                     -Value 1 `
                     -PropertyType 'DWORD'
    ## Disable Win11 Context Menu
    New-Item -Path 'HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}'
    New-Item -Path 'HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32'
    ## Disable Web Search from Start
    New-Item -Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer"
    New-ItemProperty -Path 'HKCU:\Software\Policies\Microsoft\Windows\Explorer' `
                     -Name 'DisableSearchBoxSuggestions' `
                     -Value 1 `
                     -PropertyType 'DWORD'
    ## Restart File Explorer to apply changes to context menu immediately
    Get-Process 'explorer' | Stop-Process
}

# Windows Store Default Apps removal
if ($Skip -notcontains 'Uninstall') {
    Write-Host 'Removing pre-installed applications...'
    foreach ($app in $APPS_TO_REMOVE) {
        Write-Verbose "`tRemoving $app for current user..."
        Get-AppxPackage -Name $app | Remove-AppxPackage
        Write-Verbose "`tRemoving $app for all users..."
        Get-AppxPackage -Name $app -AllUsers | Remove-AppxPackage
        Write-Verbose "`tRemoving $app for future users..."
        Get-AppxProvisionedPackage -Online | 
            Where-Object { $_.PackageName -like "*$($app)*"} |
            Remove-AppxProvisionedPackage -Online
    }
    if ($IS_VIRTUAL -or $ForceAllDefaultAppUninstall) {
        Write-Host 'Removing pre-installed applications for VM...'
        foreach ($app in $APPS_TO_REMOVE_VM) {
            Write-Verbose "`tRemoving $app for current user..."
            Get-AppxPackage -Name $app | Remove-AppxPackage
            Write-Verbose "`tRemoving $app for all users..."
            Get-AppxPackage -Name $app -AllUsers | Remove-AppxPackage
            Write-Verbose "`tRemoving $app for future users..."
            Get-AppxProvisionedPackage -Online | 
                Where-Object { $_.PackageName -like "*$($app)*"} |
                Remove-AppxProvisionedPackage -Online
        }
    }
}

# Winget Installations
if ($Skip -notcontains 'Install') {
    Write-Host 'Installing apps via WinGet...'
    foreach ($app in $WINGET_TO_INSTALL) {
        Write-Verbose "Installing $app..."
        winget install $app
    }
    foreach ($app in $WINGET_INTERACTIVE_INSTALL) {
        Write-Verbose "Installing $app..."
        if ($app -eq 'Microsoft.OfficeDeploymentTool') {
            if ($SkipOfficeDeploymentTool) { continue }
            Write-Warning "`tOffice Deployment Tool: Install to C:\Users\<UserName>"
        }
        winget install $app --interactive
    }

    if ($Skip -notcontains 'Office') {
        # MSOffice Install (2021 VL)
        Write-Host 'Installing Office via Office Deployment Tool...'
        $ODT_XML_FILE | Out-File "$env:USERPROFILE\ODT.xml"
        Invoke-Expression "$env:USERPROFILE\setup.exe /configure $env:USERPROFILE\ODT.xml"
        foreach ($file in $ODT_FILE_REMOVES) {
            Remove-Item $file
        }
    }
}

# Activate Windows and/or Office
if ($Activate) { 
    Write-Warning "If you have not purchased a Windows and/or Office license from Microsoft for the machine which you are installing Windows on (VM or directly), please do not use the '-Activate' switch. This is *only* included for hobby use or a virtual machine on a device which already has a digital license that cannot be leveraged due to a VM's different hardware signature (i.e., a Windows VM on Linux or MacOS to prevent the need for dual-booting/bootcamp)`n`n
    This script is not a part of the WindowsFirstTimeSetup.ps1 script. Please see https://massgrave.dev/ and/or https://github.com/massgravel/Microsoft-Activation-Scripts for more information.`n`n
    Press Ctrl-C to cancel."
    Pause
    Invoke-RestMethod https://massgrave.dev/get | Invoke-Expression 
}
