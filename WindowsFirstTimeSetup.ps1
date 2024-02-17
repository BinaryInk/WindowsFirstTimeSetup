<#PSScriptInfo
.VERSION 4
.GUID d10b5afa-9179-4b97-8b73-31a3f1ac045b
.AUTHOR brady.greenwood@outlook.com
.COPYRIGHT 2024 Brady Greenwood, GNU GPLv3
#>
<#
.SYNOPSIS
Prepares a new Windows 11 installation by removing unwanted applications,
installing desired applications via Winget, applying group policy changes 
via registry, and can activate Windows/Office using Massgrave.

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
    TODO Remove Microsoft Edge, if possible
    TODO Disable unnecessary services
    TODO Provide proper -Confirm parameter support; would prompt for 
         confirmation for individual application installs, removals, reg hacks,
         etc. if -Confirm is passed or if $ConfirmPreference is Medium/Low. Also
         enables -WhatIf functionality.
    TODO Make Microsoft 365 version of Office default; keep 2021 VL as an option. (enterprise or business?)
    TODO Provide method to prevent installing Visio or Project.
    TODO Error-handle if Winget is not yet updated.
    TODO Add more of the default applications that can be re-installed via Store (e.g., notepad)
        Originally based off of Pro for Workstations; may be beneficial to base off of Pro going forward.
#>

#Requires -RunAsAdministrator

param(
    # Allows skipping various steps in the script.
    # If Install is skipped, Office will be skipped.
    [string[]]
    [ValidateSet('Registry','Install','Office','Uninstall')]
    $Skip,

    # Runs Massgrave script at end
    [switch]
    $Activate,

    # Includes the following packages for removal
    [string[]]
    $IncludeAppxPackagesToRemove,

    # Excludes AppxPackages for removal
    [string[]]
    $ExcludeAppxPackagesToRemove,

    # Includes Winget applications for installation (use exact name)
    [string[]]
    $IncludeWingetApps,

    # Excludes Winget applications for installation (use exact name)
    [string[]]
    $ExcludeWingetApps,

    # Path to custom Office Deployment Tool XML config file
    [string]
    $CustomODTXML

    # Categories of default applications to remove.
    # Has no effect if $UninstallAll is $true.
    # Currently unimplemented
    # [string[]]
    # $UninstallCategories

    # Forces the uninstallation of all default applications. 
    # Has no effect if $Skip contains 'Uninstall'.
    # Currently unimplemented
    # [switch]
    # $UninstallAll,

    # Categories of winget applications to install.
    # Has no effect if $InstallAll is $true
    # Currently unimplemented
    # [string[]]
    # $InstallCategories,

    # Forces the installation of all winget applications.
    # Has no effect if $Skip contains 'Install'.
    # Currently unimplemented
    # [switch]
    # $InstallAll
)

class CustomAppxPackage {
    [string] $DisplayName
    [string] $Name
    [string] $Category

    CustomAppxPackage([string]$Name) {
        $this.Init($Name)
    }
    CustomAppxPackage([string]$Name, [string]$DisplayName) {
        $this.Init($Name, $DisplayName)
    }
    CustomAppxPackage([string]$Name, [string]$DisplayName, [string]$Category) {
        $this.Init($Name, $DisplayName, $Category)
    }

    hidden Init([string]$Name) {
        $this.Name = $Name
        $this.DisplayName = $Name
        $this.Category = "N/A"
    }
    hidden Init([string]$Name, [string]$DisplayName) {
        $this.Init($Name)
        $this.DisplayName = $DisplayName
    }
    hidden Init([string]$Name, [string]$DisplayName, [string]$Category) {
        $this.Init($Name, $DisplayName)
        $this.Category = $Category
    }

    [void] UninstallUser() {
        Write-Verbose "Uninstalling $($this.DisplayName) for $env:USERNAME..."
        Get-AppxPackage -Name $this.Name | Remove-AppxPackage
    }
    [void] UninstallAllUsers() {
        Write-Verbose "Uninstalling $($this.DisplayName) for all users..."
        Get-AppxPackage -Name $this.Name -AllUsers| Remove-AppxPackage
    }
    [void] RemoveProvisioning() {
        Write-Verbose "`tRemoving $($this.DisplayName) for future users..."
        Get-AppxProvisionedPackage -Online | 
            Where-Object { $_.PackageName -like "*$($Name)*"} |
            Remove-AppxProvisionedPackage -Online -AllUsers
    }
    [void] UninstallAll() {
        $this.UninstallUser()
        $this.UninstallAllUsers()
        # $this.RemoveProvisioning() # TODO is this necessar, if so how to fix "Removal Failed."?
    }
    
}

class CustomWingetApp {
    [string]   $DisplayName
    [string]   $Name
    [string]   $Category
    [bool]     $Interactive = $false
    [string[]] $OverrideArguments = @()
    [bool]     $Excluded = $false

    CustomWingetApp([string]$Name) {
        $this.Init($Name)
    }
    CustomWingetApp([string]$Name, [string]$DisplayName) {
        $this.Init($Name, $DisplayName)
    }
    CustomWingetApp([string]$Name, [string]$DisplayName, [string]$Category) {
        $this.Init($Name, $DisplayName, $Category)
    }

    hidden Init([string]$Name) {
        $this.Name = $Name
        $this.DisplayName = $Name
        $this.Category = 'N/A'
    }
    hidden Init([string]$Name, [string]$DisplayName) {
        $this.Init($Name)
        $this.DisplayName = $DisplayName
    }
    hidden Init([string]$Name, [string]$DisplayName, [string]$Category) {
        $this.Init($Name, $DisplayName)
        $this.Category = $Category
    }

    [void] Install() {
        [string]$Command = "winget install $($this.Name) --exact --accept-source-agreements --accept-package-agreements"
        
        if ($this.Interactive) { $Command += " --interactive" }
        if ($this.OverrideArguments.Count -gt 0) { $Command += " --override $($this.OverrideArguments)"}

        Write-Debug $Command
        Invoke-Expression $Command #FIXME Why doesn't winget show progress?
    }

    [void] Uninstall() {
        # TODO Implement CustomWingetApp.Uninstall()
    }
}

function Request-YesNoAnswer([string]$Prompt){
    while ($Answer -ne 'Y' -or $Answer -ne 'N') {
        $Answer = Read-Host -Prompt "$Prompt (Y/N)"
        if ($Answer -eq 'Y' -or $Answer -eq 'N') {
            return $Answer -eq 'Y'
        }
        Write-Host 'You must answer with "Y" or "N"'
    }
}

function Uninstall-CustomAppxPackages([CustomAppxPackage[]]$Packages) {
    foreach ($app in $Packages) {
        if ($ExcludeAppxPackagesToRemove -notcontains $app) {
            Write-Host "Removing Appx Package $($app.DisplayName)..."
            $app.UninstallAll()
        }
    }
}

function Start-AppxUninstallProcess() {
    $AppxPackagesToRemove = @(
        [CustomAppxPackage]::new('Microsoft.OutlookForWindows', 'Outlook for Windows', 'Communications' )
        [CustomAppxPackage]::new('Microsoft.Paint', 'Microsoft Paint', 'Graphics')
        [CustomAppxPackage]::new('Microsoft.XboxGameOverlay', 'Xbox Game Overlay', 'Gaming')
        [CustomAppxPackage]::new('MicrosoftCorporationII.QuickAssist', 'Quick Assist', 'RemoteAssistance')
        [CustomAppxPackage]::new('Microsoft.Xbox.TCUI', 'Xbox TCUI', 'Gaming')
        [CustomAppxPackage]::new('Microsoft.XboxSpeechToTextOverlay', 'Xbox Speech to Text Overlay', 'Gaming')
        [CustomAppxPackage]::new('Microsoft.XboxIdentityProvider', 'Xbox Identity Provider', 'Gaming')
        [CustomAppxPackage]::new('Microsoft.WindowsCalculator', 'Calculator', 'Productivity')
        [CustomAppxPackage]::new('Microsoft.WindowsSoundRecorder', 'Sound Recorder', 'Audio')
        [CustomAppxPackage]::new('Microsoft.WindowsAlarms', 'Alarms', 'Productivity')
        [CustomAppxPackage]::new('Microsoft.WindowsCamera', 'Camera', 'Video')
        [CustomAppxPackage]::new('Microsoft.PowerAutomateDesktop', 'Power Automate Desktop', 'Productivity')
        [CustomAppxPackage]::new('Microsoft.ScreenSketch', 'Screen Sketch', 'Pen')
        [CustomAppxPackage]::new('Microsoft.XboxGamingOverlay', 'Xbox Gaming Overlay', 'Gaming')
        [CustomAppxPackage]::new('Microsoft.GamingApp', 'Xbox App', 'Gaming')
        [CustomAppxPackage]::new('Microsoft.OneDriveSync', 'OneDrive', 'Cloud')
        [CustomAppxPackage]::new('Microsoft.BingNews', 'Bing News', 'News')
        [CustomAppxPackage]::new('Microsoft.People', 'People', 'Communications')
        [CustomAppxPackage]::new('Microsoft.BingWeather', 'Weather', 'News')
        [CustomAppxPackage]::new('Microsoft.WindowsMaps', 'Maps', 'News')
        [CustomAppxPackage]::new('Microsoft.GetHelp', 'Get Help', 'Assist')
        [CustomAppxPackage]::new('Microsoft.MicrosoftStickyNotes', 'Sticky Notes', 'Productivity')
        [CustomAppxPackage]::new('Microsoft.Getstarted', 'Get Started', 'Assist')
        [CustomAppxPackage]::new('Microsoft.MicrosoftOfficeHub', 'Office Hub', 'Junk')
        [CustomAppxPackage]::new('Microsoft.549981C3F5F10', 'Cortana', 'Junk')
        [CustomAppxPackage]::new('Microsoft.Todos', 'Microsoft To Do', 'Productivity')
        [CustomAppxPackage]::new('Microsoft.ZuneMusic', 'Windows Media Player', 'Audio')
        [CustomAppxPackage]::new('Microsoft.ZuneVideo', 'Movies & TV', 'Video')
        [CustomAppxPackage]::new('Clipchamp.Clipchamp', 'Clipchamp', 'Video')
        [CustomAppxPackage]::new('microsoft.windowscommunicationsapps', 'Windows Communications Apps', 'Communications')
        [CustomAppxPackage]::new('Microsoft.MicrosoftSolitaireCollection', 'Microsoft Solitaire Collection', 'Gaming')
        [CustomAppxPackage]::new('Microsoft.YourPhone', 'Your Phone',  'Communications')
        [CustomAppxPackage]::new('MicrosoftWindows.Client.WebExperience', 'Windows 11 Widgets', 'News')
    ) | Sort-Object -Property 'Category'

    # Prevent Write-Progress from consuming the screen (chances are, this is being executed on Windows PowerShell using the legacy Write-Progress)
    $Global:ProgressPreference = 'SilentlyContinue'
    Uninstall-CustomAppxPackages -Packages $AppxPackagesToRemove
    # Include
    if ($IncludeAppxPackagesToRemove) {
        [CustomAppxPackage[]]$pkgs
        $IncludeAppxPackagesToRemove | ForEach-Object {
            $pkgs += [CustomAppxPackage]::new($_)
        }
        Uninstall-CustomAppxPackages -Packages $pkgs
    }
    $Global:ProgressPreference = 'Continue'
}

function Start-WingetInstalls() {
    Write-Warning 'Winget must manually be updated through the Windows Store before installation will work. If you have not updated "App Installer" yet, ctrl-c to cancel this script, update all apps via the Windows Store (specifically: App Installer). Once updated, continue.'

    Pause

    $WingetApps = @(
        [CustomWingetApp]::new("Mozilla.Firefox", "Firefox", "Web Browser")
        [CustomWingetApp]::new("Microsoft.PowerShell", "PowerShell", "IT")
        [CustomWingetApp]::new("Microsoft.VisualStudioCode", "Visual Studio Code", "Development")
        [CustomWingetApp]::new("Microsoft.VisualStudio.2022.Community", "Visual Studio Community 2022", "Development")
    )
    # FIXME Eventually change all of this. ###
    $WingetApps[2].Interactive = $true
    $WingetApps[3].Interactive = $true

    Write-Host 'Installing apps via WinGet...'
    foreach ($app in $WingetApps) { 
        if ($ExcludeWingetApps -notcontains $app) {
            Write-Host "Installing Winget App: $($app.DisplayName)"
            $app.Install()
        }
    }
    foreach ($appName in $IncludeWingetApps) {
        if ($ExcludeWingetApps -notcontains $app) {
            $app = [CustomWingetApp]::new($appName)
            Write-Host "Installing Winget App: $($app.DisplayName)"
            $app.Install()
        }
    }
    ###
}

function Install-MicrosoftOffice() {
    $OdtConfig_VL2021 = @"
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
    $OdtInstallPath = "$env:TEMP\ODT"

    if ($CustomODTXML) {
        $ConfigXmlPath = $CustomODTXML    
    }
    else {
        $ConfigXmlPath = "$HOME\ODT.xml"
        $OdtConfig_VL2021 | Out-File $ConfigXmlPath
    }

    Write-Host 'Installing Office via Office Deployment Tool...'
    Write-Warning 'Manual Intervention Required: Accept EULA'
    $OdtWingetApp = [CustomWingetApp]::new("Microsoft.OfficeDeploymentTool", "Office Deployment Tool", "Office")
    $OdtWingetApp.OverrideArguments = "/extract:$OdtInstallPath"
    $OdtWingetApp.Install()
    
    Invoke-Expression "$OdtInstallPath\setup.exe /configure $ConfigXmlPath"

    Remove-Item -Path $OdtInstallPath -Recurse -Force
}

function Install-RegistryChanges() {
    Write-Host "Disabling Windows Update Auto-Restart (while user is logged in)..."
    New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' -Force | Out-Null
    New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' `
                    -Name 'NoAutoRebootWithLoggedOnUsers' `
                    -Value 1 `
                    -PropertyType 'DWORD' `
                    -Force |
        Out-Null

    Write-Host "Disabling Windows 11 Context Menu..."
    New-Item -Path 'HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32' `
                -Force |
        Out-Null

    Write-Host "Disabling Web Search from Start..."
    New-Item -Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer" -Force | Out-Null
    New-ItemProperty -Path 'HKCU:\Software\Policies\Microsoft\Windows\Explorer' `
                    -Name 'DisableSearchBoxSuggestions' `
                    -Value 1 `
                    -PropertyType 'DWORD' `
                    -Force |
        Out-Null

    Write-Host "Disabling Windows Telemetry..."
    New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' `
                    -Name 'AllowTelemetry' `
                    -Value 0 `
                    -PropertyType 'DWORD' `
                    -Force |
        Out-Null
    Write-Warning "Restart required for disabling of telemetry to take effect!"
    
    Write-Warning "Restarting File Explorer to apply changes to context menu immediately"
    Get-Process 'explorer' | Stop-Process
}

## SCRIPT ##
if ($Skip -notcontains 'Registry') { 
    Install-RegistryChanges 
}

if ($Skip -notcontains 'Uninstall') { 
    Start-AppxUninstallProcess
}

if ($Skip -notcontains 'Install') {
    Start-WingetInstalls
}

if ($Skip -notcontains 'Office') {
    Install-MicrosoftOffice
}

if ($Activate) { 
    Write-Warning "If you have not purchased a Windows and/or Office license from Microsoft for the machine which you are installing Windows on (VM or directly), please do not use the '-Activate' switch. This is *only* included for hobby use or a virtual machine on a device which already has a digital license that cannot be leveraged due to a VM's different hardware signature (i.e., a Windows VM on Linux or MacOS to prevent the need for dual-booting/bootcamp)`n`n
    This script is not a part of the WindowsFirstTimeSetup.ps1 script. Please see https://massgrave.dev/ and/or https://github.com/massgravel/Microsoft-Activation-Scripts for more information.`n`n
    Press Ctrl-C to cancel."
    Pause
    Invoke-RestMethod https://massgrave.dev/get | Invoke-Expression 
}
