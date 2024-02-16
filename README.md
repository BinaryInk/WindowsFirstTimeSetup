# Synopsis
Prepares a new Windows 11 installation by removing unnecessary applications, installing desired applications via Winget, applying group policy changes via registry, and activates Windows/Office using Massgrave.

# Registry Changes
- **Prevent Auto-restart while user is logged in**
  - ADD KEY: HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU
    - ADD DWORD: NoAutoRebootWithLoggedOnUsers = 1
- **Disable new Windows 11 context menus**
  - ADD KEY: HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32
- **Disable Web Search from Start Menu**
  - ADD KEY: HKCU:\Software\Policies\Microsoft\Windows\Explorer
    - ADD DWORD: DisableSearchBoxSuggestions = 1
- **Disable telemetry**
  - KEY: HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection
    - ADD DWORD: AllowTelemetry

# Application Installs & Uninstalls

## Apps that are removed from Windows:
- Microsoft.OneDriveSync
- Microsoft.BingNews
- Microsoft.People
- Microsoft.BingWeather
- Microsoft.WindowsMaps
- Microsoft.GetHelp
- Microsoft.MicrosoftStickyNotes
- Microsoft.Getstarted
- Microsoft.MicrosoftOfficeHub
- Microsoft.549981C3F5F10 (Cortana)
- Microsoft.Todos
- Microsoft.ZuneMusic
- Microsoft.ZuneVideo
- Clipchamp.Clipchamp
- microsoft.windowscommunicationsapps
- Microsoft.MicrosoftSolitaireCollection
- Microsoft.YourPhone
- MicrosoftWindows.Client.WebExperience (Widgets)

## Apps that are removed from Windows VMs
- Microsoft.OutlookForWindows
- Microsoft.Paint
- Microsoft.XboxGameOverlay
- MicrosoftCorporationII.QuickAssist
- Microsoft.Xbox.TCUI
- Microsoft.XboxSpeechToTextOverlay
- Microsoft.XboxIdentityProvider
- Microsoft.WindowsCalculator
- Microsoft.WindowsSoundRecorder
- Microsoft.WindowsAlarms
- Microsoft.WindowsCamera
- Microsoft.PowerAutomateDesktop
- Microsoft.ScreenSketch
- Microsoft.XboxGamingOverlay
- Microsoft.GamingApp

## Apps installed via Winget (silently)
- Mozilla.Firefox
- Microsoft.PowerShell

## Apps installed via Winget (interactive)
- Microsoft.VisualStudioCode
- Microsoft.VisualStudio.2022.Community
- Microsoft.OfficeDeploymentTool

## Office Versions available in-script
- Office 2021 VL + Visio Pro 2021 VL + Project Pro 2021 VL

# Disclaimer
Massgrave disclaimer: If you have not purchased a Windows and/or Office license from Microsoft for the machine which you are installing Windows on  (VM or directly), please do not use the '-Activate' switch. This is *only* included for hobby use or a virtual machine on a device which already has a digital license that cannot be leveraged due to a VM's different hardware signature (i.e., a Windows VM on Linux or MacOS to prevent the need for dual-booting/bootcamp).

Please see https://massgrave.dev/ and/or https://github.com/massgravel/Microsoft-Activation-Scripts for more information. If Microsoft ever removes Massgrave from GitHub, it will be removed from this script.
