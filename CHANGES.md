# Version 4.1 (2024-02-26)
- Fixed -IncludeWingetApps
# Version 4 (2024-02-15)
  ## New
    - Registry: Added registry change to disable telemetry
    - Appx: Added include/exclude parameters
    - Winget: Warning to ensure winget is updated through the store before proceeding with winget installs
    - Winget: Auto-accepting of source/package agreements
    - Winget: Added include/exclude parameters
  ## Fixes
    - Fixed various untested, breaking changes in V3
    - Hiding unneeded output from registry changes (and various others)
  ## Changes
    - Future: Changed -Force to -UninstallAll for eventual ShouldProcess implementation.
    - Encapsulation:
      - CustomAppxPackage class to provide future functionality
      - CustomWingetApp class to provide future functionality
    - All applications are now removed, no arbitrary definition between VM and non-VM apps.
      - Future: Will be replaced by categorical choice(s)
    - Future: Separated ODT install from rest of winget installs

# Version 3 (2024-02-14)
  - Collapsed all -Skip switch parameters into single -Skip string[] ValidateSet parameter.
    - Fixed skip flag for Office VL as a result
  - Activation is now opt-in rather than opt-out.
  - Renamed -Force parameter to match documentation.
  - Syntax unification on string array declarations.

# Version 2 (2024-02-11):
  - Uninstalls widgets instead of disabling via GP registry key.
  - Disables web search from start menu.
  - Provided switch parameters to skip registry hacks, activation, application removal, and application installations.
  - Provided switch parameter to force uninstall of "VM-only" uninstalls on a potential non-vm system.
    - Since the distinction of "VM" apps is arbitrary, makes sense to provide a way to uninstall virtually everything.
  - Removed use of cmdlet aliases (conform with best practices).
  - Removal of unnecessary Set-ExecutionPolicy, as the script is no longer in an easily copy-and-paste-able format.
  - Massgrave activation disclaimer added.
  - GNU GPL license added.
  - Proper PSScript Get-Help documentation & metadata.
  - Prevent running on non-Windows OSes.
  - Misc. syntax/readability improvements.

# Version 1 (2024-02-10):
  - Uninstalls majority of default applications
    - App lists separated for VM use
  - Installs desired applications via Winget
    - Extra scripting to install Office 2021 VL via Office Deployment Tool
  - Applies registry edits
    - Remove Win11 Widgets feature
    - Prevent auto-restart for Windows Update while user is logged in
    - Disable Win11 context menu in favor of <= Win10 context menu
  - Invokes Massgrave script to activate Windows and/or Office 2021 VL (internet connection req.)
