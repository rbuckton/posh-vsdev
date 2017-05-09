# Visual Studio Developer Command Prompt for PowerShell
This PowerShell module allows you to configure your PowerShell session as a Visual Studio Developer
Command Prompt by loading the environment variables from the **VsDevCmd.bat** batch file for a
specified Visual Studio version.

# Installation

```powershell
# clone the repository
git clone https://github.com/rbuckton/posh-vs.git;

# Option 1 - Deploy to $env:USERPROFILE\Documents\WindowsPowerShell\Modules
./posh-vs/deploy.ps1;
Import-Module posh-vs;

# Option 2 - Run locally
./posh-vs/install.ps1;
```

# Usage

```powershell
# Get all supported installed Visual Studio instances
Get-VisualStudioVersion

# Get a specific version
Get-VisualStudioVersion -Version 14

# Get a version by release channel
Get-VisualStudioVersion -Channel VisualStudio.15.Release

# Use the Developer Command Prompt Environment
Use-VisualStudioVersion

# Restore previous environment variables
Reset-VisualStudioEnvironment

# Reset cache of Visual Studio instances
Reset-VisualStudioVersionCache

# Add posh-vs to your profile for the current PowerShell host
Add-VisualStudioEnvironmentToProfile -UseEnvironment
```