# Visual Studio Developer Command Prompt for PowerShell
[![psget](https://img.shields.io/github/release/rbuckton/posh-vsdev.svg?label=psget&colorB=0072c6)](https://www.powershellgallery.com/packages/posh-vsdev/0.1.0)

This PowerShell module allows you to configure your PowerShell session as a Visual Studio Developer
Command Prompt by loading the environment variables from the **VsDevCmd.bat** batch file for a
specified Visual Studio version.

# Installation

> NOTE: `posh-vsdev` requires PowerShell 5.0 or higher.

### Inspect
```powershell
PS> Save-Module -Name posh-vsdev
```

### Install
```powershell
PS> Install-Module -Name posh-vsdev -Scope CurrentUser
```

# Usage
```powershell
# Import the posh-vsdev module module
Import-Module posh-vsdev

# Get all installed Visual Studio instances
Get-VisualStudioVersion

# Get a specific version
Get-VisualStudioVersion -Version 14

# Use the Developer Command Prompt Environment
Use-VisualStudioVersion

# Restore the non-Developer environment
Reset-VisualStudioEnvironment

# Reset cache of Visual Studio instances
Reset-VisualStudioVersionCache

# Add posh-vsdev to your PowerShell profile
Add-VisualStudioEnvironmentToProfile -UseEnvironment
```