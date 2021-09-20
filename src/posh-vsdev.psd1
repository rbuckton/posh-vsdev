@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'posh-vsdev.psm1'

    # Version number of this module.
    ModuleVersion = '0.2.1'

    # ID used to uniquely identify this module
    GUID = '796b8bb3-07d8-41b0-8394-231f8359c7a6'

    # Author of this module
    Author = 'Ron Buckton'

    # Copyright statement for this module
    Copyright = '(c) 2021 Ron Buckton'

    # Description of the functionality provided by this module
    Description = 'Sets up the Visual Studio Developmer Command Prompt environment in Powershell.'

    # Minimum version of the Windows PowerShell engine required by this module
    PowerShellVersion = '5.0'

    # Functions to export from this module
    FunctionsToExport = @(
        'Get-VisualStudioInstance',
        'Use-VisualStudioEnvironment',
        'Reset-VisualStudioEnvironment',
        'Reset-VisualStudioInstanceCache',
        'Add-VisualStudioEnvironmentToProfile',
        'Get-WindowsSdk'
    )

    # Cmdlets to export from this module
    CmdletsToExport = @()

    # Variables to export from this module
    VariablesToExport = @(
        'VisualStudioVersion'
    )

    # Aliases to export from this module
    AliasesToExport = @(
        'Get-VisualStudioVersion',
        'Reset-VisualStudioVersionCache',
        'Get-VSInstance',
        'Get-VS',
        'Use-VSEnvironment',
        'Use-VS',
        'Reset-VSEnvironment',
        'Reset-VSInstanceCache'
    )

    # Private data to pass to the module specified in RootModule/ModuleToProcess.
    # This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData = @{
        PSData = @{
            # Tags applied to this module. These help with module discovery in online galleries.
            Tags = @('visualstudio', 'visual', 'studio', 'vs', 'developer')

            # A URL to the license for this module.
            LicenseUri = 'https://github.com/rbuckton/posh-vsdev/blob/master/LICENSE'

            # A URL to the main website for this project.
            ProjectUri = 'https://github.com/rbuckton/posh-vsdev'
        }
    }
}
