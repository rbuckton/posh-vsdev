if (Get-Module posh-vsdev) { return; }

. $PSScriptRoot/posh-vsdev.ps1;

# Reset the environment when the module is removed
$ExecutionContext.SessionState.Module.OnRemove = {
    Reset-VisualStudioEnvironment;
};

# Aliases
@{
    # Backwards compatibility
    "Get-VisualStudioVersion" = "Get-VisualStudioInstance";
    "Reset-VisualStudioVersionCache" = "Reset-VisualStudioInstanceCache";

    # Shortcuts
    "Get-VSInstance" = "Get-VisualStudioInstance";
    "Get-VS" = "Get-VisualStudioInstance";
    "Use-VSEnvironment" = "Use-VisualStudioEnvironment";
    "Use-VS" = "Use-VisualStudioEnvironment";
    "Reset-VSEnvironment" = "Reset-VisualStudioEnvironment";
    "Reset-VSInstanceCache" = "Reset-VisualStudioInstanceCache";
}.GetEnumerator() | ForEach-Object { Set-Alias $_.Key $_.Value; };

# Export members
Export-ModuleMember `
    -Function:(
        'Get-VisualStudioInstance',
        'Use-VisualStudioEnvironment',
        'Reset-VisualStudioEnvironment',
        'Reset-VisualStudioInstanceCache',
        'Add-VisualStudioEnvironmentToProfile',
        'Get-WindowsSdk'
    ) `
    -Variable:(
        'VisualStudioVersion'
    ) `
    -Alias:(
        'Get-VisualStudioVersion',
        'Reset-VisualStudioVersionCache',
        'Get-VSInstance',
        'Get-VS',
        'Use-VSEnvironment',
        'Use-VS',
        'Reset-VSEnvironment',
        'Reset-VSInstanceCache'
    );
