if (Get-Module posh-vs) { return; }

. $PSScriptRoot\VisualStudio-scripts.ps1;

Export-ModuleMember -Function:(
    'Get-VisualStudioVersion',
    'Use-VisualStudioEnvironment',
    'Reset-VisualStudioEnvironment',
    'Reset-VisualStudioVersionCache',
    'Add-VisualStudioEnvironmentToProfile'
);