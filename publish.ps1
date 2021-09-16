param ([string]$NugetApiKey);

$Module = @{
    Path = (Resolve-Path "$PSScriptRoot/src").Path
    NugetApiKey = $NugetApiKey
};

Publish-Module @Module;
