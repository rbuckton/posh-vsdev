param ([string]$NugetApiKey);

$Module = @{
    Name = "posh-vsdev"
    RequiredVersion = "0.2.1"
    NugetApiKey = $NugetApiKey
};

Publish-Module @Module;
