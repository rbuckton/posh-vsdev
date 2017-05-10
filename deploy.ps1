param([switch] $Force);

$script:ModuleInfo = Get-Module src\posh-vsdev.psd1 -ListAvailable;
if (-not $script:ModuleInfo) {
    Write-Warning "Could not find src\posh-vsdev.psd1";
    return;
}

$script:Name = $script:ModuleInfo.Name;
$script:Version = $script:ModuleInfo.Version;
if (-not $profile) {
    Write-Warning "Could not find user profile.";
    return;
}

$script:UserProfileDir = Split-Path $profile -Parent;
$script:ModulePaths = $env:PSModulePath -split ";";
$script:CurrentUserModulePath = $script:ModulePaths | Where-Object {
    $_.StartsWith($script:UserProfileDir, [System.StringComparison]::InvariantCultureIgnoreCase);
};
if (-not $script:CurrentUserModulePath) {
    Write-Warning "Could not determine module path";
    return;
}

$script:ModuleDir = "$script:CurrentUserModulePath\$script:Name\$script:Version";
if (-not (Test-Path $script:ModuleDir -PathType:Container)) {
    mkdir $script:ModuleDir -ErrorAction:SilentlyContinue;
}
elseif (-not $Force) {
    Write-Warning "Version $local:Version of $local:Name is already installed. Use -Force to overwrite.";
    return;
}

Copy-Item "src/*.*" -Destination $script:ModuleDir -Container -Recurse -Force;