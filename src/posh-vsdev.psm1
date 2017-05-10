if (Get-Module posh-vsdev) { return; }

# Simplifies access to HashSet<string>
class Set : System.Collections.Generic.HashSet[string] {
    Set() { }
    Set([string[]] $Data) {
        foreach($local:Item in $Data) {
            $this.Add($local:Item);
        }
    }
}

# Encapsulates environment variables and their values
class Environment : System.Collections.Generic.Dictionary[string,string] {
    hidden static [Environment] $_Default;

    Environment() {}

    static [Environment] GetDefault() {
        if ([Environment]::_Default -eq $null) {
            [Environment]::_Default = [Environment]::GetCurrent();
        }
        return [Environment]::_Default;
    }

    static [Environment] GetCurrent() {
        $local:Env = [Environment]::new();
        foreach($local:Item in Get-ChildItem "ENV:\") {
            $local:Env[$local:Item.Name] = $local:Item.Value;
        }
        return $local:Env;
    }

    hidden [string] get_Item([string] $Key) {
        $Value = $null;
        [void]($this.TryGetValue($Key, [ref]$Value));
        return $Value;
    }

    [void] Apply() {
        [void]([Environment]::GetDefault());
        $local:Current = [Environment]::GetCurrent();
        foreach ($local:Item in $local:Current.GetEnumerator()) {
            if (-not $this.ContainsKey($local:Item.Key)) {
                script:SetEnvironmentVariable $local:Item.Key $null;
            }
        }
        foreach ($local:Item in $this.GetEnumerator()) {
            script:SetEnvironmentVariable $local:Item.Key $local:Item.Value;
        }
    }

    [Environment] Clone() {
        [Environment] $local:Env = [Environment]::new();
        foreach ($local:Entry in $this.GetEnumerator()) {
            $local:Env[$local:Entry.Key] = $local:Entry.Value;
        }
        return $local:Env;
    }
}

# Stores a diff between two paths
class PathDiff {
    hidden [string[]] $Added;
    hidden [string[]] $Removed;
    hidden [Set] $RemovedSet;

    hidden PathDiff([string[]] $Added, [string[]] $Removed) {
        $this.Added = @() + $Added;
        $this.Removed = @() + $Removed;
        $this.RemovedSet = [Set]::new($Removed);
    }

    static [PathDiff] FromObject([psobject] $Object) {
        if ($Object -eq $null) { return $null; }
        if ($Object -is [PathDiff]) { return $Object; }
        return [PathDiff]::new($Object.Added, $Object.Removed);
    }

    static [psobject] ToObject([PathDiff] $Object) {
        if ($Object -eq $null) { return $null; }
        return @{
            Added = @() + $Object.Added;
            Removed = @() + $Object.Removed;
        };
    }

    static [PathDiff] DiffBetween([string[]] $OldPaths, [string[]] $NewPaths) {
        [Set] $local:OldSet = [Set]::new($OldPaths);
        [Set] $local:NewSet = [Set]::new($NewPaths);
        [string[]] $local:Added = @();
        [string[]] $local:Removed = @();
        foreach ($local:Path in $NewSet.GetEnumerator()) {
            if (-not $OldSet.Contains($local:Path)) {
                $local:Added += $local:Path;
            }
        }
        foreach ($local:Path in $OldSet.GetEnumerator()) {
            if (-not $NewSet.Contains($local:Path)) {
                $local:Removed += $local:Path;
            }
        }
        return [PathDiff]::new($local:Added, $local:Removed);
    }

    [string] Apply([string] $Path) {
        return $this.ApplyToPaths($Path -split ";") -join ";";
    }

    [string[]] Apply([string[]] $Paths) {
        return $this.ApplyToPaths($Paths);
    }

    hidden [string[]] ApplyToPaths([string[]] $Paths) {
        $local:Result = @();
        foreach ($local:Path in $Paths) {
            if ($local:Path -and $local:Path.Trim() -and -not $this.RemovedSet.Contains($local:Path)) {
                $local:Result += $local:Path;
            }
        }
        foreach ($local:Path in $this.Added) {
            if ($local:Path -and $local:Path.Trim()) {
                $local:Result += $local:Path;
            }
        }
        return $local:Result;
    }
}

# Stores a diff between two environments
class EnvironmentDiff : System.Collections.Generic.Dictionary[string,psobject] {
    EnvironmentDiff() { }

    static [EnvironmentDiff] FromObject([psobject] $Object) {
        if ($Object -eq $null) { return $null; }
        if ($Object -is [EnvironmentDiff]) { return $Object; }
        $Object = script:ConvertToHashTable $Object;
        [EnvironmentDiff] $local:Changes = [EnvironmentDiff]::new();
        foreach ($local:Entry in $Object.GetEnumerator()) {
            $local:Key = $local:Entry.Key;
            $local:Value = $local:Entry.Value;
            if ($local:Key -ieq "Path") {
                $local:Value = [PathDiff]::FromObject($local:Value);
            }
            $local:Changes[$local:Key] = $local:Value;
        }
        return $local:Changes;
    }

    static [psobject] ToObject([EnvironmentDiff] $Object) {
        if ($Object -eq $null) { return $null; }
        $local:Changes = @{};
        foreach ($local:Entry in $Object.GetEnumerator()) {
            $local:Key = $local:Entry.Key;
            $local:Value = $local:Entry.Value;
            if ($local:Key -ieq "Path") {
                $local:Value = [PathDiff]::ToObject($local:Value);
            }
            $local:Changes[$local:Key] = $local:Value;
        }
        return $local:Changes;
    }

    static [EnvironmentDiff] DiffBetween([Environment] $OldEnv, [Environment] $NewEnv) {
        [EnvironmentDiff] $local:Changes = [EnvironmentDiff]::new();
        foreach ($local:Entry in $OldEnv.GetEnumerator()) {
            if (-not $NewEnv.ContainsKey($local:Entry.Key)) {
                $local:Changes[$local:Entry.Key] = $null;
            }
        }
        foreach ($local:Entry in $NewEnv.GetEnumerator()) {
            $local:Key = $local:Entry.Key;
            $local:Value = $local:Entry.Value;
            $local:OldValue = $OldEnv[$local:Key];
            if ($local:Value -ne $local:OldValue) {
                if ($local:Key -ieq "Path") {
                    $local:Value = [PathDiff]::DiffBetween($local:OldValue, $local:Value);
                }
                $local:Changes[$local:Key] = $local:Value;
            }
        }
        return $local:Changes;
    }

    hidden [psobject] get_Item([string] $Key) {
        [psobject] $Value = $null;
        [void]($this.TryGetValue($Key, [ref]$Value));
        return $Value;
    }

    hidden [void] set_Item([string] $Key, [psobject] $Value) {
        if (-not $this.ValidateKeyValue($Key, $Value)) { return; }
        ([System.Collections.Generic.Dictionary[string, psobject]]$this)[$Key] = $Value;
    }

    hidden [void] Add([string] $Key, [psobject] $Value) {
        if (-not $this.ValidateKeyValue($Key, $Value)) { return; }
        [void](([System.Collections.Generic.Dictionary[string, psobject]]$this).Add($Key, $Value));
    }

    [Environment] Apply([Environment]$Env) {
        [Environment] $local:NewEnv = $Env.Clone();
        foreach ($local:Entry in $this.GetEnumerator()) {
            $local:Key = $local:Entry.Key;
            $local:Value = $local:Entry.Value;
            if ($local:Value -is [PathDiff]) {
                $local:Value = $local:Value.Apply($Env[$local:Key]);
            }
            if ($local:Value) {
                $local:NewEnv[$local:Key] = $local:Value;
            }
            else {
                $local:NewEnv.Remove($local:Key);
            }
        }
        return $local:NewEnv;
    }

    hidden [bool] ValidateKeyValue([string] $Key, [psobject] $Value) {
        if (($Key -ieq "Path") -and -not ($Value -eq $null -or $Value -is [PathDiff])) {
            throw [System.ArgumentException]::new("Invalid argument: Value");
            return $false;
        }
        if (($Key -ine "Path") -and -not ($Value -eq $null -or $Value -is [string])) {
            throw [System.ArgumentException]::new("Invalid argument: Value");
            return $false;
        }
        return $true;
    }
}

# Represents an instance of Visual Studio
class VisualStudioInstance {
    [string] $Name;
    [string] $Channel;
    [string] $Version;
    [string] $Path;
    hidden [EnvironmentDiff] $Env;

    VisualStudioInstance([string] $Name, [string] $Channel, [string] $Version, [string] $Path, [EnvironmentDiff] $Env) {
        $this.Name = $Name;
        $this.Channel = $Channel;
        $this.Version = $Version;
        $this.Path = $Path;
        $this.Env = $Env;
    }

    static [VisualStudioInstance] FromObject([psobject] $Object) {
        if ($Object -eq $null) { return $null; }
        if ($Object -is [VisualStudioInstance]) { return $Object; }
        return [VisualStudioInstance]::new(
            $Object.Name,
            $Object.Channel,
            $Object.Version,
            $Object.Path,
            [EnvironmentDiff]::FromObject($Object.Env)
        );
    }

    static [psobject] ToObject([VisualStudioInstance] $Object) {
        if ($Object -eq $null) { return $null; }
        return @{
            Name = $Object.Name;
            Channel = $Object.Channel;
            Version = $Object.Version;
            Path = $Object.Path;
            Env = [EnvironmentDiff]::ToObject($Object.Env);
        };
    }

    [EnvironmentDiff] GetEnvironment() {
        if ($this.Env -eq $null) {
            $local:CurrentEnv = [Environment]::GetCurrent();
            $local:DefaultEnvironment = [Environment]::GetDefault();
            $local:DefaultEnvironment.Apply();
            $local:Env = [Environment]::GetCurrent();
            $local:CommandPath = Join-Path $this.Path $script:VSDEVCMD_PATH;
            $local:Command = '"' + ($local:CommandPath) + '"&set';
            cmd /c $local:Command | ForEach-Object {
                if ($_ -match "^(.*?)=(.*)$") {
                    $local:Key = $Matches[1];
                    $local:Value = $Matches[2];
                    $local:Env[$local:Key] = $local:Value;
                }
            }
            $this.Env = [EnvironmentDiff]::DiffBetween($local:DefaultEnvironment, $local:Env);
            $local:CurrentEnv.Apply();
            $script:HasChanges = $true;
        }
        return $this.Env;
    }

    hidden [void] Apply() {
        $this.GetEnvironment().Apply([Environment]::GetDefault()).Apply();
    }

    [void] Save() {
        $script:HasChanges = $true;
        script:SaveChanges;
    }
}

# Converts a JSON object (from ConvertFrom-Json) into a Hashtable
function script:ConvertToHashTable([psobject] $Object) {
    if ($Object -eq $null) { return $null; }
    if ($Object -is [hashtable]) { return $Object };
    $local:Table = @{};
    foreach ($local:Key in $Object | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name) {
        $local:Value = $Object | Select-Object -ExpandProperty $local:Key;
        $local:Table[$local:Key] = $local:Value;
    }
    return $local:Table;
}

# Sets or removes an environment variable
function script:SetEnvironmentVariable([string] $Key, [string] $Value) {
    if ($Value -ne $null) {
        [void](Set-Item -Force "ENV:\$Key" -Value $Value);
    }
    else {
        [void](Remove-Item -Force "ENV:\$Key");
    }
}

# Populates $script:VisualStudioVersions from cache if it is empty
function script:PopulateVisualStudioVersionsFromCache() {
    if ($script:VisualStudioVersions -eq $null) {
        if (Test-Path $script:CACHE_PATH) {
            $script:VisualStudioVersions = (Get-Content $script:CACHE_PATH | ConvertFrom-Json) `
                | ForEach-Object {
                    [VisualStudioInstance]::FromObject($_);
                };
        }
    }
}

# Populates $script:VisualStudioVersions from disk if it is empty
function script:PopulateVisualStudioVersions() {
    if ($script:VisualStudioVersions -eq $null) {
        # Add Legacy instances
        $script:VisualStudioVersions = Get-ChildItem ${env:ProgramFiles(x86)} `
            | Where-Object -Property Name -Match "Microsoft Visual Studio (\d+.0)" `
            | ForEach-Object {
                [VisualStudioInstance]::new(
                    $Matches[0],
                    "Release",
                    $Matches[1],
                    $_.FullName,
                    $null
                );
            };

        # Add Dev15+ instances
        if (Test-Path $script:VS_INSTANCES_DIR) {
            $script:VisualStudioVersions += Get-ChildItem $script:VS_INSTANCES_DIR `
                | ForEach-Object {
                    $local:StatePath = Join-Path $_.FullName "state.json";
                    $local:State = Get-Content $local:StatePath | ConvertFrom-Json;
                    [VisualStudioInstance]::new(
                        $local:State.installationName,
                        $local:State.channelId,
                        $local:State.installationVersion,
                        $local:State.installationPath,
                        $null
                    );
                };
        }

        # Sort by version descending and remove versions that don't exist
        $script:VisualStudioVersions = $script:VisualStudioVersions `
            | Sort-Object -Property Version -Descending `
            | Where-Object { Test-Path (Join-Path $_.Path $script:VSDEVCMD_PATH) };

        if ($script:VisualStudioVersions) {
            $script:HasChanges = $true;
        }
    }
}

# Saves any changes to the $script:VisualStudioVersions cache to disk
function script:SaveChanges() {
    if ($script:HasChanges -and $script:VisualStudioVersions) {
        $local:Content = $script:VisualStudioVersions `
            | ForEach-Object {
                [VisualStudioInstance]::ToObject($_);
            } `
            | ConvertTo-Json;
        if ($script:VisualStudioVersions.Length -eq 1) {
            $local:Content = "[" + $local:Content + "]";
        }
        $local:CacheDir = Split-Path $script:CACHE_PATH -Parent;
        if (-not (Test-Path $local:CacheDir)) {
            [void](mkdir $local:CacheDir -ErrorAction:SilentlyContinue);
        }

        $local:Content | Out-File $script:CACHE_PATH;
        $script:HasChanges = $false;
    }
}

# Indicates whether the specified profile path exists
function script:HasProfile([string] $ProfilePath) {
    if (-not $ProfilePath) { return $false; }
    if (-not (Test-Path -LiteralPath $ProfilePath)) { return $false; }
    return $true;
}

# Indicates whether "posh-vsdev" is referenced in the specified profile
function script:IsInProfile([string] $ProfilePath) {
    if (-not (script:HasProfile $ProfilePath)) { return $false; }
    $local:Content = Get-Content $ProfilePath -ErrorAction:SilentlyContinue;
    if ($local:Content -match "posh-vsdev") { return $true; }
    return $false;
}

# Indicates whether the Use-VisualStudioEnvironment cmdlet is referenced int he specified profile
function script:IsUsingEnvironment([string] $ProfilePath) {
    if (-not (script:HasProfile $ProfilePath)) { return $false; }
    $local:Content = Get-Content $ProfilePath -ErrorAction:SilentlyContinue;
    if ($local:Content -match "Use-VisualStudioEnvironment") { return $true; }
    return $false;
}

# Indicates whether the specified profile is signed
function script:IsProfileSigned([string] $ProfilePath) {
    if (-not (script:HasProfile $ProfilePath)) { return $false; }
    $local:Sig = Get-AuthenticodeSignature $ProfilePath;
    if (-not $local:Sig) { return $false; }
    if (-not $local:Sig.SignerCertificate) { return $false; }
    return $true;
}

# Indicates whether this module is installed within a PowerShell common module path
function script:IsInModulePaths() {
    foreach ($local:Path in $env:PSModulePath -split ";") {
        if (-not $local:Path.EndsWith("\")) { $local:Path += "\"; }
        if ($PSScriptRoot.StartsWith($local:Path, [System.StringComparison]::InvariantCultureIgnoreCase)) {
            return $true;
        }
    }
    return $false;
}

<#
.SYNOPSIS
    Get installed Visual Studio instances.
.DESCRIPTION
    The Get-VisualStudioVersion cmdlet gets information about the installed Visual Studio instances on this machine.
.PARAMETER Name
    Specifies a name that can be used to filter the results.
.PARAMETER Channel
    Specifies a release channel that can be used to filter the results.
.PARAMETER Version
    Specifies a version number that can be used to filter the results.
.INPUTS
    None. You cannot pipe objects to Get-VisualStudioVersion.
.OUTPUTS
    VisualStudioInstance. Get-VisualStudioVersion returns a VisualStudioInstance object for each matching instance.
.EXAMPLE
    PS> Get-VisualStudioVersion
    Name                                              Channel                    Version      Path
    ----                                              -------                    -------      ----
    VisualStudio/15.0.0+26228.9.d15rtwsvc             VisualStudio.15.int.d15rel 15.0.26228.9 C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise
    Microsoft Visual Studio 14.0                      Release                    14.0         C:\Program Files (x86)\Microsoft Visual Studio 14.0
.EXAMPLE
    PS> Get-VisualStudioVersion -Channel Release
    Name                                              Channel                    Version      Path
    ----                                              -------                    -------      ----
    Microsoft Visual Studio 14.0                      Release                    14.0         C:\Program Files (x86)\Microsoft Visual Studio 14.0
#>
function Get-VisualStudioVersion([string] $Name, [string] $Channel, [string] $Version) {
    script:PopulateVisualStudioVersionsFromCache;
    script:PopulateVisualStudioVersions;
    $local:Versions = $script:VisualStudioVersions;
    if ($Name) {
        $local:Versions = $local:Versions | Where-Object -Property Name -Like $Name;
    }
    if ($Channel) {
        $local:Versions = $local:Versions | Where-Object -Property Channel -Like $Channel;
    }
    if ($Version) {
        $local:Versions = $local:Versions | Where-Object -Property Version -Like $Version;
    }
    $local:Versions;
    script:SaveChanges;
}

<#
.SYNOPSIS
    Uses the developer environment variables for an instance of Visual Studio.
.DESCRIPTION
    The Use-VisualStudioEnvironment cmdlet overwrites the current environment variables with ones from the
    Developer Command Prompt for a specific instance of Visual Studio.
    If a developer environment is already in use, the environment is first reset to the state at the time
    the "posh-vsdev" module was loaded.
.PARAMETER Name
    Specifies a name that can be used to filter the results.
.PARAMETER Channel
    Specifies a release channel that can be used to filter the results.
.PARAMETER Version
    Specifies a version number that can be used to filter the results.
.PARAMETER InputObject
    A VisualStudioInstance whose environment should be used.
.INPUTS
    VisualStudioInstance.
        You can pipe a VisualStudioInstance to Use-VisualStudioEnvironment.
.OUTPUTS
    None.
.EXAMPLE
    PS> Use-VisualStudioVersion
    Using Development Environment from 'Microsoft Visual Studio 14.0'.
#>
function Use-VisualStudioEnvironment {
    [CmdletBinding()]
    param (
        [Parameter(ParameterSetName = "Match")]
        [string] $Name,
        [Parameter(ParameterSetName = "Match")]
        [string] $Channel,
        [Parameter(ParameterSetName = "Match")]
        [version] $Version,
        [Parameter(ParameterSetName = "Pipeline", Position = 0, ValueFromPipeline = $true, Mandatory = $true)]
        [psobject] $InputObject
    );

    [void]([Environment]::GetDefault());
    [VisualStudioInstance] $local:Instance = $null;
    if ($InputObject) {
        $local:Instance = [VisualStudioInstance]::FromObject($InputObject);
    } else {
        $local:Instance = Get-VisualStudioVersion -Name:$Name -Channel:$Channel -Version:$Version | Select-Object -First:1;
    }

    if ($local:Instance) {
        $local:Instance.Apply();
        script:SaveChanges;
        Write-Host "Using Development Environment from '$($local:Instance.Name)'." -ForegroundColor:DarkGray;
        $script:VisualStudioVersion = $local:Instance;
    }
    else {
        [string] $local:Message = "Could not find Visual Studio";
        [string[]] $local:MessageParts = @();
        if ($Name) { $local:MessageParts += "Name='$Name'"; }
        if ($Channel) { $local:MessageParts += "Channel='$Channel'"; }
        if ($Version) { $local:MessageParts += "Version='$Version'"; }
        if ($local:MessageParts.Length > 0) {
            $local:Message += "for " + $local:MessageParts[0];
            if ($local:MessageParts.Length -eq 2) {
            }
            elseif ($local:MessageParts.Length -gt 2) {
                for ($local:I = 1; $local:I -lt $local:MessageParts.Length - 1; $local:I++) {
                    $local:Message += ", " + $local:MessageParts[$local:I];
                }
                if ($local:MessageParts.Length > 2) {
                    $local:Message += ", and " + $local:MessageParts[$local:MessageParts.Length - 1];
                }
            }
        }
        $local:Message += ".";
        Write-Warning $local:Message;
    }
}

<#
.SYNOPSIS
    Restores the original enironment.
.DESCRIPTION
    The Reset-VisualStudioEnvironment cmdlet restores all environment variables to their values
    at the point the "posh-vsdev" module was first imported.
.PARAMETER Force
    Indicates that all environment variables should be restored even if no development environment
    was used.
.INPUTS
    None. You cannot pipe objects to Reset-VisualStudioEnvironment.
.OUTPUTS
    None.
#>
function Reset-VisualStudioEnvironment([switch] $Force) {
    if ($script:VisualStudioVersion -or $Force) {
        $script:VisualStudioVersion = $null;
        [Environment]::GetDefault().Apply();
        Write-Host "Restored default environment" -ForegroundColor DarkGray;
    }
}

<#
.SYNOPSIS
    Resets the cache of installed Visual Studio instances.
.DESCRIPTION
    Resets the cache of installed Visual Studio instances and their respective environment
    settings.
.INPUTS
    None. You cannot pipe objects to Reset-VisualStudioVersionCache.
.OUTPUTS
    None.
#>
function Reset-VisualStudioVersionCache() {
    $script:VisualStudioVersions = $null;
    if (Test-Path $script:CACHE_PATH) {
        [void](Remove-Item $script:CACHE_PATH -Force);
    }
}

<#
.SYNOPSIS
    Adds "posh-vsdev" to your profile.
.DESCRIPTION
    Adds an import to "posh-vsdev" to your PowerShell profile.
.PARAMETER AllHosts
    Specifies that "posh-vsdev" should be installed to your PowerShell profile for all PowerShell hosts.
    If not provided, only the current profile is used.
.PARAMETER UseEnvironment
    Specifies that an invocation of the Use-VisualStudioEnvironment cmdlet should be added to your
    PowerShell profile.
.PARAMETER Force
    Indicates that "posh-vsdev" should be added to your profile, even if it may already be present.
.INPUTS
    None. You cannot pipe objects to Reset-VisualStudioVersionCache.
.OUTPUTS
    None.
#>
function Add-VisualStudioEnvironmentToProfile([switch] $AllHosts, [switch] $UseEnvironment, [switch] $Force) {
    [string] $local:ProfilePath = if ($AllHosts) { $profile.CurrentUserAllHosts; } else { $profile.CurrentUserCurrentHost; }
    [bool] $local:IsInProfile = script:IsInProfile $local:ProfilePath;
    [bool] $local:IsUsingEnvironment = script:IsUsingEnvironment $local:ProfilePath;
    if (-not $Force -and $local:IsInProfile -and -not $UseEnvironment) {
        Write-Warning "'posh-vsdev' is already installed.";
        return;
    }
    if (-not $Force -and $local:IsUsingEnvironment -and $UseEnvironment) {
        Write-Warning "'posh-vsdev' is already using a VisualStudio environment.";
        return;
    }
    if (script:IsProfileSigned $local:ProfilePath) {
        Write-Warning "Cannot modify signed profile.";
        return;
    }
    [string] $local:Content = $null;
    if ($Force -or -not $local:IsInProfile) {
        if (-not (script:HasProfile $local:ProfilePath)) {
            $local:ProfileDir = Split-Path $local:ProfilePath -Parent;
            if (-not (Test-Path -LiteralPath:$local:ProfileDir)) {
                [void](mkdir $local:ProfileDir -ErrorAction:SilentlyContinue);
            }
        }
        if (script:IsInModulePaths) {
            $local:Content += "`nImport-Module posh-vsdev;";
        }
        else {
            $local:Content += "`nImport-Module `"$PSScriptRoot\posh-vsdev.psd1`";";
        }
    }
    if ($Force -or (-not $local:IsUsingEnvironment -and $UseEnvironment)) {
        $local:Content += "`nUse-VisualStudioEnvironment;";
    }
    if ($local:Content) {
        Add-Content -LiteralPath:$local:ProfilePath -Value $local:Content -Encoding UTF8;
    }
}

# constants
[string] $script:VSDEVCMD_PATH = "Common7\Tools\VsDevCmd.bat";
[string] $script:VS_INSTANCES_DIR = "$env:ProgramData\Microsoft\VisualStudio\Packages\_Instances";
[string] $script:CONFIG_DIR = "$env:USERPROFILE\.posh-vsdev";
[string] $script:CACHE_PATH = "$script:CONFIG_DIR\instances.json";

# state
[bool] $script:HasChanges = $false;                             # Indicates whether the in-memory cache has changes
[VisualStudioInstance[]] $script:VisualStudioVersions = $null;  # In-memory cache of instances
[VisualStudioInstance] $script:VisualStudioVersion = $null;     # Current VS instance

# Save the default environment.
[void]([Environment]::GetDefault());

# Reset the environment when the module is removed
$ExecutionContext.SessionState.Module.OnRemove = {
    if ($script:VisualStudioVersion) {
        Reset-VisualStudioEnvironment;
    }
};

# Export members
Export-ModuleMember `
    -Function:(
        'Get-VisualStudioVersion',
        'Use-VisualStudioEnvironment',
        'Reset-VisualStudioEnvironment',
        'Reset-VisualStudioVersionCache',
        'Add-VisualStudioEnvironmentToProfile'
    ) `
    -Variable:(
        'VisualStudioVersion'
    );