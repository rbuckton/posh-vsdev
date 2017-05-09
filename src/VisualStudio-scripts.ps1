# constants
$script:VSDEVCMD_PATH = "Common7\Tools\VsDevCmd.bat";
$script:VS_INSTANCES_DIR = "$env:ProgramData\Microsoft\VisualStudio\Packages\_Instances";
$script:POSH_VS_CONFIG_DIR = "$env:USERPROFILE\.posh-vs";
$script:CACHE_PATH = "$script:POSH_VS_CONFIG_DIR\instances.json";

$script:VisualStudioVersions = $null;   # In-memory cache of instances
$script:HasChanges = $false;            # Indicates whether the in-memory cache has changes

# simplifies access to HashSet<string>
class Set : System.Collections.Generic.HashSet[string] {
    Set() { }
    Set([string[]] $Data) {
        foreach($local:Item in $Data) {
            $this.Add($local:Item);
        }
    }
}

class Env : System.Collections.Generic.Dictionary[string,string] {
    hidden static [Env] $_Default;

    Env() {}
    Env([Env] $Other) {
        if ($Other) {
            foreach ($local:Entry in $Other.GetEnumerator()) {
                $this[$local:Entry.Key] = $Other[$local:Entry.Value];
            }
        }
    }

    static [Env] GetDefault() {
        if ([Env]::_Default -eq $null) {
            [Env]::_Default = [Env]::GetCurrent();
        }
        return [Env]::_Default;
    }

    static [Env] GetCurrent() {
        $local:Env = [Env]::new();
        foreach($local:Item in Get-ChildItem "ENV:\") {
            $local:Env[$local:Item.Name] = $local:Item.Value;
        }
        return $local:Env;
    }

    [string] get_Item([string] $Key) {
        $Value = $null;
        [void]($this.TryGetValue($Key, [ref]$Value));
        return $Value;
    }

    [void] Apply() {
        [void]([Env]::GetDefault());
        $local:Current = [Env]::GetCurrent();
        foreach ($local:Item in $local:Current.GetEnumerator()) {
            if (-not $this.ContainsKey($local:Item.Key)) {
                script:SetEnvironmentVariable $local:Item.Key $null;
            }
        }
        foreach ($local:Item in $this.GetEnumerator()) {
            script:SetEnvironmentVariable $local:Item.Key $local:Item.Value;
        }
    }
}

class PathsDiff {
    hidden [string[]] $Added;
    hidden [string[]] $Removed;
    hidden [Set] $RemovedSet;

    PathsDiff([string[]] $Added, [string[]] $Removed) {
        $this.Added = @() + $Added;
        $this.Removed = @() + $Removed;
        $this.RemovedSet = [Set]::new($Removed);
    }

    static [PathsDiff] FromObject([psobject] $Object) {
        if ($Object -eq $null) { return $null; }
        if ($Object -is [PathsDiff]) { return $Object; }
        return [PathsDiff]::new($Object.Added, $Object.Removed);
    }

    static [psobject] ToObject([PathsDiff] $Object) {
        if ($Object -eq $null) { return $null; }
        return @{
            Added = @() + $Object.Added;
            Removed = @() + $Object.Removed;
        };
    }

    static [PathsDiff] DiffBetween([string[]] $OldPaths, [string[]] $NewPaths) {
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
        return [PathsDiff]::new($local:Added, $local:Removed);
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
            if (-not $this.RemovedSet.Contains($local:Path)) {
                $local:Result += $local:Path;
            }
        }
        return $local:Result + $this.Added;
    }
}

class Diff : System.Collections.Generic.Dictionary[string,psobject] {
    Diff() { }

    static [Diff] FromObject([psobject] $Object) {
        if ($Object -eq $null) { return $null; }
        if ($Object -is [Diff]) { return $Object; }
        $Object = script:ConvertToHashTable $Object;
        [Diff] $local:Changes = [Diff]::new();
        foreach ($local:Entry in $Object.GetEnumerator()) {
            $local:Key = $local:Entry.Key;
            $local:Value = $local:Entry.Value;
            if ($local:Key -ieq "Path") {
                $local:Value = [PathsDiff]::FromObject($local:Value);
            }
            $local:Changes[$local:Key] = $local:Value;
        }
        return $local:Changes;
    }

    static [psobject] ToObject([Diff] $Object) {
        if ($Object -eq $null) { return $null; }
        $local:Changes = @{};
        foreach ($local:Entry in $Object.GetEnumerator()) {
            $local:Key = $local:Entry.Key;
            $local:Value = $local:Entry.Value;
            if ($local:Key -ieq "Path") {
                $local:Value = [PathsDiff]::ToObject($local:Value);
            }
            $local:Changes[$local:Key] = $local:Value;
        }
        return $local:Changes;
    }

    static [Diff] DiffBetween([Env] $OldEnv, [Env] $NewEnv) {
        [Diff] $local:Changes = [Diff]::new();
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
                    $local:Value = [PathsDiff]::DiffBetween($local:OldValue, $local:Value);
                }
                $local:Changes[$local:Key] = $local:Value;
            }
        }
        return $local:Changes;
    }

    [psobject] get_Item([string] $Key) {
        [psobject] $Value = $null;
        [void]($this.TryGetValue($Key, [ref]$Value));
        return $Value;
    }

    [void] set_Item([string] $Key, [psobject] $Value) {
        if (-not $this.ValidateKeyValue($Key, $Value)) { return; }
        ([System.Collections.Generic.Dictionary[string, psobject]]$this)[$Key] = $Value;
    }

    [void] Add([string] $Key, [psobject] $Value) {
        if (-not $this.ValidateKeyValue($Key, $Value)) { return; }
        [void](([System.Collections.Generic.Dictionary[string, psobject]]$this).Add($Key, $Value));
    }

    [Env] Apply([Env]$Env) {
        [Env] $local:NewEnv = [Env]::new($Env);
        foreach ($local:Entry in $this.GetEnumerator()) {
            $local:Key = $local:Entry.Key;
            $local:Value = $local:Entry.Value;
            if ($local:Value -is [PathsDiff]) {
                $local:Value = $local:Value.Apply($local:NewEnv[$local:Key]);
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
        if (($Key -ieq "Path") -and -not ($Value -eq $null -or $Value -is [PathsDiff])) {
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

class Instance {
    [string] $Name;
    [string] $Channel;
    [string] $Version;
    [string] $BasePath;
    [string] $CommandPath;
    hidden [Diff] $Env;

    Instance([string] $Name, [string] $Channel, [string] $Version, [string] $BasePath, [string] $CommandPath, [Diff] $Env) {
        $this.Name = $Name;
        $this.Channel = $Channel;
        $this.Version = $Version;
        $this.BasePath = $BasePath;
        $this.CommandPath = $CommandPath;
        $this.Env = $Env;
    }

    static [Instance] FromObject([psobject] $Object) {
        if ($Object -eq $null) { return $null; }
        if ($Object -is [Instance]) { return $Object; }
        return [Instance]::new(
            $Object.Name,
            $Object.Channel,
            $Object.Version,
            $Object.BasePath,
            $Object.CommandPath,
            [Diff]::FromObject($Object.Env)
        );
    }

    static [psobject] ToObject([Instance] $Object) {
        if ($Object -eq $null) { return $null; }
        return @{
            Name = $Object.Name;
            Channel = $Object.Channel;
            Version = $Object.Version;
            BasePath = $Object.BasePath;
            CommandPath = $Object.CommandPath;
            Env = [Diff]::ToObject($Object.Env);
        };
    }

    [Diff] GetEnvironment() {
        if ($this.Env -eq $null) {
            $local:CurrentEnv = [Env]::GetCurrent();
            $local:DefaultEnvironment = [Env]::GetDefault();
            $local:DefaultEnvironment.Apply();
            $local:Env = [Env]::new();
            $local:Command = '"' + ($this.CommandPath) + '"&set';
            cmd /c $local:Command | ForEach-Object {
                if ($_ -match "^(.*?)=(.*)$") {
                    $local:Key = $Matches[1];
                    $local:Value = $Matches[2];
                    $local:Env[$local:Key] = $local:Value;
                }
            }
            $local:CurrentEnv.Apply();
            $this.Env = [Diff]::DiffBetween($local:DefaultEnvironment, $local:Env);
            $script:HasChanges = $true;
        }
        return $this.Env;
    }

    [void] Apply() {
        $local:Def = [Env]::GetDefault();
        $local:Diff = $this.GetEnvironment();
        $local:Env = $local:Diff.Apply($local:Def);
        $local:Env.Apply();
    }
}

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

function script:SetEnvironmentVariable([string] $Key, [string] $Value) {
    if ($Value -ne $null) {
        [void](Set-Item -Force "ENV:\$Key" -Value $Value);
    }
    else {
        [void](Remove-Item -Force "ENV:\$Key");
    }
}

function script:PopulateVisualStudioVersionsFromCache() {
    if ($script:VisualStudioVersions -eq $null) {
        if (Test-Path $script:CACHE_PATH) {
            $script:VisualStudioVersions = (Get-Content $script:CACHE_PATH | ConvertFrom-Json) `
                | ForEach-Object {
                    [Instance]::FromObject($_);
                };
        }
    }
}

function script:PopulateVisualStudioVersions() {
    if ($script:VisualStudioVersions -eq $null) {
        # Add Legacy instances
        $script:VisualStudioVersions = Get-ChildItem ${env:ProgramFiles(x86)} `
            | Where-Object -Property Name -Match "Microsoft Visual Studio (\d+.0)" `
            | ForEach-Object {
                [Instance]::new(
                    $Matches[0],
                    "Release",
                    $Matches[1],
                    $_.FullName,
                    (Join-Path $_.FullName $script:VSDEVCMD_PATH),
                    $null
                );
            };

        # Add Dev15+ instances
        if (Test-Path $script:VS_INSTANCES_DIR) {
            $script:VisualStudioVersions += Get-ChildItem $script:VS_INSTANCES_DIR `
                | ForEach-Object {
                    $local:StatePath = Join-Path $_.FullName "state.json";
                    $local:State = Get-Content $local:StatePath | ConvertFrom-Json;
                    [Instance]::new(
                        $local:State.installationName,
                        $local:State.channelId,
                        $local:State.installationVersion,
                        $local:State.installationPath,
                        (Join-Path $local:State.installationPath $script:VSDEVCMD_PATH),
                        $null
                    );
                };
        }

        # Sort by version descending and remove versions that don't exist
        $script:VisualStudioVersions = $script:VisualStudioVersions `
            | Sort-Object -Property Version -Descending `
            | Where-Object { Test-Path $_.CommandPath };

        if ($script:VisualStudioVersions) {
            $script:HasChanges = $true;
        }
    }
}

function script:SaveChanges() {
    if ($script:HasChanges -and $script:VisualStudioVersions) {
        $local:Content = $script:VisualStudioVersions `
            | ForEach-Object {
                [Instance]::ToObject($_);
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

    [void]([Env]::GetDefault());
    [Instance] $local:VisualStudioVersion = $null;
    if ($InputObject) {
        $local:VisualStudioVersion = [Instance]::FromObject($InputObject);
    } else {
        $local:VisualStudioVersion = Get-VisualStudioVersion -Name:$Name -Channel:$Channel -Version:$Version | Select-Object -First:1;
    }

    if ($local:VisualStudioVersion) {
        $local:VisualStudioVersion.Apply();
        script:SaveChanges;
        Write-Output "Using Development Environment from '$($local:VisualStudioVersion.Name)'.";
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

function Reset-VisualStudioEnvironment {
    [Env]::GetDefault().Apply();
}

function Reset-VisualStudioVersionCache() {
    $script:VisualStudioVersions = $null;
    if (Test-Path $script:CACHE_PATH) {
        [void](Remove-Item $script:CACHE_PATH -Force);
    }
}

function script:HasProfile([string] $ProfilePath) {
    if (-not $ProfilePath) { return $false; }
    if (-not (Test-Path -LiteralPath $ProfilePath)) { return $false; }
    return $true;
}

function script:IsInProfile([string] $ProfilePath) {
    if (-not (script:HasProfile $ProfilePath)) { return $false; }
    $local:Content = Get-Content $ProfilePath -ErrorAction:SilentlyContinue;
    if ($local:Content -match "posh-vs") { return $true; }
    return $false;
}

function script:IsUsingEnvironment([string] $ProfilePath) {
    if (-not (script:HasProfile $ProfilePath)) { return $false; }
    $local:Content = Get-Content $ProfilePath -ErrorAction:SilentlyContinue;
    if ($local:Content -match "Use-VisualStudioEnvironment") { return $true; }
    return $false;
}

function script:IsProfileSigned([string] $ProfilePath) {
    if (-not (script:HasProfile $ProfilePath)) { return $false; }
    $local:Sig = Get-AuthenticodeSignature $ProfilePath;
    if (-not $local:Sig) { return $false; }
    if (-not $local:Sig.SignerCertificate) { return $false; }
    return $true;
}

function script:IsInModulePaths() {
    foreach ($local:Path in $env:PSModulePath -split ";") {
        if (-not $local:Path.EndsWith("\")) { $local:Path += "\"; }
        if ($PSScriptRoot.StartsWith($local:Path, [System.StringComparison]::InvariantCultureIgnoreCase)) {
            return $true;
        }
    }
    return $false;
}

function Add-VisualStudioEnvironmentToProfile([switch] $AllHosts, [switch] $UseEnvironment) {
    $local:ProfilePath = if ($AllHosts) { $profile.CurrentUserAllHosts; } else { $profile.CurrentUserCurrentHost; }
    $local:IsInProfile = script:IsInProfile $local:ProfilePath;
    $local:IsUsingEnvironment = script:IsUsingEnvironment $local:ProfilePath;
    if ($local:IsInProfile -and -not $UseEnvironment) {
        Write-Warning "'posh-vs' is already installed.";
        return;
    }
    if ($local:IsUsingEnvironment -and $UseEnvironment) {
        Write-Warning "'posh-vs' is already using a VisualStudio environment.";
        return;
    }
    if (script:IsProfileSigned $local:ProfilePath) {
        Write-Warning "Cannot modify signed profile.";
        return;
    }
    if (-not $local:IsInProfile) {
        if (-not (script:HasProfile $local:ProfilePath)) {
            $local:ProfileDir = Split-Path $local:ProfilePath -Parent;
            if (-not (Test-Path -LiteralPath:$local:ProfileDir)) {
                [void](mkdir $local:ProfileDir -ErrorAction:SilentlyContinue);
            }
        }
        if (script:IsInModulePaths) {
            Add-Content -LiteralPath:$local:ProfilePath -Value "Import-Module posh-vs;" -Encoding UTF8;
        }
        else {
            Add-Content -LiteralPath:$local:ProfilePath -Value "Import-Module `"$PSScriptRoot\posh-vs.psd1`";" -Encoding UTF8;
        }
    }
    if (-not $local:IsUsingEnvironment -and $UseEnvironment) {
        Add-Content -LiteralPath:$local:ProfilePath -Value "Use-VisualStudioEnvironment;" -Encoding UTF8;
    }
}

[void]([Env]::GetDefault());