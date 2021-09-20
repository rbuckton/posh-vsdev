# Simplifies access to HashSet<string>
class Set : System.Collections.Generic.HashSet[string] {
    Set() { }
    Set([string[]] $Data) {
        foreach ($local:Item in $Data) {
            $this.Add($local:Item);
        }
    }
}

# Encapsulates environment variables and their values
class Environment : System.Collections.Generic.Dictionary[string, string] {
    hidden static [Environment] $_Clean;
    hidden static [Environment] $_Default;

    # Get a clean environment
    static [Environment] GetClean() {
        if ($null -eq [Environment]::_Clean) {
            $local:Entries = script:ExecuteCommandInNewEnvironment { Get-ChildItem env:; };
            $local:Env = [Environment]::new();
            foreach ($local:Entry in $local:Entries) {
                if (script:IsIgnoredEnvironmentVariable $local:Entry.Key) {
                    continue;
                }
                $local:Env[$local:Entry.Name] = $local:Entry.Value;
            }
            [Environment]::_Clean = $local:Env;
        }
        return [Environment]::_Clean;
    }

    # Get the default environment without VS environment variables
    static [Environment] GetDefault() {
        if ($null -eq [Environment]::_Default) {
            [Environment] $local:CleanEnv = [Environment]::GetClean();
            [Environment] $local:Env = $null;
            if ($env:PoshVsDevDefaultEnvironment) {
                $local:EnvDiffObject = $env:PoshVsDevDefaultEnvironment | ConvertFrom-Json -ErrorAction:SilentlyContinue;
                if ($local:EnvDiffObject) {
                    $local:EnvDiff = [EnvironmentDiff]::FromObject($local:EnvDiffObject);
                    $local:Env = $local:EnvDiff.Apply($local:CleanEnv);
                }
            }
            if (-not $local:Env -and $env:PoshVsDevEnvironment) {
                if ($script:VisualStudioVersion) {
                    $local:Env = $script:VisualStudioVersion.Unapply([Environment]::GetCurrent());
                }
                else {
                    $local:Env = $local:CleanEnv;
                }
            }
            if (-not $local:Env) {
                $local:CurrentEnv = [Environment]::GetCurrent();
                $local:EnvDiff = [EnvironmentDiff]::DiffBetween($local:CleanEnv, $local:CurrentEnv);
                $local:EnvDiffObject = [EnvironmentDiff]::ToObject($local:EnvDiff);
                $local:Env = $local:CurrentEnv;
                $env:PoshVsDevDefaultEnvironment = $local:EnvDiffObject | ConvertTo-Json;
            }
            [Environment]::_Default = $local:Env;
        }
        return [Environment]::_Default;
    }

    # Gets the current environment (excluding PoshVsDev* environment variables)
    static [Environment] GetCurrent() {
        $local:Env = [Environment]::new();
        foreach ($local:Entry in Get-ChildItem "ENV:\") {
            if (script:IsIgnoredEnvironmentVariable $local:Entry.Key) {
                continue;
            }
            $local:Env[$local:Entry.Name] = $local:Entry.Value;
        }
        return $local:Env;
    }

    hidden [string] get_Item([string] $Key) {
        $Value = $null;
        [void]($this.TryGetValue($Key, [ref]$Value));
        return $Value;
    }

    # Applies the this environment's variables, replacing the current environment.
    [void] Apply([string] $Name) {
        # Clear the current environment
        $local:Current = [Environment]::GetCurrent();
        foreach ($local:Entry in $local:Current.GetEnumerator()) {
            if (script:IsIgnoredEnvironmentVariable $local:Entry.Key) {
                continue;
            }
            if (-not $this.ContainsKey($local:Entry.Key)) {
                script:SetEnvironmentVariable $local:Entry.Key $null;
            }
        }

        # Apply the new environment.
        foreach ($local:Entry in $this.GetEnumerator()) {
            if (script:IsIgnoredEnvironmentVariable $local:Entry.Key) {
                continue;
            }
            script:SetEnvironmentVariable $local:Entry.Key $local:Entry.Value;
        }

        # Set a PoshVsDevEnvironment variable for the current environment
        # (helps improve startup time in a nested shell)
        script:SetEnvironmentVariable "PoshVsDevEnvironment" $Name;
    }

    [Environment] Clone() {
        [Environment] $local:Env = [Environment]::new();
        foreach ($local:Entry in $this.GetEnumerator()) {
            if (script:IsIgnoredEnvironmentVariable $local:Entry.Key) {
                continue;
            }
            $local:Env[$local:Entry.Key] = $local:Entry.Value;
        }
        return $local:Env;
    }
}

# Stores a diff between two paths
class PathDiff {
    hidden [string[]] $Added;
    hidden [Set] $AddedSet;
    hidden [string[]] $Removed;
    hidden [Set] $RemovedSet;

    hidden PathDiff([string[]] $Added, [string[]] $Removed) {
        $this.Added = @() + $Added;
        $this.AddedSet = [Set]::new($Added);
        $this.Removed = @() + $Removed;
        $this.RemovedSet = [Set]::new($Removed);
    }

    static [PathDiff] FromObject([psobject] $Object) {
        if ($null -eq $Object) { return $null; }
        if ($Object -is [PathDiff]) { return $Object; }
        return [PathDiff]::new($Object.Added, $Object.Removed);
    }

    static [psobject] ToObject([PathDiff] $Object) {
        if ($null -eq $Object) { return $null; }
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

    [string] Unapply([string] $Path) {
        return $this.UnapplyToPaths($Path -split ";") -join ";";
    }

    [string[]] Unapply([string[]] $Paths) {
        return $this.UnapplyToPaths($Paths);
    }

    hidden [string[]] UnapplyToPaths([string[]] $Paths) {
        $local:Result = @();
        foreach ($local:Path in $Paths) {
            if ($local:Path -and $local:Path.Trim() -and -not $this.AddedSet.Contains($local:Path)) {
                $local:Result += $local:Path;
            }
        }
        foreach ($local:Path in $this.Removed) {
            if ($local:Path -and $local:Path.Trim()) {
                $local:Result += $local:Path;
            }
        }
        return $local:Result;
    }
}

# Stores a diff between two environments
class EnvironmentDiff : System.Collections.Generic.Dictionary[string, psobject] {
    EnvironmentDiff() {
    }

    # Create an EnvironmentDiff from a psobject (for deserialization purposes)
    static [EnvironmentDiff] FromObject([psobject] $Object) {
        if ($null -eq $Object) {
            return $null;
        }

        if ($Object -is [EnvironmentDiff]) {
            return $Object;
        }

        $Object = script:ConvertToHashTable $Object;
        [EnvironmentDiff] $local:Changes = [EnvironmentDiff]::new();
        foreach ($local:Entry in $Object.GetEnumerator()) {
            if (script:IsIgnoredEnvironmentVariable $local:Entry.Key) {
                continue;
            }
            $local:Key = $local:Entry.Key;
            $local:Value = $local:Entry.Value;
            if ($local:Key -ieq "Path") {
                $local:Value = [PathDiff]::FromObject($local:Value);
            }
            $local:Changes[$local:Key] = $local:Value;
        }
        return $local:Changes;
    }

    # Creates a psobject from an EnvironmentDiff (for serialization purposes)
    static [psobject] ToObject([EnvironmentDiff] $Object) {
        if ($null -eq $Object) {
            return $null;
        }

        $local:Changes = @{};
        foreach ($local:Entry in $Object.GetEnumerator()) {
            if (script:IsIgnoredEnvironmentVariable $local:Entry.Key) {
                continue;
            }
            $local:Key = $local:Entry.Key;
            $local:Value = $local:Entry.Value;
            if ($local:Key -ieq "Path") {
                $local:Value = [PathDiff]::ToObject($local:Value);
            }
            $local:Changes[$local:Key] = $local:Value;
        }
        return $local:Changes;
    }

    # Calculates the difference between two Environments
    static [EnvironmentDiff] DiffBetween([Environment] $OldEnv, [Environment] $NewEnv) {
        [EnvironmentDiff] $local:Changes = [EnvironmentDiff]::new();
        foreach ($local:Entry in $OldEnv.GetEnumerator()) {
            if (script:IsIgnoredEnvironmentVariable $local:Entry.Key) {
                continue;
            }
            if (-not $NewEnv.ContainsKey($local:Entry.Key)) {
                $local:Changes[$local:Entry.Key] = $null;
            }
        }
        foreach ($local:Entry in $NewEnv.GetEnumerator()) {
            if (script:IsIgnoredEnvironmentVariable $local:Entry.Key) {
                continue;
            }
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

    # Applies this EnvironmentDiff to the provided Environment, producing a new Environment
    [Environment] Apply([Environment] $Env) {
        [Environment] $local:NewEnv = $Env.Clone();
        foreach ($local:Entry in $this.GetEnumerator()) {
            if (script:IsIgnoredEnvironmentVariable $local:Entry.Key) {
                continue;
            }
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

    # Unapplies this EnvironmentDiff to the provided Environment, producing a new Environment
    [Environment] Unapply([Environment] $Env) {
        [Environment] $local:NewEnv = $Env.Clone();
        foreach ($local:Entry in $this.GetEnumerator()) {
            if (script:IsIgnoredEnvironmentVariable $local:Entry.Key) {
                continue;
            }
            $local:Key = $local:Entry.Key;
            $local:Value = $local:Entry.Value;
            if ($local:Value -is [PathDiff]) {
                $local:Value = $local:Value.Unapply($Env[$local:Key]);
            }
            if ($local:Value) {
                $local:NewEnv.Remove($local:Key);
            }
        }
        return $local:NewEnv;
    }

    hidden [bool] ValidateKeyValue([string] $Key, [psobject] $Value) {
        if (script:IsIgnoredEnvironmentVariable $Key) {
            throw [System.ArgumentException]::new("Invalid argument: Key");
            return $false;
        }
        if (($Key -ieq "Path") -and -not ($null -eq $Value -or $Value -is [PathDiff])) {
            throw [System.ArgumentException]::new("Invalid argument: Value");
            return $false;
        }
        if (($Key -ine "Path") -and -not ($null -eq $Value -or $Value -is [string])) {
            throw [System.ArgumentException]::new("Invalid argument: Value");
            return $false;
        }
        return $true;
    }
}

enum TargetArchitectureIn {
    x86 = 0;
    i686 = 0;

    amd64 = 1;
    x64 = 1;
    x86_64 = 1;

    arm = 2;

    arm64 = 3;
    aarch64 = 3;
}

enum TargetArchitecture {
    x86 = 0;
    amd64 = 1;
    arm = 2;
    arm64 = 3;
}

enum HostArchitectureIn {
    x86 = 0;

    amd64 = 1;
    x64 = 1;
    x86_64 = 1;
}

enum HostArchitecture {
    x86 = 0;
    amd64 = 1;
}

enum AppPlatform {
    Desktop = 0;
    UWP = 1;
}

class WindowsSdk {
    hidden static [WindowsSdk[]] $_None;
    hidden static [WindowsSdk[]] $_All;

    [version] $Version;
    [string] $WindowsSdkDir;
    [string] $WindowsLibPath;
    [string] $WindowsSdkBinPath;
    [AppPlatform[]] $SupportedPlatforms;

    WindowsSdk(
        [version] $Version,
        [string] $WindowsSdkDir,
        [string] $WindowsLibPath,
        [string] $WindowsSdkBinPath,
        [AppPlatform[]] $SupportedPlatforms
    ) {
        $this.Version = $Version;
        $this.WindowsSdkDir = $WindowsSdkDir;
        $this.WindowsLibPath = $WindowsLibPath;
        $this.WindowsSdkBinPath = $WindowsSdkBinPath;
        $this.SupportedPlatforms = [AppPlatform[]]($SupportedPlatforms);
        if ($null -eq $this.SupportedPlatforms) {
            $this.SupportedPlatforms = [AppPlatform[]](@());
        }
    }

    [string] GetVersionString() {
        if ($null -eq $this.Version) {
            return "none";
        }
        else {
            return $this.Version;
        }
    }

    [bool] IsNone() {
        return $null -eq $this.Version;
    }

    static [WindowsSdk] None() {
        if (-not [WindowsSdk]::_None) {
            [WindowsSdk]::_None = [WindowsSdk]::new($null, $null, $null, $null, [AppPlatform[]](@()));
        }
        return [WindowsSdk]::_None;
    }

    static [WindowsSdk[]] All() {
        if (-not [WindowsSdk]::_All) {
            [string[]]$local:RegRoots = @(
                "HKCU:\SOFTWARE\Wow6432Node",
                "HKLM:\SOFTWARE\Wow6432Node",
                "HKCU:\SOFTWARE",
                "HKLM:\SOFTWARE"
            );

            [Set] $local:Seen = [Set]::new();
            [WindowsSdk[]] $local:Sdks = @();

            foreach ($local:RegRoot in $local:RegRoots) {
                [psobject[]] $local:VersionRegKeys = Get-ChildItem "$RegRoot\Microsoft\Microsoft SDKs\Windows" -ErrorAction Ignore;
                foreach ($local:RegKey in $local:VersionRegKeys) {
                    [string] $local:VersionKey = $local:RegKey.PSChildName;
                    if ($local:VersionKey -notmatch "^v\d+\.\d+$") {
                        continue;
                    }

                    [string] $local:VersionString = $local:VersionKey.Substring(1)
                    [version] $local:Version = $local:VersionString;
                    [string] $local:WindowsSdkDir = $local:RegKey | Get-ItemPropertyValue -Name InstallationFolder;

                    # Special Case for Windows SDK <= 8.1
                    if ($local:Version.Major -le 8) {
                        if ($local:Seen.Contains($local:VersionString)) {
                            continue;
                        }

                        $local:Seen.Add($local:VersionString) | Out-Null;

                        [string] $local:WindowsLibPath = Join-Path -Path:$local:WindowsSdkDir -ChildPath:"References\CommonConfiguration\Neutral";
                        [string] $local:WindowsSdkBinPath = Join-Path -Path:$local:WindowsSdkDir -ChildPath:"bin";
                        [version] $local:WindowsSdkVersion = $local:Version;

                        $local:Sdks += [WindowsSdk]::new(
                            $local:WindowsSdkVersion,
                            $local:WindowsSdkDir,
                            $local:WindowsLibPath,
                            $local:WindowsSdkBinPath,
                            [AppPlatform[]](@([AppPlatform]::Desktop))
                        );
                        continue;
                    }

                    # Windows SDK >= 10

                    $local:WindowsSdkIncludeDir = Join-Path -Path:$local:WindowsSdkDir -ChildPath:"include";
                    $local:SdkVersions = Get-ChildItem $local:WindowsSdkIncludeDir | ForEach-Object {
                        if ($_.PSChildName -match "^\d+\.\d+\.\d+\.\d+$") {
                            $_.PSChildName -as [version];
                        }
                    };

                    foreach ($local:WindowsSdkVersion in $local:SdkVersions) {
                        if ($local:Seen.Contains($local:WindowsSdkVersion)) {
                            continue;
                        }

                        [AppPlatform[]] $local:SupportedPlatforms = @();
                        [string] $local:UWPTestPath = Join-Path -Path:$local:WindowsSdkDir -ChildPath:"include\$local:WindowsSdkVersion\um\Windows.h";
                        if (Test-Path -LiteralPath:$local:UWPTestPath) {
                            $local:SupportedPlatforms += @([AppPlatform]::UWP);
                        }

                        [string] $local:DestopTestPath = Join-Path -Path:$local:WindowsSdkDir -ChildPath:"include\$local:WindowsSdkVersion\um\winsdkver.h";
                        if (Test-Path -LiteralPath:$local:DestopTestPath) {
                            $local:SupportedPlatforms += @([AppPlatform]::Desktop);
                        }

                        if (-not $local:SupportedPlatforms) {
                            continue;
                        }

                        [string] $local:WindowsSdkBinPath = Join-Path -Path:$local:WindowsSdkDir -ChildPath:"bin";
                        [string] $local:WindowsSdkVersionedBinPath = Join-Path -Path:$local:WindowsSdkBinPath -ChildPath:$local:WindowsSdkVersion;
                        if (Test-Path -LiteralPath:$local:WindowsSdkVersionedBinPath) {
                            $local:WindowsSdkBinPath = $local:WindowsSdkVersionedBinPath;
                        }

                        [string] $local:WindowsSdkUnionMetadataPath = Join-Path -Path:$local:WindowsSdkDir -ChildPath:"UnionMetadata";
                        [string] $local:WindowsSdkVersionedUnionMetadataPath = Join-Path -Path:$local:WindowsSdkUnionMetadataPath -ChildPath:$local:WindowsSdkVersion;
                        if (Test-Path -LiteralPath:$local:WindowsSdkVersionedUnionMetadataPath) {
                            $local:WindowsSdkUnionMetadataPath = $local:WindowsSdkVersionedUnionMetadataPath;
                        }

                        [string] $local:WindowsSdkReferencesPath = Join-Path -Path:$local:WindowsSdkDir -ChildPath:"References";
                        [string] $local:WindowsSdkVersionedReferencesPath = Join-Path -Path:$local:WindowsSdkReferencesPath -ChildPath:$local:WindowsSdkVersion;
                        if (Test-Path -LiteralPath:$local:WindowsSdkVersionedReferencesPath) {
                            $local:WindowsSdkReferencesPath = $local:WindowsSdkVersionedReferencesPath;
                        }

                        [string] $local:WindowsSdkLibPath = "$local:WindowsSdkUnionMetadataPath;$local:WindowsSdkReferencesPath";
                        $local:Seen.Add($local:WindowsSdkVersion) | Out-Null;
                        $local:Sdks += [WindowsSdk]::new(
                            $local:WindowsSdkVersion,
                            $local:WindowsSdkDir,
                            $local:WindowsSdkLibPath,
                            $local:WindowsSdkBinPath,
                            [AppPlatform[]]($local:SupportedPlatforms)
                        );
                    }
                }
            }
            [WindowsSdk]::_All = $local:Sdks | Sort-Object -Descending -Property:{ $_.Version -as [version] };
        }

        return [WindowsSdk]::_All;
    }

    # Find a matching WindowsSdk
    static [WindowsSdk] Match([string]$SdkVersion, [AppPlatform]$AppPlatform) {
        if (-not $SdkVersion) { $SdkVersion = "latest"; }
        switch -RegEx ($SdkVersion.ToLower()) {
            '^none$' { return [WindowsSdk]::None(); }
            '^8\.1$' { return [WindowsSdk]::All() | Where-Object { $_.Version -like "8.1" } | Select-Object -First 1; }
            '^\d+\.\d+\.\d+\.\d+$' { return [WindowsSdk]::All() | Where-Object { $_.Version -eq $SdkVersion } | Select-Object -First 1; }
            '^\d+\.\d+(\.\d+)?$' { return [WindowsSdk]::All() | Where-Object { $_.Version -like "$SdkVersion.*" } | Select-Object -First 1; }
            '^latest$' { return [WindowsSdk]::All() | Select-Object -First 1; }
        }
        throw "Unsupported SDK Version format: $SdkVersion";
    }

    static [bool] Equals([WindowsSdk] $Left, [WindowsSdk] $Right) {
        if ([object]::ReferenceEquals($null, $Left)) {
            return [object]::ReferenceEquals($null, $Right);
        }
        if ([object]::ReferenceEquals($null, $Right)) {
            return $false;
        }
        if ($Left.Version -ne $Right.Version -or
            $Left.WindowsSdkDir -ne $Right.WindowsSdkDir -or
            $Left.WindowsLibPath -ne $Right.WindowsLibPath -or
            $Left.WindowsSdkBinPath -ne $Right.WindowsSdkBinPath) {
            return $false;
        }
        if ([object]::ReferenceEquals($null, $Left.SupportedPlatforms)) {
            return [object]::ReferenceEquals($null, $Right.SupportedPlatforms);
        }
        if ([object]::ReferenceEquals($null, $Right.SupportedPlatforms)) {
            return $false;
        }
        if ($Left.SupportedPlatforms.Length -ne $Right.SupportedPlatforms.Length) {
            return $false;
        }
        foreach ($local:LeftPlatform in $Left.SupportedPlatforms) {
            if ($local:LeftPlatform -notin $Right.SupportedPlatforms) {
                return $false;
            }
        }
        return $true;
    }

    static [bool] op_Equality([WindowsSdk] $Left, [WindowsSdk] $Right) {
        return [WindowsSdk]::Equals($Left, $Right);
    }

    static [bool] op_Inequality([WindowsSdk] $Left, [WindowsSdk] $Right) {
        return -not [WindowsSdk]::Equals($Left, $Right);
    }
}

class EnvironmentOptions {
    [TargetArchitecture] $Arch;
    [HostArchitecture] $HostArch;
    [AppPlatform] $AppPlatform;
    [string] $WindowsSdk;
    [bool] $NoExtensions;

    EnvironmentOptions() {
        $this._Init(
            [TargetArchitectureIn]::x86,
            [HostArchitectureIn]::x86,
            [AppPlatform]::Desktop,
            "latest",
            $false);
    }

    EnvironmentOptions(
        [TargetArchitectureIn] $Arch = [TargetArchitectureIn]::x86,
        [HostArchitectureIn] $HostArch = (if ($Arch -eq [TargetArchitectureIn]::amd64) { [HostArchitecture]::amd64 } else { [HostArchitecture]::x86 }),
        [AppPlatform] $AppPlatform = [AppPlatform]::Desktop,
        [string] $WindowsSdk = "latest",
        [bool] $NoExtensions = $false
    ) {
        $this._Init($Arch, $HostArch, $AppPlatform, $WindowsSdk, $NoExtensions);
    }

    hidden _Init(
        [TargetArchitectureIn] $Arch,
        [HostArchitectureIn] $HostArch,
        [AppPlatform] $AppPlatform,
        [string] $WindowsSdk,
        [bool] $NoExtensions
    ) {
        $this.Arch = [TargetArchitecture]$Arch;
        $this.HostArch = [HostArchitecture]$HostArch;
        $this.AppPlatform = $AppPlatform;
        $this.NoExtensions = $NoExtensions;

        $local:MatchingWindowsSdk = [WindowsSdk]::Match($WindowsSdk, $AppPlatform);
        if ($local:MatchingWindowsSdk) {
            $this.WindowsSdk = $local:MatchingWindowsSdk.GetVersionString();
        } elseif ($WindowsSdk -eq "latest") {
            $this.WindowsSdk = $null;
        } else {
            throw "Windows SDK $WindowsSdk could not be found";
        }
    }

    [bool] Equals([object]$Other) {
        return [EnvironmentOptions]::Equals($this, $Other);
    }

    [int] GetHashCode() {
        return $this.Arch.GetHashCode()
            -bxor $this.HostArch.GetHashCode()
            -bxor $this.AppPlatform.GetHashCode()
            -bxor [System.Collections.Generic.EqualityComparer[object]]::Default.GetHashCode($this.WindowsSdk)
            -bxor $this.NoExtensions.GetHashCode();
    }

    [string] ToString() {
        [string] $local:CommandArgs = "";
        if ([TargetArchitecture]::x86 -ne $this.Arch) { $local:CommandArgs += " -arch=$($this.Arch)"; }
        if ($this.Arch -ne $this.HostArch) { $local:CommandArgs += " -host_arch=$($this.HostArch)"; }
        if ([AppPlatform]::Desktop -ne $this.AppPlatform) { $local:CommandArgs += " -app_platform=$($this.AppPlatform)"; }
        if ($this.WindowsSdk) { $local:CommandArgs += " -winsdk=$($this.WindowsSdk)"; }
        if ($this.NoExtensions) { $local:CommandArgs += " -no_ext"; }
        return $local:CommandArgs;
    }

    static [bool] Equals([EnvironmentOptions] $Left, [EnvironmentOptions] $Right) {
        if ([object]::ReferenceEquals($null, $Left)) {
            return [object]::ReferenceEquals($null, $Right);
        }
        if ([object]::ReferenceEquals($null, $Right)) {
            return $false;
        }
        return $Left.Arch -eq $Right.Arch -and
            $Left.HostArch -eq $Right.HostArch -and
            $Left.AppPlatform -eq $Right.AppPlatform -and
            $Left.WindowsSdk -eq $Right.WindowsSdk -and
            $Left.NoExtensions -eq $Right.NoExtensions;
    }

    static [bool] op_Equality([EnvironmentOptions] $Left, [EnvironmentOptions] $Right) {
        return [EnvironmentOptions]::Equals($Left, $Right);
    }

    static [bool] op_Inequality([EnvironmentOptions] $Left, [EnvironmentOptions] $Right) {
        return -not [EnvironmentOptions]::Equals($Left, $Right);
    }

    # Create an EnvironmentOptions from a psobject (for deserialization purposes)
    static [EnvironmentOptions] FromObject([psobject] $Object) {
        if ($null -eq $Object) {
            return $null;
        }
        if ($Object -is [EnvironmentOptions]) {
            return $Object;
        }
        return [EnvironmentOptions]::new(
            $Object.Arch,
            $Object.HostArch,
            $Object.AppPlatform,
            $Object.WindowsSdk,
            $Object.NoExtensions
        );
    }

    # Create a psobject from an EnvironmentOptions (for serialization purposes)
    static [psobject] ToObject([EnvironmentOptions] $Object) {
        if ($null -eq $Object) {
            return $null;
        }
        return @{
            Arch = [string]($Object.Arch);
            HostArch = [string]($Object.HostArch);
            AppPlatform = [string]($Object.AppPlatform);
            WindowsSdk = $Object.WindowsSdk;
            NoExtensions = $Object.NoExtensions;
        };
    }
}

class EnvironmentDiffMap : System.Collections.DictionaryBase, System.Collections.IEnumerable {
    [bool] TryGetValue([EnvironmentOptions] $Key, [ref] $Value) {
        if ($this.InnerHashtable.Contains($Key)) {
            $Value.Value = $this.InnerHashtable[$Key];
            return $true;
        }
        $Value.Value = $null;
        return $false;
    }

    [void] Add([EnvironmentOptions] $Key, [EnvironmentDiff] $Value) {
        $this.InnerHashtable.Add($Key, $Value);
    }

    [bool] TryAdd([EnvironmentOptions] $Key, [EnvironmentDiff] $Value) {
        return $this.InnerHashtable.TryAdd($Key, $Value);
    }

    [bool] ContainsKey([EnvironmentOptions] $Key) {
        return $this.InnerHashtable.ContainsKey($Key);
    }

    [bool] ContainsValue([EnvironmentDiff] $Value) {
        return $this.InnerHashtable.ContainsValue($Value);
    }

    [bool] Remove([EnvironmentOptions] $Key) {
        return $this.InnerHashtable.Remove($Key);
    }

    hidden [System.Collections.Generic.ICollection[EnvironmentOptions]] get_Keys() {
        return [EnvironmentOptions[]](@($this.InnerHashtable.Keys));
    }

    hidden [System.Collections.Generic.ICollection[EnvironmentDiff]] get_Values() {
        return [EnvironmentOptions[]](@($this.InnerHashtable.Values));
    }

    hidden [EnvironmentDiff] get_Item([EnvironmentOptions] $Key) {
        [EnvironmentDiff] $Value = $null;
        [void]($this.TryGetValue($Key, [ref]$Value));
        return $Value;
    }

    hidden [void] set_Item([EnvironmentOptions] $Key, [EnvironmentDiff] $Value) {
        $this.InnerHashtable[$Key] = $Value;
    }

    hidden [object] OnGet([object] $Key, [object] $CurrentValue) {
        if (-not $Key -is [EnvironmentOptions]) {
            throw [System.ArgumentException]::new("Invalid key");
        }
        return $CurrentValue;
    }

    hidden [void] OnSet([object] $Key, [object] $OldValue, [object] $NewValue) {
        if (-not $Key -is [EnvironmentOptions]) {
            throw [System.ArgumentException]::new("Invalid key");
        }
        if (-not $NewValue -is [EnvironmentOptions]) {
            throw [System.ArgumentException]::new("Invalid value");
        }
    }

    hidden [void] OnInsert([object] $Key, [object] $Value) {
        if (-not $Key -is [EnvironmentOptions]) {
            throw [System.ArgumentException]::new("Invalid key");
        }
        if (-not $Value -is [EnvironmentOptions]) {
            throw [System.ArgumentException]::new("Invalid value");
        }
    }

    hidden [void] OnRemove([object] $Key, [object] $Value) {
        if (-not $Key -is [EnvironmentOptions]) {
            throw [System.ArgumentException]::new("Invalid key");
        }
    }

    hidden [void] OnValidate([object] $Key, [object] $Value) {
        if (-not $Key -is [EnvironmentOptions]) {
            throw [System.ArgumentException]::new("Invalid key");
        }
        if (-not $Value -is [EnvironmentOptions]) {
            throw [System.ArgumentException]::new("Invalid value");
        }
    }

    [System.Collections.Generic.IEnumerator[System.Collections.Generic.KeyValuePair[EnvironmentOptions, EnvironmentDiff]]] GetEnumerator() {
        $OfTypeMethodOpen = [System.Linq.Enumerable].GetMethod("OfType");
        $OfTypeMethodClosed = $OfTypeMethodOpen.MakeGenericMethod([System.Collections.DictionaryEntry]);
        $DictionaryEntries = [System.Collections.Generic.IEnumerable[System.Collections.DictionaryEntry]]$OfTypeMethodClosed.Invoke($null, @($this.InnerHashtable));
        $Projection = [System.Func[System.Collections.DictionaryEntry, System.Collections.Generic.KeyValuePair[EnvironmentOptions, EnvironmentDiff]]]{
            param($pair);
            return [System.Collections.Generic.KeyValuePair[EnvironmentOptions, EnvironmentDiff]]::new($pair.Key, $pair.Value);
        };
        $Selected = [System.Linq.Enumerable]::Select($DictionaryEntries, $Projection);
        return $Selected.GetEnumerator();
    }

    # Create an EnvironmentDiffMap from a psobject (for deserialization purposes)
    static [EnvironmentDiffMap] FromObject([psobject] $Object) {
        if ($null -eq $Object) {
            return $null;
        }
        if ($Object -is [EnvironmentDiffMap]) {
            return $Object
        }

        $local:Map = [EnvironmentDiffMap]::new();
        foreach ($local:Entry in $Object) {
            $local:Key = $local:Entry.Key;
            $local:Value = $local:Entry.Value;
            $local:Map.Add([EnvironmentOptions]::FromObject($local:Key), [EnvironmentDiff]::FromObject($local:Value));
        }

        return $local:Map;
    }

    # Create a psobject from an EnvironmentDiffMap (for serialization purposes)
    static [psobject] ToObject([EnvironmentDiffMap] $Object) {
        if ($null -eq $Object) {
            return $null;
        }

        [psobject[]] $Result = @();
        foreach ($local:Entry in $Object.GetEnumerator()) {
            $local:Key = $local:Entry.Key;
            $local:Value = $local:Entry.Value;
            $local:Result += @(@{
                Key = [EnvironmentOptions]::ToObject($local:Key);
                Value = [EnvironmentDiff]::ToObject($local:Value);
            });
        }
        return [System.Collections.Generic.List[object]]::new($Result);
    }
}

# Represents an instance of Visual Studio
class VisualStudioInstance {
    [string] $Name;
    [string] $Channel;
    [version] $Version;
    [string] $Path;
    hidden [EnvironmentDiffMap] $Envs;

    VisualStudioInstance([string] $Name, [string] $Channel, [version] $Version, [string] $Path, [EnvironmentDiffMap] $Envs) {
        $this.Name = $Name;
        $this.Channel = $Channel;
        $this.Version = $Version;
        $this.Path = $Path;
        $this.Envs = $Envs;
        if (-not $this.Envs) {
            $this.Envs = [EnvironmentDiffMap]::new();
        }
    }

    static [VisualStudioInstance] FromObject([psobject] $Object) {
        if ($null -eq $Object) {
            return $null;
        }
        if ($Object -is [VisualStudioInstance]) {
            return $Object;
        }
        return [VisualStudioInstance]::new(
            $Object.Name,
            $Object.Channel,
            $Object.Version -as [version],
            $Object.Path,
            [EnvironmentDiffMap]::FromObject($Object.Envs)
        );
    }

    static [psobject] ToObject([VisualStudioInstance] $Object) {
        if ($null -eq $Object) {
            return $null;
        }
        return @{
            Name = $Object.Name;
            Channel = $Object.Channel;
            Version = $Object.Version -as [string];
            Path = $Object.Path;
            Envs = [EnvironmentDiffMap]::ToObject($Object.Envs);
        };
    }

    hidden [string] GetPoshVsDevEnvironmentId([EnvironmentOptions] $Options) {
        return @{
            Name = $this.Name;
            Channel = $this.Channel;
            Version = $this.Version -as [string];
            Options = [EnvironmentOptions]::ToObject($Options);
        } | ConvertTo-Json;
    }

    [EnvironmentDiff] GetEnvironmentDiff(
        [TargetArchitectureIn]$Arch = [TargetArchitectureIn]::x86,
        [HostArchitecture]$HostArch = [HostArchitecture]::x86,
        [AppPlatform]$AppPlatform = [AppPlatform]::Desktop,
        [string]$WindowsSdk = "latest",
        [bool]$NoExtensions = $false
    ) {
        return $this.GetEnvironmentDiff([EnvironmentOptions]::new(
            $Arch,
            $HostArch,
            $AppPlatform,
            $WindowsSdk,
            $NoExtensions
        ));
    }

    [EnvironmentDiff] GetEnvironmentDiff([EnvironmentOptions] $Options) {
        [EnvironmentDiff] $local:EnvDiff = $null;
        if (-not $this.Envs) {
            $this.Envs = [EnvironmentDiffMap]::new();
        }
        if (-not $this.Envs.TryGetValue($Options, [ref]$local:EnvDiff)) {
            $local:CommandPath = Join-Path -Path:$this.Path -ChildPath:$script:VSDEVCMD_PATH;
            $local:Command = [scriptblock]::Create("
                & `"$local:CommandPath`" $($Options.ToString()) -no_logo | Out-Null;
                Get-ChildItem env:
            ");

            $local:Entries = script:ExecuteCommandInNewEnvironment $local:Command;
            $local:NewEnv = [Environment]::new();
            foreach ($local:Entry in $local:Entries) {
                if (script:IsIgnoredEnvironmentVariable $local:Entry.Key) {
                    continue;
                }
                $local:NewEnv[$local:Entry.Name] = $local:Entry.Value;
            }

            $local:CleanEnv = [Environment]::GetClean();
            $local:EnvDiff = [EnvironmentDiff]::DiffBetween($local:CleanEnv, $local:NewEnv);
            $this.Envs.Add($Options, $local:EnvDiff);
            $script:HasChanges = $true;
        }
        return $local:EnvDiff;
    }

    hidden [void] Apply(
        [TargetArchitectureIn]$Arch = [TargetArchitectureIn]::x86,
        [HostArchitecture]$HostArch = [HostArchitecture]::x86,
        [AppPlatform]$AppPlatform = [AppPlatform]::Desktop,
        [string]$WindowsSdk = "latest",
        [bool]$NoExtensions = $false
    ) {
        $this.Apply([EnvironmentOptions]::new(
            $Arch,
            $HostArch,
            $AppPlatform,
            $WindowsSdk,
            $NoExtensions
        ));
    }

    hidden [void] Apply([EnvironmentOptions] $Options) {
        $this.GetEnvironmentDiff($Options).
            Apply([Environment]::GetCurrent()).
            Apply($this.GetPoshVsDevEnvironmentId($Options));
    }

    hidden [void] Unapply([EnvironmentOptions] $Options) {
        $this.GetEnvironmentDiff($Options).
            Unapply([Environment]::GetCurrent()).
            Apply($null);
    }

    [void] Save() {
        $script:HasChanges = $true;
        script:SaveChanges;
    }
}

$script:OperatorPattern = "(?<Operator>[<>=^~]|<=|>=)?";
$script:RevisionPattern = "(?:\.(?<Revision>[*x]|\d+))?";
$script:BuildPattern = "(?:\.(?:(?<Build>[*x])|(?<Build>\d+)$script:RevisionPattern))?";
$script:MinorPattern = "(?:\.(?:(?<Minor>[*x])|(?<Minor>\d+)$script:BuildPattern))?";
$script:MajorPattern = "(?:(?<Major>[*x])|${script:OperatorPattern}v?(?<Major>\d+)$script:MinorPattern)";
$script:VersionSpecPattern = "^\s*(?:${script:MajorPattern})?\s*$";
$script:TransientEnvironmentVariables = @(
    "PoshVsDevVsName",
    "PoshVsDevVsChannel",
    "PoshVsDevVsVersion",
    "PoshVsDevArch",
    "PoshVsDevHostArch",
    "PoshVsDevPlatform",
    "PoshVsDevWindowsSdk",
    "PoshVsDevNoExtensions",
    "PoshVsDevClean"
);
$script:IgnoredEnvironmentVariables = $script:TransientEnvironmentVariables + @(
    "PoshVsDevEnvironment",
    "PoshVsDevDefaultEnvironment"
);

class VersionSpec {
    hidden static $VERSION_FRAGMENT_UNSPECIFIED = -1;
    hidden static $VERSION_FRAGMENT_STAR = -2;
    hidden static $VERSION_OPERATORS = @{
        ""   = [VersionComparisonOperator]::None;
        "="  = [VersionComparisonOperator]::Equal;
        "<"  = [VersionComparisonOperator]::LessThan;
        "<=" = [VersionComparisonOperator]::LessThanOrEqual;
        ">"  = [VersionComparisonOperator]::GreaterThan;
        ">=" = [VersionComparisonOperator]::GreaterThanOrEqual;
        "~"  = [VersionComparisonOperator]::Tilde;
        "^"  = [VersionComparisonOperator]::Caret;
    };

    static [bool] TryParse([string] $Text, [ref] $Value) {

        function TryParseLogicalOr([string] $Text, [ref] $Value) {
            $Text = $Text.Trim();
            [int] $local:End = $Text.IndexOf('||');
            if ($local:End -eq 0) {
                $Value.Value = $null;
                return $false;
            }

            if ($local:End -eq -1) {
                return TryParseRange $Text $Value;
            }

            $local:Left = $null;
            $local:Right = $null;
            if (-not (TryParseRange $Text.Substring(0, $End) ([ref]$local:Left)) -or
                -not (TryParseLogicalOr $Text.Substring($End + 2) ([ref]$local:Right))) {
                $Value.Value = $null;
                return $false;
            }


            $Value.Value = [VersionSpec]::Or($local:Left, $local:Right);
            return $true;
        }

        function TryParseRange([string] $Text, [ref] $Value) {
            $Text = $Text.Trim();
            [int] $local:End = $Text.IndexOf(' ');
            if ($local:End -eq -1) {
                return TryParsePrimitive $Text $Value;
            }

            $local:Left = $null;
            $local:Right = $null;
            if (-not (TryParsePrimitive $Text.Substring(0, $End) ([ref]$local:Left)) -or
                -not (TryParseRange $Text.Substring($End + 1) ([ref]$local:Right))) {
                $Value.Value = $null;
                return $false;
            }


            $Value.Value = [VersionSpec]::Range($local:Left, $local:Right);
            return $true;
        }

        function TryParsePrimitive([string] $Text, [ref] $Value) {
            [int] $local:Major = [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED;
            [int] $local:Minor = [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED;
            [int] $local:Build = [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED;
            [int] $local:Revision = [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED;
            if (-not ($Text -match $script:VersionSpecPattern) -or
                -not (TryParseXOrNumber $Matches.Major ([ref]$local:Major)) -or
                -not (TryParseXOrNumber $Matches.Minor ([ref]$local:Minor)) -or
                -not (TryParseXOrNumber $Matches.Build ([ref]$local:Build)) -or
                -not (TryParseXOrNumber $Matches.Revision ([ref]$local:Revision)) -or (
                    $Matches.Operator -and (
                        $local:Major -eq [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED -or
                        $local:Major -eq [VersionSpec]::VERSION_FRAGMENT_STAR -or
                        $local:Minor -eq [VersionSpec]::VERSION_FRAGMENT_STAR -or
                        $local:Build -eq [VersionSpec]::VERSION_FRAGMENT_STAR -or
                        $local:Revision -eq [VersionSpec]::VERSION_FRAGMENT_STAR))) {
                $Value.Value = $null;
                return $false;
            }
            [VersionComparisonOperator] $local:Operator = [VersionSpec]::VERSION_OPERATORS[$Matches.Operator -as [string]];
            $Value.Value = [VersionPrimitive]::new($local:Operator, $local:Major, $local:Minor, $local:Build, $local:Revision);
            return $true;
        }

        function TryParseXOrNumber([string] $Text, [ref] $Value) {
            if ($Text.Length -eq 0) {
                $Value.Value = [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED;
                return $true;
            }
            if ($Text.Length -eq 1 -and ($Text -eq '*' -or $Text -eq 'x' -or $Text -eq 'X')) {
                $Value.Value = [VersionSpec]::VERSION_FRAGMENT_STAR;
                return $true;
            }
            if ([int]::TryParse($Text, $Value) -and $Value.Value -ge 0) {
                return $true;
            }
            $Value.Value = 0;
            return $false;
        }

        if (TryParseLogicalOr $Text $Value) {
            $Value.Value = $Value.Value.Normalize();
            return $true;
        }
        return $false;
    }

    static [VersionRange] Range([VersionSpec] $Left, [VersionSpec] $Right) {
        return [VersionRange]::new([VersionRangeOperator]::Range, $Left, $Right);
    }

    static [VersionRange] Or([VersionSpec] $Left, [VersionSpec] $Right) {
        return [VersionRange]::new([VersionRangeOperator]::Or, $Left, $Right);
    }

    static [VersionPrimitive] Primitive([int] $Major, [int] $Minor, [int] $Build, [int] $Revision) {
        return [VersionPrimitive]::new([VersionComparisonOperator]::None, $Major, $Minor, $Build, $Revision);
    }

    static [VersionPrimitive] EQ([int] $Major, [int] $Minor, [int] $Build, [int] $Revision) {
        return [VersionPrimitive]::new([VersionComparisonOperator]::Equal, $Major, $Minor, $Build, $Revision);
    }

    static [VersionPrimitive] LT([int] $Major, [int] $Minor, [int] $Build, [int] $Revision) {
        return [VersionPrimitive]::new([VersionComparisonOperator]::LessThan, $Major, $Minor, $Build, $Revision);
    }

    static [VersionPrimitive] LE([int] $Major, [int] $Minor, [int] $Build, [int] $Revision) {
        return [VersionPrimitive]::new([VersionComparisonOperator]::LessThanOrEqual, $Major, $Minor, $Build, $Revision);
    }

    static [VersionPrimitive] GT([int] $Major, [int] $Minor, [int] $Build, [int] $Revision) {
        return [VersionPrimitive]::new([VersionComparisonOperator]::GreaterThan, $Major, $Minor, $Build, $Revision);
    }

    static [VersionPrimitive] GE([int] $Major, [int] $Minor, [int] $Build, [int] $Revision) {
        return [VersionPrimitive]::new([VersionComparisonOperator]::GreaterThanOrEqual, $Major, $Minor, $Build, $Revision);
    }

    [VersionSpec] Normalize() {
        return $this;
    }

    [bool] IsMatch([Version] $Version) {
        return $false;
    }
}

enum VersionRangeOperator {
    Range = 0;
    Or = 1;
}

class VersionRange : VersionSpec {
    [VersionRangeOperator] $Operator;
    [VersionSpec] $Left;
    [VersionSpec] $Right;

    VersionRange([VersionRangeOperator] $Operator, [VersionSpec] $Left, [VersionSpec] $Right) {
        $this.Operator = $Operator;
        $this.Left = $Left;
        $this.Right = $Right;
    }

    [VersionSpec] Normalize() {
        $local:Left = $this.Left.Normalize();
        $local:Right = $this.Right.Normalize();
        if ($this.Left -ne $local:Left -or $this.Right -ne $local:Right) {
            return [VersionRange]::new($this.Operator, $local:Left, $local:Right);
        }
        return $this;
    }

    [bool] IsMatch([Version] $Version) {
        if ($this.Operator -eq [VersionRangeOperator]::Range) {
            return $this.Left.IsMatch($Version) -and $this.Right.IsMatch($Version);
        }
        else {
            return $this.Left.IsMatch($Version) -or $this.Right.IsMatch($Version);
        }
    }

    [string] ToString() {
        if ($this.Operator -eq [VersionRangeOperator]::Range) {
            return "$($this.Left) $($this.Right)";
        }
        else {
            return "$($this.Left) || $($this.Right)";
        }
    }
}

enum VersionComparisonOperator {
    None = 0;
    Equal = 1;
    LessThan = 2;
    GreaterThan = 3;
    LessThanOrEqual = 4;
    GreaterThanOrEqual = 5;
    Tilde = 6;
    Caret = 7;
}

class VersionPrimitive : VersionSpec {
    [VersionComparisonOperator] $Operator = [VersionComparisonOperator]::GreaterThanOrEqual;
    [int] $Major = 0;
    [int] $Minor = 0;
    [int] $Build = 0;
    [int] $Revision = 0;

    VersionPrimitive([VersionComparisonOperator] $Operator, [int] $Major, [int] $Minor, [int] $Build, [int] $Revision) {
        $this.Operator = $Operator;
        $this.Major = $Major;
        $this.Minor = $Minor;
        $this.Build = $Build;
        $this.Revision = $Revision;
    }

    [VersionSpec] Normalize() {
        if ($this.Operator -eq [VersionComparisonOperator]::None) {
            if ($this.Major -eq [VersionSpec]::VERSION_FRAGMENT_STAR -or
                $this.Major -eq [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED) {
                return [VersionPrimitive]::GE(0, 0, 0, 0);
            }
            if ($this.Minor -eq [VersionSpec]::VERSION_FRAGMENT_STAR -or
                $this.Minor -eq [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED) {
                return [VersionSpec]::Range(
                    [VersionSpec]::GE($this.Major, 0, 0, 0),
                    [VersionSpec]::LT($this.Major + 1, 0, 0, 0)
                );
            }
            if ($this.Build -eq [VersionSpec]::VERSION_FRAGMENT_STAR -or
                $this.Build -eq [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED) {
                return [VersionSpec]::Range(
                    [VersionSpec]::GE($this.Major, $this.Minor, 0, 0),
                    [VersionSpec]::LT($this.Major, $this.Minor + 1, 0, 0)
                );
            }
            if ($this.Revision -eq [VersionSpec]::VERSION_FRAGMENT_STAR -or
                $this.Revision -eq [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED) {
                return [VersionSpec]::Range(
                    [VersionSpec]::GE($this.Major, $this.Minor, $this.Build, 0),
                    [VersionSpec]::LT($this.Major, $this.Minor, $this.Build + 1, 0)
                );
            }
            return [VersionSpec]::EQ($this.Major, $this.Minor, $this.Build, $this.Revision);
        }
        elseif ($this.Operator -eq [VersionComparisonOperator]::Tilde) {
            if ($this.Revision -ne [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED) {
                return [VersionSpec]::Range(
                    [VersionSpec]::GE($this.Major, $this.Minor, $this.Build, $this.Revision),
                    [VersionSpec]::LT($this.Major, $this.Minor, $this.Build + 1, 0)
                );
            }
            if ($this.Build -ne [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED) {
                return [VersionSpec]::Range(
                    [VersionSpec]::GE($this.Major, $this.Minor, $this.Build, 0),
                    [VersionSpec]::LT($this.Major, $this.Minor + 1, 0, 0)
                );
            }
            if ($this.Minor -ne [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED) {
                return [VersionSpec]::Range(
                    [VersionSpec]::GE($this.Major, $this.Minor, 0, 0),
                    [VersionSpec]::LT($this.Major + 1, 0, 0, 0)
                );
            }
        }
        elseif ($this.Operator -eq [VersionComparisonOperator]::Caret) {
            if ($this.Major -gt 0) {
                return [VersionSpec]::Range(
                    [VersionSpec]::GE($this.Major, [Math]::Max(0, $this.Minor), [Math]::Max($this.Build, 0), [Math]::Max($this.Revision, 0)),
                    [VersionSpec]::LT($this.Major + 1, 0, 0, 0)
                );
            }
            if ($this.Minor -gt 0) {
                return [VersionSpec]::Range(
                    [VersionSpec]::GE(0, $this.Minor, [Math]::Max($this.Build, 0), [Math]::Max($this.Revision, 0)),
                    [VersionSpec]::LT(0, $this.Minor + 1, 0, 0)
                );
            }
            if ($this.Minor -lt 0) {
                return [VersionSpec]::Range(
                    [VersionSpec]::GE(0, 0, 0, 0),
                    [VersionSpec]::LT(1, 0, 0, 0)
                );
            }
            if ($this.Build -gt 0) {
                return [VersionSpec]::Range(
                    [VersionSpec]::GE(0, 0, $this.Build, [Math]::Max($this.Revision, 0)),
                    [VersionSpec]::LT(0, 0, $this.Build + 1, 0)
                );
            }
            if ($this.Build -lt 0) {
                return [VersionSpec]::Range(
                    [VersionSpec]::GE(0, 0, 0, 0),
                    [VersionSpec]::LT(0, 1, 0, 0)
                );
            }
            if ($this.Revision -gt 0) {
                return [VersionSpec]::Range(
                    [VersionSpec]::GE(0, 0, 0, $this.Revision),
                    [VersionSpec]::LT(0, 0, 0, $this.Revision + 1)
                );
            }
            if ($this.Revision -lt 0) {
                return [VersionSpec]::Range(
                    [VersionSpec]::GE(0, 0, 0, 0),
                    [VersionSpec]::LT(0, 0, 1, 0)
                );
            }
        }
        return $this.Update(
            [Math]::Max(0, $this.Major),
            [Math]::Max(0, $this.Minor),
            [Math]::Max(0, $this.Build),
            [Math]::Max(0, $this.Revision)
        );
    }

    [VersionPrimitive] Update([int] $Major, [int] $Minor, [int] $Build, [int] $Revision) {
        if ($this.Major -ne $Major -or
            $this.Minor -ne $Minor -or
            $this.Build -ne $Build -or
            $this.Revision -ne $Revision) {
            return [VersionPrimitive]::new($this.Operator, $this.Major, $this.Minor, $this.Build, $this.Revision);
        }
        return $this;
    }

    [bool] IsMatch([Version] $Version) {
        [version] $local:NormalThis = [version]::new($this.Major, $this.Minor, $this.Build, $this.Revision);
        [version] $local:NormalVersion = [version]::new($Version.Major, $Version.Minor, [Math]::Max(0, $Version.Build), [Math]::Max(0, $Version.Revision));
        [int] $local:Result = $local:NormalVersion.CompareTo($local:NormalThis);
        switch ($this.Operator) {
            Equal {
                return $local:Result -eq 0;
            }
            LessThan {
                return $local:Result -lt 0;
            }
            LessThanOrEqual {
                return $local:Result -le 0;
            }
            GreaterThan {
                return $local:Result -gt 0;
            }
            GreaterThanOrEqual {
                return $local:Result -ge 0;
            }
        }
        throw "How did we get here! $($this.Operator)";
    }

    [string] ToString() {
        [string] $local:Text = switch ($this.Operator) {
            Equal { "=" }
            LessThan { "<" }
            LessThanOrEqual { "<=" }
            GreaterThan { ">" }
            GreaterThanOrEqual { ">=" }
            Tilde { "~" }
            Caret { "^"}
            default { "" }
        };
        if ($this.Major -eq [VersionSpec]::VERSION_FRAGMENT_STAR -or
            $this.Major -eq [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED) {
            return $local:Text + "*";
        }

        $local:Text += $this.Major -as [string];

        if ($this.Minor -eq [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED) {
            return $local:Text;
        }

        $local:Text += ".";

        if ($this.Minor -eq [VersionSpec]::VERSION_FRAGMENT_STAR) {
            return $local:Text + "*";
        }

        $local:Text += $this.Minor -as [string];

        if ($this.Build -eq [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED) {
            return $local:Text;
        }

        $local:Text += ".";

        if ($this.Build -eq [VersionSpec]::VERSION_FRAGMENT_STAR) {
            return $local:Text + "*";
        }

        $local:Text += $this.Build -as [string];

        if ($this.Revision -eq [VersionSpec]::VERSION_FRAGMENT_UNSPECIFIED) {
            return $local:Text;
        }

        $local:Text += ".";

        if ($this.Revision -eq [VersionSpec]::VERSION_FRAGMENT_STAR) {
            return $local:Text + "*";
        }

        $local:Text += $this.Revision -as [string];

        return $local:Text;
    }
}

function script:ExecuteCommandInNewEnvironment([scriptblock] $Command) {
    $local:TempOutputFile = New-TemporaryFile;
    $local:TempErrorFile = New-TemporaryFile;
    $local:EncodedCommand = script:EncodeCommand $Command;
    $local:ProcessArgs = @(
        "-NoLogo",
        "-NoProfile",
        "-OutputFormat", "XML",
        "-EncodedCommand", $local:EncodedCommand
    );
    Start-Process `
        -UseNewEnvironment `
        -NoNewWindow `
        -Wait `
        -FilePath "powershell" `
        -ArgumentList:$local:ProcessArgs `
        -RedirectStandardOutput:$local:TempOutputFile `
        -RedirectStandardError:$local:TempErrorFile `
        | Out-Null;
    $local:Content = Import-Clixml -LiteralPath:$local:TempOutputFile.FullName;
    Remove-Item $local:TempOutputFile -Force | Out-Null;
    Remove-Item $local:TempErrorFile -Force | Out-Null;
    $local:Content | ForEach-Object { $_ };
}

function script:EncodeCommand([scriptblock] $Command) {
    return [System.Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($Command.ToString()))
}

function script:IsIgnoredEnvironmentVariable([string] $Key) {
    return $Key -iin $script:IgnoredEnvironmentVariables;
}

function script:ClearTransientEnvironmentVariables() {
    foreach ($local:Key in $script:TransientEnvironmentVariables) {
        script:SetEnvironmentVariable $local:Key $null;
    }
}

# Converts a JSON object (from ConvertFrom-Json) into a Hashtable
function script:ConvertToHashTable([psobject] $Object) {
    if ($null -eq $Object) {
        return $null;
    }
    if ($Object -is [hashtable]) {
        return $Object;
    }
    $local:Table = @{};
    foreach ($local:Key in $Object | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name) {
        $local:Value = $Object | Select-Object -ExpandProperty $local:Key;
        $local:Table[$local:Key] = $local:Value;
    }
    return $local:Table;
}

# Sets or removes an environment variable
function script:SetEnvironmentVariable([string] $Key, [string] $Value) {
    if ($null -ne $Value) {
        [void](Set-Item -Force "ENV:\$Key" -Value $Value);
    }
    else {
        [void](Remove-Item -Force "ENV:\$Key");
    }
}

# Populates $script:VisualStudioVersions from cache if it is empty
function script:PopulateVisualStudioVersionsFromCache() {
    if ($null -eq $script:VisualStudioVersions) {
        if (Test-Path $script:CACHE_PATH) {
            $script:VisualStudioVersions = (Get-Content $script:CACHE_PATH | ConvertFrom-Json) `
                | ForEach-Object {
                    [VisualStudioInstance]::FromObject($_);
                };
        }
    }
}

# Gets the installed legacy visual studio instances from the registry
function script:GetLegacyVisualStudioInstancesFromRegistry() {
    Get-ChildItem -Path:"HKCU:\Software\Microsoft\VisualStudio\*.0" -PipelineVariable:ProductKey `
        | ForEach-Object -PipelineVariable:ConfigKey {
            Join-Path -Path:$local:ProductKey.PSParentPath -ChildPath:($local:ProductKey.PSChildName + "_Config") `
                | Get-Item -ErrorAction:SilentlyContinue;
        } `
        | ForEach-Object -PipelineVariable:ProfileKey {
            Join-Path -Path:$local:ProductKey.PSPath -ChildPath:"Profile" `
                | Get-Item -ErrorAction:SilentlyContinue
        } `
        | ForEach-Object {
            $local:Version = $local:ProfileKey | Get-ItemPropertyValue -Name BuildNum;
            $local:Path = $local:ConfigKey | Get-ItemPropertyValue -Name ShellFolder;
            $local:Name = "VisualStudio/$local:Version";
            if (Join-Path -Path:$local:Path -ChildPath:$script:VSDEVCMD_PATH | Test-Path) {
                [VisualStudioInstance]::new(
                    $local:Name,
                    "Release",
                    $local:Version -as [version],
                    $local:Path,
                    $null
                );
            }
        };
}

# Gets the installed visual studio instances from the VS instances directory
function script:GetVisualStudioInstancesFromVSInstancesDir() {
    Get-ChildItem $script:VS_INSTANCES_DIR `
        | ForEach-Object {
            $local:StatePath = Join-Path -Path:$_.FullName -ChildPath:"state.json";
            $local:State = Get-Content -Path:$local:StatePath | ConvertFrom-Json;
            $local:VsDevCmdPath = Join-Path -Path:$local:State.installationPath -ChildPath:$script:VSDEVCMD_PATH;
            if (Test-Path -LiteralPath:$local:VsDevCmdPath) {
                # other interesting data:
                # $local:State.installDate
                # $local:State.catalogInfo.buildBranch
                # $local:State.catalogInfo.productDisplayVersion
                # $local:State.catalogInfo.productSemanticVersion
                # $local:State.catalogInfo.productLineVersion
                # $local:State.catalogInfo.productMilestone
                # $local:State.catalogInfo.productMilestoneIsPreRelease
                # $local:State.catalogInfo.productName
                # $local:State.catalogInfo.productPatchVersion
                # $local:State.catalogInfo.productRelease
                # $local:State.catalogInfo.channelUri
                # $local:State.launchParams.fileName
                [VisualStudioInstance]::new(
                    $local:State.installationName,
                    $local:State.channelId,
                    $local:State.installationVersion -as [version],
                    $local:State.installationPath,
                    $null
                );
            }
        };
}

# Gets the installed visual studio instances
function script:GetVisualStudioInstances() {
    script:GetLegacyVisualStudioInstancesFromRegistry;
    script:GetVisualStudioInstancesFromVSInstancesDir;
}

# Populates $script:VisualStudioVersions from disk if it is empty
function script:PopulateVisualStudioVersions() {
    if ($null -eq $script:VisualStudioVersions) {
        $script:VisualStudioVersions = script:GetVisualStudioInstances `
            | Sort-Object -Property Version -Descending;

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
            | ConvertTo-Json -Depth 10;

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
    if (-not $ProfilePath) {
        return $false;
    }
    if (-not (Test-Path -LiteralPath:$ProfilePath)) {
        return $false;
    }
    return $true;
}

# Indicates whether "posh-vsdev" is referenced in the specified profile
function script:IsInProfile([string] $ProfilePath) {
    if (-not (script:HasProfile $ProfilePath)) {
        return $false;
    }
    $local:Content = Get-Content $ProfilePath -ErrorAction:SilentlyContinue;
    if ($local:Content -match "posh-vsdev") {
        return $true;
    }
    return $false;
}

# Indicates whether the Use-VisualStudioEnvironment cmdlet is referenced in the specified profile
function script:IsUsingEnvironment([string] $ProfilePath) {
    if (-not (script:HasProfile $ProfilePath)) {
        return $false;
    }
    $local:Content = Get-Content $ProfilePath -ErrorAction:SilentlyContinue;
    if ($local:Content -match "Use-VisualStudioEnvironment") {
        return $true;
    }
    return $false;
}

# Indicates whether the specified profile is signed
function script:IsProfileSigned([string] $ProfilePath) {
    if (-not (script:HasProfile $ProfilePath)) {
        return $false;
    }
    $local:Sig = Get-AuthenticodeSignature $ProfilePath;
    if (-not $local:Sig) {
        return $false;
    }
    if (-not $local:Sig.SignerCertificate) {
        return $false;
    }
    return $true;
}

# Indicates whether this module is installed within a PowerShell common module path
function script:IsInModulePaths() {
    foreach ($local:Path in $env:PSModulePath -split ";") {
        if (-not $local:Path.EndsWith("\")) {
            $local:Path += "\";
        }
        if ($PSScriptRoot.StartsWith($local:Path, [System.StringComparison]::InvariantCultureIgnoreCase)) {
            return $true;
        }
    }
    return $false;
}

function script:VisualStudioArgumentCompleter {
    param (
        $commandName,
        $parameterName,
        $wordToComplete,
        $commandAst,
        $fakeBoundParameters
    )

    $local:InstanceArgs = @{};
    $local:Name = $null;
    $local:LiteralName = $null;
    if ($fakeBoundParameters.ContainsKey("Name")) {
        $local:Name = $fakeBoundParameters["Name"];
        $local:InstanceArgs["Name"] = $local:Name;
    }
    elseif ($fakeBoundParameters.ContainsKey("LiteralName")) {
        $local:LiteralName = $fakeBoundParameters["LiteralName"];
        $local:InstanceArgs["LiteralName"] = $local:LiteralName;
    }

    $local:Channel = $null;
    $local:LiteralChannel = $null;
    if ($fakeBoundParameters.ContainsKey("Channel")) {
        $local:Channel = $fakeBoundParameters["Channel"];
        $local:InstanceArgs["Channel"] = $local:Channel;
    }
    elseif ($fakeBoundParameters.ContainsKey("LiteralChannel")) {
        $local:LiteralChannel = $fakeBoundParameters["LiteralChannel"];
        $local:InstanceArgs["LiteralChannel"] = $local:LiteralChannel;
    }

    $local:Version = $null;
    if ($fakeBoundParameters.ContainsKey("Version")) {
        $local:Version = $fakeBoundParameters["Version"];
        $local:InstanceArgs["Version"] = $local:Version;
    }

    $local:InstanceArgs.Remove($parameterName) | Out-Null;

    [string[]]$local:possibleValues = @();

    switch ($parameterName) {
        'Name' {
            $local:possibleValues = Get-VisualStudioInstance `
                @local:InstanceArgs `
                | ForEach-Object { $_.Name; };
        }
        'LiteralName' {
            $local:possibleValues = Get-VisualStudioInstance `
                @local:InstanceArgs `
                | ForEach-Object { $_.Name; };
        }
        'Channel' {
            $local:possibleValues = Get-VisualStudioInstance `
                @local:InstanceArgs `
                | ForEach-Object { $_.Channel; };
        }
        'LiteralChannel' {
            $local:possibleValues = Get-VisualStudioInstance `
                @local:InstanceArgs `
                | ForEach-Object { $_.Channel; };
        }
        'Version' {
            $local:possibleValues = Get-VisualStudioInstance `
                @local:InstanceArgs `
                | ForEach-Object { $_.Version; };
        }
    }

    $local:possibleValues | Where-Object {
        $_ -ilike "$wordToComplete*"
    };
}

function script:WindowsSdkArgumentCompleter {
    param (
        $commandName,
        $parameterName,
        $wordToComplete,
        $commandAst,
        $fakeBoundParameters
    )

    [bool]$local:HasAppPlatform = $false;
    [AppPlatform]$local:AppPlatform = [AppPlatform]::Desktop;
    if ($fakeBoundParameters.ContainsKey("AppPlatform")) {
        try {
            $local:AppPlatform = [AppPlatform]($fakeBoundParameters["AppPlatform"]);
            $local:HasAppPlatform = $true;
        }
        catch {
        }
    }

    [string[]]$local:possibleValues = @(
        "latest"
    );

    $local:possibleValues += [WindowsSdk]::All() `
        | Where-Object {
            (-not $local:HasAppPlatform) -or ($local:AppPlatform -in $_.SupportedPlatforms);
        } `
        | ForEach-Object {
            $_.Version -as [string];
        };

    $local:possibleValues | Where-Object { $_ -ilike "$wordToComplete*" };
}

<#
.SYNOPSIS
    Get installed Visual Studio instances.
.DESCRIPTION
    The Get-VisualStudioInstance cmdlet gets information about the installed Visual Studio instances on this machine.
.PARAMETER Name
    Specifies a name pattern that can be used to filter the results.
.PARAMETER LiteralName
    Specifies a literal name that can be used to filter the results.
.PARAMETER Channel
    Specifies a release channel pattern that can be used to filter the results.
.PARAMETER LiteralChannel
    Specifies a literal release channel that can be used to filter the results.
.PARAMETER Version
    Specifies a version number specification that can be used to filter the results.
.INPUTS
    None. You cannot pipe objects to Get-VisualStudioInstance.
.OUTPUTS
    VisualStudioInstance. Get-VisualStudioInstance returns a VisualStudioInstance object for each matching instance.
.EXAMPLE
    PS> Get-VisualStudioInstance
    Name                                              Channel                    Version      Path
    ----                                              -------                    -------      ----
    Microsoft Visual Studio 14.0                      Release                    14.0         C:\Program Files (x86)\Microsoft Visual Studio 14.0
.EXAMPLE
    PS> Get-VisualStudioInstance -Channel Release
    Name                                              Channel                    Version      Path
    ----                                              -------                    -------      ----
    Microsoft Visual Studio 14.0                      Release                    14.0         C:\Program Files (x86)\Microsoft Visual Studio 14.0
#>
function Get-VisualStudioInstance {
    [CmdletBinding(DefaultParameterSetName="NameChannel")]
    param (
        [Parameter(Position=0, ParameterSetName="NameChannel")]
        [Parameter(Position=0, ParameterSetName="NameLiteralChannel")]
        [SupportsWildcards()]
        [ArgumentCompleter({ script:VisualStudioArgumentCompleter @args })]
        [string] $Name,

        [Parameter(Position=0, ParameterSetName="LiteralNameChannel")]
        [Parameter(Position=0, ParameterSetName="LiteralNameLiteralChannel")]
        [ArgumentCompleter({ script:VisualStudioArgumentCompleter @args })]
        [string] $LiteralName,

        [Parameter(Position=1, ParameterSetName="NameChannel")]
        [Parameter(Position=1, ParameterSetName="LiteralNameChannel")]
        [SupportsWildcards()]
        [ArgumentCompleter({ script:VisualStudioArgumentCompleter @args })]
        [string] $Channel,

        [Parameter(Position=1, ParameterSetName="NameLiteralChannel")]
        [Parameter(Position=1, ParameterSetName="LiteralNameLiteralChannel")]
        [ArgumentCompleter({ script:VisualStudioArgumentCompleter @args })]
        [string] $LiteralChannel,

        [Parameter()]
        [ArgumentCompleter({ script:VisualStudioArgumentCompleter @args })]
        [string] $Version
    );

    script:PopulateVisualStudioVersionsFromCache;
    script:PopulateVisualStudioVersions;
    [VisualStudioInstance[]] $local:Versions = $script:VisualStudioVersions;
    if ($Name) {
        $local:Versions = $local:Versions | Where-Object -Property Name -ILike $Name;
    }
    if ($LiteralName) {
        $local:Versions = $local:Versions | Where-Object -Property Name -IEQ $LiteralName;
    }
    if ($Channel) {
        $local:Versions = $local:Versions | Where-Object -Property Channel -ILike $Channel;
    }
    if ($LiteralChannel) {
        $local:Versions = $local:Versions | Where-Object -Property Channel -IEQ $LiteralChannel;
    }
    if ($Version) {
        $local:VersionSpec = $null;
        if ([VersionSpec]::TryParse($Version, [ref]$local:VersionSpec)) {
            $local:Versions = $local:Versions | Where-Object { return $local:VersionSpec.IsMatch($_.Version); };
        }
        else {
            throw [System.FormatException]::new();
        }
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
.PARAMETER None
    Indicates no environment should be used (similar to calling Reset-VisualStudioEnvironment).
.PARAMETER Arch
    Indicates the target compilation architecture (default: `x86`).
.PARAMETER HostArch
    Indicates the host tools architecture (default: the same value as -Arch).
.PARAMETER AppPlatform
    Indicates the intended application platform (default: `Desktop`).
.PARAMETER WindowsSdk
    Indicates the Windows SDK to use for build tools (default: `latest`).
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
    [CmdletBinding(DefaultParameterSetName = "NameChannel")]
    param (
        [Parameter(Position=0, ParameterSetName = "NameChannel")]
        [Parameter(Position=0, ParameterSetName = "NameChannelOptions")]
        [Parameter(Position=0, ParameterSetName = "NameLiteralChannel")]
        [Parameter(Position=0, ParameterSetName = "NameLiteralChannelOptions")]
        [SupportsWildcards()]
        [ArgumentCompleter({ script:VisualStudioArgumentCompleter @args })]
        [string] $Name = $env:PoshVsDevVsName,

        [Parameter(Position=0, ParameterSetName = "LiteralNameChannel")]
        [Parameter(Position=0, ParameterSetName = "LiteralNameChannelOptions")]
        [Parameter(Position=0, ParameterSetName = "LiteralNameLiteralChannel")]
        [Parameter(Position=0, ParameterSetName = "LiteralNameLiteralChannelOptions")]
        [ArgumentCompleter({ script:VisualStudioArgumentCompleter @args })]
        [string] $LiteralName = $env:PoshVsDevVsName,

        [Parameter(Position=1, ParameterSetName = "NameChannel")]
        [Parameter(Position=1, ParameterSetName = "NameChannelOptions")]
        [Parameter(Position=1, ParameterSetName = "LiteralNameChannel")]
        [Parameter(Position=1, ParameterSetName = "LiteralNameChannelOptions")]
        [SupportsWildcards()]
        [ArgumentCompleter({ script:VisualStudioArgumentCompleter @args })]
        [string] $Channel = $env:PoshVsDevVsChannel,

        [Parameter(Position=1, ParameterSetName = "NameLiteralChannel")]
        [Parameter(Position=1, ParameterSetName = "NameLiteralChannelOptions")]
        [Parameter(Position=1, ParameterSetName = "LiteralNameLiteralChannel")]
        [Parameter(Position=1, ParameterSetName = "LiteralNameLiteralChannelOptions")]
        [SupportsWildcards()]
        [ArgumentCompleter({ script:VisualStudioArgumentCompleter @args })]
        [string] $LiteralChannel = $env:PoshVsDevVsChannel,

        [Parameter(Position=2, ParameterSetName = "NameChannel")]
        [Parameter(Position=2, ParameterSetName = "NameChannelOptions")]
        [Parameter(Position=2, ParameterSetName = "NameLiteralChannel")]
        [Parameter(Position=2, ParameterSetName = "NameLiteralChannelOptions")]
        [Parameter(Position=2, ParameterSetName = "LiteralNameChannel")]
        [Parameter(Position=2, ParameterSetName = "LiteralNameChannelOptions")]
        [Parameter(Position=2, ParameterSetName = "LiteralNameLiteralChannel")]
        [Parameter(Position=2, ParameterSetName = "LiteralNameLiteralChannelOptions")]
        [ArgumentCompleter({ script:VisualStudioArgumentCompleter @args })]
        [string] $Version = $env:PoshVsDevVsVersion,

        [Parameter(ParameterSetName = "Pipeline", Position = 0, ValueFromPipeline = $true, Mandatory = $true)]
        [Parameter(ParameterSetName = "PipelineOptions", Position = 0, ValueFromPipeline = $true, Mandatory = $true)]
        [psobject] $InputObject,

        [Parameter(ParameterSetName = "None")]
        [switch] $None = $($env:PoshVsDevClean -iin ("true","t","1")),

        [Parameter(ParameterSetName = "NameChannel")]
        [Parameter(ParameterSetName = "NameLiteralChannel")]
        [Parameter(ParameterSetName = "LiteralNameChannel")]
        [Parameter(ParameterSetName = "LiteralNameLiteralChannel")]
        [Parameter(ParameterSetName = "Pipeline")]
        [ArgumentCompletions('x86', 'i686', 'amd64', 'x64', 'x86_64', 'arm', 'arm64', 'aarch64')]
        [TargetArchitectureIn] $Arch = $(
            if ($env:PoshVsDevArch) {
                $env:PoshVsDevArch
            } else {
                [TargetArchitectureIn]::x86
            }
        ),

        [Parameter(ParameterSetName = "NameChannel")]
        [Parameter(ParameterSetName = "NameLiteralChannel")]
        [Parameter(ParameterSetName = "LiteralNameChannel")]
        [Parameter(ParameterSetName = "LiteralNameLiteralChannel")]
        [Parameter(ParameterSetName = "Pipeline")]
        [ArgumentCompletions('x86', 'amd64', 'x64', 'x86_64')]
        [HostArchitectureIn] $HostArch = $(
            if ($env:PoshVsDevHostArch) {
                $env:PoshVsDevHostArch
            } elseif ($Arch -eq [TargetArchitectureIn]::amd64) {
                [HostArchitectureIn]::amd64
            } else {
                [HostArchitectureIn]::x86
            }
        ),

        [Parameter(ParameterSetName = "NameChannel")]
        [Parameter(ParameterSetName = "NameLiteralChannel")]
        [Parameter(ParameterSetName = "LiteralNameChannel")]
        [Parameter(ParameterSetName = "LiteralNameLiteralChannel")]
        [Parameter(ParameterSetName = "Pipeline")]
        [ArgumentCompletions('Desktop', 'UWP')]
        [AppPlatform] $AppPlatform = $(
            if ($env:PoshVsDevAppPlatform) {
                $env:PoshVsDevAppPlatform
            } else {
                [AppPlatform]::Desktop
            }
        ),

        [Parameter(ParameterSetName = "NameChannel")]
        [Parameter(ParameterSetName = "NameLiteralChannel")]
        [Parameter(ParameterSetName = "LiteralNameChannel")]
        [Parameter(ParameterSetName = "LiteralNameLiteralChannel")]
        [Parameter(ParameterSetName = "Pipeline")]
        [ArgumentCompleter({ script:WindowsSdkArgumentCompleter @args })]
        [string] $WindowsSdk = $(
            if ($env:PoshVsDevWindowsSdk) {
                $env:PoshVsDevWindowsSdk
            } else {
                "latest"
            }
        ),

        [Parameter(ParameterSetName = "NameChannel")]
        [Parameter(ParameterSetName = "NameLiteralChannel")]
        [Parameter(ParameterSetName = "LiteralNameChannel")]
        [Parameter(ParameterSetName = "LiteralNameLiteralChannel")]
        [Parameter(ParameterSetName = "Pipeline")]
        [bool] $NoExtensions = $env:PoshVsDevNoExtensions -iin ("true","t","1"),

        [Parameter(ParameterSetName = "NameChannelOptions")]
        [Parameter(ParameterSetName = "NameLiteralChannelOptions")]
        [Parameter(ParameterSetName = "LiteralNameChannelOptions")]
        [Parameter(ParameterSetName = "LiteralNameLiteralChannelOptions")]
        [Parameter(ParameterSetName = "PipelineOptions")]
        [EnvironmentOptions] $Options,

        [switch] $Force
    );

    script:ClearTransientEnvironmentVariables;

    if ($None) {
        Reset-VisualStudioEnvironment;
        return;
    }

    if (-not $PSBoundParameters.ContainsKey("Options")) {
        $Options = [EnvironmentOptions]::new(
            $Arch,
            $HostArch,
            $AppPlatform,
            $WindowsSdk,
            $NoExtensions
        )
    }

    if (-not $Options) {
        $Options = [EnvironmentOptions]::new();
    }

    [VisualStudioInstance] $local:Instance = $null;
    if ($InputObject) {
        $local:Instance = [VisualStudioInstance]::FromObject($InputObject);
    } else {
        $local:InstanceArgs = @{};
        if ($Name) { $local:InstanceArgs["Name"] = $Name; }
        elseif ($LiteralName) { $local:InstanceArgs["LiteralName"] = $LiteralName; }

        if ($Channel) { $local:InstanceArgs["Channel"] = $Channel; }
        elseif ($LiteralChannel) { $local:InstanceArgs["LiteralChannel"] = $LiteralChannel; }

        $local:Instance = Get-VisualStudioInstance `
            @local:InstanceArgs `
            -Version:$Version `
            | Select-Object -First:1;
    }

    if ($local:Instance) {
        $local:PoshVsDevEnvironment = $local:Instance.GetPoshVsDevEnvironmentId($Options);
        if ($Force -or 
            ($env:PoshVsDevEnvironment -ine $local:PoshVsDevEnvironment) -or 
            ($script:VisualStudioVersion -ne $local:Instance) -or
            ($script:CurrentEnvironmentOptions -ne $Options)) {
            if ($script:VisualStudioVersion -and $script:CurrentEnvironmentOptions) {
                $script:VisualStudioVersion.Unapply($script:CurrentEnvironmentOptions);
            }
            $local:Instance.Apply($Options);
            script:SaveChanges;
            Write-Host "Using Development Environment from '$($local:Instance.Name)' [$($Options.ToString().Trim())]." -ForegroundColor:DarkGray;
            $script:VisualStudioVersion = $local:Instance;
            $script:CurrentEnvironmentOptions = $Options;
        }
    }
    else {
        [string] $local:Message = "Could not find Visual Studio";
        [string[]] $local:MessageParts = @();
        if ($Name) {
            $local:MessageParts += "Name='$Name'";
        }
        elseif ($LiteralName) {
            $local:MessageParts += "LiteralName='$LiteralName'";
        }
        if ($Channel) {
            $local:MessageParts += "Channel='$Channel'";
        }
        elseif ($LiteralChannel) {
            $local:MessageParts += "LiteralChannel='$LiteralChannel'";
        }
        if ($Version) {
            $local:MessageParts += "Version='$Version'";
        }
        if ($local:MessageParts.Length -gt 0) {
            $local:Message += "for " + $local:MessageParts[0];
            if ($local:MessageParts.Length -eq 2) {
                $local:Message += " and " + $local:MessageParts[1];
            }
            elseif ($local:MessageParts.Length -gt 2) {
                for ($local:I = 1; $local:I -lt $local:MessageParts.Length - 1; $local:I++) {
                    $local:Message += ", " + $local:MessageParts[$local:I];
                }
                if ($local:MessageParts.Length -gt 2) {
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
function Reset-VisualStudioEnvironment {
    [CmdletBinding()]
    param (
        [switch] $Force
    );

    script:ClearTransientEnvironmentVariables;

    if ($Force -or $script:VisualStudioVersion -or $env:PoshVsDevEnvironment) {
        if ($script:VisualStudioVersion -and $script:CurrentEnvironmentOptions) {
            $script:VisualStudioVersion.Unapply($script:CurrentEnvironmentOptions);
        }
        else {
            [Environment]::GetDefault().Apply($null);
        }
        $script:VisualStudioVersion = $null;
        $script:CurrentEnvironmentOptions = $null;
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
    None. You cannot pipe objects to Reset-VisualStudioInstanceCache.
.OUTPUTS
    None.
#>
function Reset-VisualStudioInstanceCache {
    [CmdletBinding()]
    param (
    );

    $script:VisualStudioVersions = $null;
    if (Test-Path $script:CACHE_PATH) {
        Remove-Item $script:CACHE_PATH -Force | Out-Null;
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
    None. You cannot pipe objects to Reset-VisualStudioInstanceCache.
.OUTPUTS
    None.
#>
function Add-VisualStudioEnvironmentToProfile {
    [CmdletBinding()]
    param (
        [switch] $AllHosts,
        [switch] $UseEnvironment,
        [switch] $Force
    );

    [string] $local:ProfilePath = (
        if ($AllHosts) {
            $profile.CurrentUserAllHosts
        } else {
            $profile.CurrentUserCurrentHost
        }
    );
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

<#
.SYNOPSIS
    Get matching Windows SDKs installed on the machine.
.DESCRIPTION
    Gets matching Windows SDKs installed on the machine.
.PARAMETER SdkVersion
    The Windows SDK version to match.
.INPUTS
    None. You cannot pipe objects to Get-WindowsSdk
.OUTPUTS
    WindowsSdk. Get-WindowsSdk returns one or more WindowsSdk objects matching the provided inputs.
#>
function Get-WindowsSdk {
    [CmdletBinding(DefaultParameterSetName="SdkVersion")]
    param (
        [Parameter(Position=0, ParameterSetName="SdkVersion")]
        [SupportsWildcards()]
        [ArgumentCompleter({ script:WindowsSdkArgumentCompleter @args })]
        [string] $SdkVersion,

        [Parameter(Position=0, ParameterSetName="LiteralSdkVersion")]
        [ArgumentCompleter({ script:WindowsSdkArgumentCompleter @args })]
        [string] $LiteralSdkVersion,

        [ArgumentCompletions("Desktop", "UWP")]
        [AppPlatform] $AppPlatform
    )

    if ($LiteralSdkVersion -ieq "latest" -or $SdkVersion -ieq "latest") {
        return [WindowsSdk]::All() | Select-Object -First 1;
    }

    [WindowsSdk]::All() |
        Where-Object {
            ((-not $SdkVersion) -or ($_.Version -ilike $SdkVersion)) -and `
            ((-not $LiteralSdkVersion) -or ($_.Version -ieq $LiteralSdkVersion)) -and `
            ((-not $AppPlatform) -or ($AppPlatform -in $_.SupportedPlatforms));
        } |
        ForEach-Object {
            [WindowsSdk]::new(
                $_.Version,
                $_.WindowsSdkDir,
                $_.WindowsLibPath,
                $_.WindowsSdkBinPath,
                $_.SupportedPlatforms
            );
        };
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
[EnvironmentOptions] $script:CurrentEnvironmentOptions = $null; # Current Environment Options

if ($env:PoshVsDevEnvironment) {
    $local:PoshVsDevEnvironmentObject = $env:PoshVsDevEnvironment | ConvertFrom-Json -ErrorAction:SilentlyContinue;
    if ($local:PoshVsDevEnvironmentObject) {
        $script:VisualStudioVersion = Get-VisualStudioInstance `
            -LiteralName:$local:PoshVsDevEnvironmentObject.Name `
            -LiteralChannel:$local:PoshVsDevEnvironmentObject.Channel `
            -Version:$local:PoshVsDevEnvironmentObject.Version;
        $script:CurrentEnvironmentOptions = [EnvironmentOptions]::FromObject($local:PoshVsDevEnvironmentObject.Options);
    }
}

[Environment]::GetDefault() | Out-Null;
