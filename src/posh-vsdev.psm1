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
    [version] $Version;
    [string] $Path;
    hidden [EnvironmentDiff] $Env;

    VisualStudioInstance([string] $Name, [string] $Channel, [version] $Version, [string] $Path, [EnvironmentDiff] $Env) {
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
            $Object.Version -as [version],
            $Object.Path,
            [EnvironmentDiff]::FromObject($Object.Env)
        );
    }

    static [psobject] ToObject([VisualStudioInstance] $Object) {
        if ($Object -eq $null) { return $null; }
        return @{
            Name = $Object.Name;
            Channel = $Object.Channel;
            Version = $Object.Version -as [string];
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

$script:OperatorPattern = "(?<Operator>[<>=^~]|<=|>=)?";
$script:RevisionPattern = "(?:\.(?<Revision>[*x]|\d+))?";
$script:BuildPattern = "(?:\.(?:(?<Build>[*x])|(?<Build>\d+)$script:RevisionPattern))?";
$script:MinorPattern = "(?:\.(?:(?<Minor>[*x])|(?<Minor>\d+)$script:BuildPattern))?";
$script:MajorPattern = "(?:(?<Major>[*x])|${script:OperatorPattern}v?(?<Major>\d+)$script:MinorPattern)";
$script:VersionSpecPattern = "^\s*(?:${script:MajorPattern})?\s*$";

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

    static [VersionRange] Range([VersionSpec] $Left, [VersionSpec] $Right) { return [VersionRange]::new([VersionRangeOperator]::Range, $Left, $Right); }
    static [VersionRange] Or([VersionSpec] $Left, [VersionSpec] $Right) { return [VersionRange]::new([VersionRangeOperator]::Or, $Left, $Right); }
    static [VersionPrimitive] Primitive([int] $Major, [int] $Minor, [int] $Build, [int] $Revision) { return [VersionPrimitive]::new([VersionComparisonOperator]::None, $Major, $Minor, $Build, $Revision); }
    static [VersionPrimitive] EQ([int] $Major, [int] $Minor, [int] $Build, [int] $Revision) { return [VersionPrimitive]::new([VersionComparisonOperator]::Equal, $Major, $Minor, $Build, $Revision); }
    static [VersionPrimitive] LT([int] $Major, [int] $Minor, [int] $Build, [int] $Revision) { return [VersionPrimitive]::new([VersionComparisonOperator]::LessThan, $Major, $Minor, $Build, $Revision); }
    static [VersionPrimitive] LE([int] $Major, [int] $Minor, [int] $Build, [int] $Revision) { return [VersionPrimitive]::new([VersionComparisonOperator]::LessThanOrEqual, $Major, $Minor, $Build, $Revision); }
    static [VersionPrimitive] GT([int] $Major, [int] $Minor, [int] $Build, [int] $Revision) { return [VersionPrimitive]::new([VersionComparisonOperator]::GreaterThan, $Major, $Minor, $Build, $Revision); }
    static [VersionPrimitive] GE([int] $Major, [int] $Minor, [int] $Build, [int] $Revision) { return [VersionPrimitive]::new([VersionComparisonOperator]::GreaterThanOrEqual, $Major, $Minor, $Build, $Revision); }
    [VersionSpec] Normalize() { return $this; }
    [bool] IsMatch([Version] $Version) { return $false; }
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
            Equal { return $local:Result -eq 0; }
            LessThan { return $local:Result -lt 0; }
            LessThanOrEqual { return $local:Result -le 0; }
            GreaterThan { return $local:Result -gt 0; }
            GreaterThanOrEqual { return $local:Result -ge 0; }
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

# Gets the installed legacy visual studio instances from the registry
function script:GetLegacyVisualStudioInstancesFromRegistry() {
    Get-ChildItem HKCU:\Software\Microsoft\VisualStudio\*.0 -PipelineVariable:ProductKey `
        | ForEach-Object { Join-Path $local:ProductKey.PSParentPath ($local:ProductKey.PSChildName + "_Config") | Get-Item -ErrorAction:SilentlyContinue; } -PipelineVariable:ConfigKey `
        | ForEach-Object { Join-Path $local:ProductKey.PSPath Profile | Get-Item -ErrorAction:SilentlyContinue } -PipelineVariable:ProfileKey `
        | ForEach-Object {
            $local:Version = $local:ProfileKey | Get-ItemPropertyValue -Name BuildNum;
            $local:Path = $local:ConfigKey | Get-ItemPropertyValue -Name ShellFolder;
            $local:Name = "VisualStudio/$local:Version";
            if (Join-Path $local:Path $script:VSDEVCMD_PATH | Test-Path) {
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
            $local:StatePath = Join-Path $_.FullName "state.json";
            $local:State = Get-Content $local:StatePath | ConvertFrom-Json;
            if (Join-Path $local:State.installationPath $script:VSDEVCMD_PATH | Test-Path) {
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
    if ($script:VisualStudioVersions -eq $null) {
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
    The Get-VisualStudioInstance cmdlet gets information about the installed Visual Studio instances on this machine.
.PARAMETER Name
    Specifies a name that can be used to filter the results.
.PARAMETER Channel
    Specifies a release channel that can be used to filter the results.
.PARAMETER Version
    Specifies a version number that can be used to filter the results.
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
function Get-VisualStudioInstance([string] $Name, [string] $Channel, [string] $Version) {
    script:PopulateVisualStudioVersionsFromCache;
    script:PopulateVisualStudioVersions;
    [VisualStudioInstance[]] $local:Versions = $script:VisualStudioVersions;
    if ($Name) {
        $local:Versions = $local:Versions | Where-Object -Property Name -Like $Name;
    }
    if ($Channel) {
        $local:Versions = $local:Versions | Where-Object -Property Channel -Like $Channel;
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
    [CmdletBinding(DefaultParameterSetName = "Match")]
    param (
        [Parameter(ParameterSetName = "Match")]
        [string] $Name,
        [Parameter(ParameterSetName = "Match")]
        [string] $Channel,
        [Parameter(ParameterSetName = "Match")]
        [string] $Version,
        [Parameter(ParameterSetName = "None")]
        [switch] $None,
        [Parameter(ParameterSetName = "Pipeline", Position = 0, ValueFromPipeline = $true, Mandatory = $true)]
        [psobject] $InputObject
    );

    [void]([Environment]::GetDefault());
    if ($None) {
        Reset-VisualStudioEnvironment;
        return;
    }

    [VisualStudioInstance] $local:Instance = $null;
    if ($InputObject) {
        $local:Instance = [VisualStudioInstance]::FromObject($InputObject);
    } else {
        $local:Instance = Get-VisualStudioInstance -Name:$Name -Channel:$Channel -Version:$Version | Select-Object -First:1;
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
    None. You cannot pipe objects to Reset-VisualStudioInstanceCache.
.OUTPUTS
    None.
#>
function Reset-VisualStudioInstanceCache() {
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
    None. You cannot pipe objects to Reset-VisualStudioInstanceCache.
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
        'Add-VisualStudioEnvironmentToProfile'
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