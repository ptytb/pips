# License: MIT
# Copyright (c) 2019 Ilya Pronin. All Rights Reserved.


class Irvine {
    
    static hidden [string] $PluginName = 'Irvine'
    static hidden [string] $IndexUrl = 'https://www.lfd.uci.edu/~gohlke/pythonlibs/'
    static hidden [string] $Description = @"
Provides search, install and local caching for packages published at Christoph Gohlke's page: https://www.lfd.uci.edu/~gohlke/pythonlibs/, Laboratory for Fluorescence Dynamics, University of California, Irvine.
"@
    static hidden [string] $PackageType = 'wheel_irvine'
    static hidden [string] $DatabaseFileName = 'packages.bin'
    
    hidden [hashtable] $Database
  
    hidden [string] $PluginConfigurationDirectory
    hidden [System.Func[string, bool]] $CreatePath
    hidden [System.Func[string, string, hashtable]] $DownloadArchive
    hidden [System.Func[string, bool]] $WriteLog
    hidden [System.Func[string, bool]] $FileExists
    hidden [System.Func[object, string, bool]] $Serialize
    hidden [System.Func[string, object]] $Deserialize     
    hidden [System.Func[string, hashtable]] $GetPackageInfoFromWheelName
    hidden [System.Func[string, string, System.Func[hashtable, bool]]] $TestCanInstallPackageTo

    Irvine() {
    }

    [void] Init($PluginConfigurationDirectory, $DownloadArchive, $WriteLog, $FileExists, $Serialize, $Deserialize,
                $GetPackageInfoFromWheelName, $TestCanInstallPackageTo) {
                    
        $this.PluginConfigurationDirectory = $PluginConfigurationDirectory
        $this.DownloadArchive = $DownloadArchive         
        $this.WriteLog = $WriteLog
        $this.FileExists = $FileExists
        $this.Serialize = $Serialize
        $this.Deserialize = $Deserialize
        $this.GetPackageInfoFromWheelName = $GetPackageInfoFromWheelName
        $this.TestCanInstallPackageTo = $TestCanInstallPackageTo
        
        $this.LoadDatabase()
        
        $null = New-Item -Type Directory -ErrorAction Ignore -Path $PluginConfigurationDirectory
        $null = New-Item -Type Directory -ErrorAction Ignore -Path $this.GetCachePath()         
    }
    
    [void] Release() {         
    }
    
    [string] GetCachePath() {
        return "$($this.PluginConfigurationDirectory)\cache\archive"
    }
    
    hidden [string] GetLocallyCreatedDatabasePath() {
        return "$($this.PluginConfigurationDirectory)\$([Irvine]::DatabaseFileName)"
    }
    
    hidden [string] GetDatabasePath() {
        $LocallyCreatedDatabasePath = $this.GetLocallyCreatedDatabasePath()
        $DefaultDatabasePath = "$PSScriptRoot\$([Irvine]::DatabaseFileName)"
        if ($this.FileExists.Invoke($LocallyCreatedDatabasePath)) {
            return $LocallyCreatedDatabasePath
        } elseif ($this.FileExists.Invoke($DefaultDatabasePath)) {
            return $DefaultDatabasePath
        } else {
            return $null
        }
    }
    
    hidden [string] GetArchiveFilenameFromLocalCache([string] $archiveName) {
        return "$($this.GetCachePath())\$archiveName"         
    }
    
    hidden [string] GetArchiveLocation([string] $name, [string] $version, [string] $type,
                                       [string] $pythonVersion, [string] $pythonArch) {
        
        $FuncCanInstall = $this.TestCanInstallPackageTo.Invoke($pythonVersion, $pythonArch)
        $archive = $null
        
        $candidates = $this.GetPackageVersions($name, $FuncCanInstall)
        
        if ([string]::IsNullOrEmpty($version)) {
            $archive = $candidates | Select-Object -First 1
            if ($archive) {
                $this.WriteLog.Invoke("Version is not specified, guessed the latest $($archive.info.Version) of $(($candidates | ForEach-Object { $_.info.Version }) -join ', ')")
            }
        } else {
            foreach ($candidate in $candidates) {
                if ($version -ne $candidate.info.Version) {
                    continue
                }
                
                $archive = $candidate
                break
            }
        }
        
        if (-not $archive) {
            $this.WriteLog.Invoke("Can't find a package '$name'$(if ($version) { " version $version" } else {''}) for $pythonArch-bit Python $pythonVersion at $([Irvine]::IndexUrl)")
            return $null
        }
        
        $location = $this.GetArchiveFilenameFromLocalCache($archive.filename)
        
        if (-not $this.FileExists.Invoke($location)) {
            $null = $this.WriteLog.Invoke("Downloading $($archive.url)")
            $result = $this.DownloadArchive.Invoke($archive.url, $this.GetCachePath())
            $this.WriteLog.Invoke($result.output)
        } else {
            $null = $this.WriteLog.Invoke("Found in cache: $($archive.url)")
        }
        
        if ($this.FileExists.Invoke($location)) {
            return $location
        } else {
            $this.WriteLog.Invoke("Failed to download archive $name from $($archive.url)")
            return $null
        }         
    }
    
    [string] GetPluginName() {
        return [Irvine]::PluginName
    }
    
    [string] GetDescription() {
        return [Irvine]::Description
    }
    
    [array] GetSupportedPackageTypes() {
        return @( [Irvine]::PackageType )
    }
    
    [System.Collections.ArrayList] GetAllPackageNames() {
        return $this.Database.items.Keys
    }
    
    [object[]] GetPackageVersions([string] $name, $FuncCanInstall) {
        return ($this.Database.items."$name".wheels |
            ForEach-Object {
                $info = $this.GetPackageInfoFromWheelName.Invoke($_.filename)
                $_ | Add-Member -MemberType NoteProperty -Name info -Value $info
                $_
            } |
            Where-Object { $FuncCanInstall.Invoke($_.info) } |
            Sort-Object -Property @{ Expression={  [version] ($_.info.VersionCanonical) }; Descending = $true })
    }     
    
    [System.Collections.ArrayList] GetSearchResults([string] $query, [string] $pythonVersion, [string] $pythonArch) {
        $results = [System.Collections.ArrayList]::new()
        $FuncCanInstall = $this.TestCanInstallPackageTo.Invoke($pythonVersion, $pythonArch)
        
        foreach ($package in $this.Database.items.GetEnumerator()) {
            
            ($name, $info) = ($package.Key, $package.Value)             
            
            if (($name -inotmatch $query) -and ($info.description -inotmatch $query)) {
                continue
            }
            
            foreach ($wheel in $info.wheels) {                 
                
                $wheelInfo = $this.GetPackageInfoFromWheelName.Invoke($wheel.filename)
                
                # For debugging
                $meta = "py=$($wheelInfo.Python) abi=$($wheelInfo.ABI) plat=$($wheelInfo.Platform) file=$($wheel.filename) |"
                
                if (-not $FuncCanInstall.Invoke($wheelInfo)) {
                    continue
                }                 
                
                [void] $results.Add(@{
                    'Name'=$name;
                    'Version'=$wheelInfo.version;
                    'Description'="$meta $($info.description)";
                    'Type'=[Irvine]::PackageType;
                    })                 
                    
            }             
        }
        
        return $results
    }
    
    [hashtable] PackageActionHook([string] $name, [string] $version, [string] $type,
                               [string] $pythonVersion, [string] $pythonArch, [ref] $hookError) {
        if ($type -ne [Irvine]::PackageType) {
            return $null
        }
        
        $location = $this.GetArchiveLocation($name, $version, $type, $pythonVersion, $pythonArch)
        
        if (-not [string]::IsNullOrEmpty($location)) {
            return @{
                'Name'=$location;
                'Version'=$version;
                'Type'='wheel';
            }
        } else {
            $hookError.Value = $true
            return $null
        }
    }
    
    [string] GetPackageHomepage([string] $name, [string] $type) {
        if ($type -ne [Irvine]::PackageType) {
            return $null
        } elseif ($this.Database.items.Contains($name)) {
            return $this.Database.items."$name".url
        } else {
            return $null
        }
    }
    
    [array] GetToolMenuCommands() {
        $self = $this
        
        return @(
            @{
                Persistent=$true;
                MenuText = "$([Irvine]::PluginName): update package index";
                Code = {
                    $self.UpdateDatabase()
                }.GetNewClosure();
            };
        )         
    }
        
    hidden [void] LoadDatabase() {
        $path = $this.GetDatabasePath()
        if ($path) {
            $this.Database = $this.Deserialize.Invoke($path)
            [void] $this.WriteLog.Invoke(@"
Loaded package index: $($this.Database.items.Count) packages
Index updated: $($this.Database.originTimestamp)
"@)
        } else {             
            [void] $this.WriteLog.Invoke('Database not found.')
        }
    }
    
    hidden [void] SaveDatabase() {
        if ($this.Database) {
            $path = $this.GetLocallyCreatedDatabasePath()
            [void] $this.Serialize.Invoke($this.Database, $path)
        }
    }
    
    hidden [void] UpdateDatabase() {
        $this.SaveDatabase()
    }     
    
}

Function NewPluginInstance() {
    return [Irvine]::new()
}

Export-ModuleMember -Function NewPluginInstance
