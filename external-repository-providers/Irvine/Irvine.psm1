# License: MIT
# Copyright (c) 2019 Ilya Pronin. All Rights Reserved.


class Irvine {

    static hidden [string] $PluginName = 'Irvine'
    static hidden [string] $IndexUrl = 'https://www.lfd.uci.edu/~gohlke/pythonlibs/'
    static hidden [string] $PluginDescription = @"
Provides search, install and local caching for packages published at Christoph Gohlke's page: https://www.lfd.uci.edu/~gohlke/pythonlibs/, Laboratory for Fluorescence Dynamics, University of California, Irvine.
"@
    static hidden [string] $InputPackageType = 'wheel_irvine'
    static hidden [string] $OutputPackageType = 'wheel'
    static hidden [string] $DatabaseFileName = 'packages.bin'

    hidden [hashtable] $Database

    hidden [string] $PluginConfigurationDirectory
    hidden [System.Func[string, bool]] $CreatePath
    hidden [System.Func[string, string, hashtable]] $DownloadArchive
    hidden [System.Func[string, byte[]]] $DownloadString
    hidden [System.Action[string]] $WriteLog
    hidden [System.Func[string, bool]] $FileExists
    hidden [System.Func[object, string, bool]] $Serialize
    hidden [System.Func[string, object]] $Deserialize
    hidden [System.Func[string, hashtable]] $GetPackageInfoFromWheelName
    hidden [System.Func[string, string, System.Func[hashtable, bool]]] $TestCanInstallPackageTo
    hidden [System.Func[string, bool]] $TestPackageInIndex
    hidden [System.Action] $DeferredInit

    Irvine() {
    }

    [void] Init($PluginConfigurationDirectory, $DownloadArchive, $DownloadString, $WriteLog,
        $FileExists, $Serialize, $Deserialize, $GetPackageInfoFromWheelName, $TestCanInstallPackageTo,
        $TestPackageInIndex, $DeferredInit) {

        $this.PluginConfigurationDirectory = $PluginConfigurationDirectory
        $this.DownloadArchive = $DownloadArchive
        $this.DownloadString = $DownloadString
        $this.WriteLog = $WriteLog
        $this.FileExists = $FileExists
        $this.Serialize = $Serialize
        $this.Deserialize = $Deserialize
        $this.GetPackageInfoFromWheelName = $GetPackageInfoFromWheelName
        $this.TestCanInstallPackageTo = $TestCanInstallPackageTo
        $this.TestPackageInIndex = $TestPackageInIndex
        $this.DeferredInit = $DeferredInit

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
        return [Irvine]::PluginDescription
    }

    [array] GetSupportedPackageTypes() {
        return @( [Irvine]::InputPackageType )
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
                # $meta = "py=$($wheelInfo.Python) abi=$($wheelInfo.ABI) plat=$($wheelInfo.Platform) file=$($wheel.filename) |"

                if (-not $FuncCanInstall.Invoke($wheelInfo)) {
                    continue
                }

                [void] $results.Add(@{
                    'Name'=$name;
                    'Version'=$wheelInfo.version;
                    # 'Description'="$meta $($info.description)";  # for debugging
                    'Description'=$($info.description);
                    'Type'=[Irvine]::InputPackageType;
                    })

            }
        }

        return $results
    }

    [hashtable] PackageActionHook([string] $name, [string] $version, [string] $type,
                               [string] $pythonVersion, [string] $pythonArch, [ref] $hookError) {
        if ($type -ne [Irvine]::InputPackageType) {
            return $null
        }

        $location = $this.GetArchiveLocation($name, $version, $type, $pythonVersion, $pythonArch)

        if (-not [string]::IsNullOrEmpty($location)) {
            return @{
                'Name'=$location;
                'Version'=$version;
                'Type'=[Irvine]::OutputPackageType;
            }
        } else {
            $hookError.Value = $true
            return $null
        }
    }

    [string] GetPackageHomepage([string] $name, [string] $type) {
        if ($type -ne [Irvine]::InputPackageType) {
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
            $this.WriteDatabaseInfoToLog()
        } else {
            [void] $this.WriteLog.Invoke('Package index not found.')
        }
    }

    hidden [void] WriteDatabaseInfoToLog() {
        [void] $this.WriteLog.Invoke(@"
Loaded package index: $($this.Database.items.Count) packages
Packages origin updated: $($this.Database.originTimestamp)
Index creaded: $($this.Database.databaseTimestamp)
"@)
    }

    hidden [void] SaveDatabase() {
        if ($this.Database) {
            $path = $this.GetLocallyCreatedDatabasePath()
            [void] $this.Serialize.Invoke($this.Database, $path)
        }
    }

    hidden [void] UpdateDatabase() {
        $this.DeferredInit.Invoke()
        [byte[]] $bytes = $this.DownloadString.Invoke([Irvine]::IndexUrl)

        # $cache = "$($this.PluginConfigurationDirectory)\cache\index.html"
        # [IO.File]::WriteAllBytes($cache, $bytes)  # cache for debugging
        # $bytes = [IO.File]::ReadAllBytes($cache)

        if ((-not $bytes) -or ($bytes.Length -eq 0)) {
            $this.WriteLog.Invoke("Update failed, couldn't get hands on $([Irvine]::IndexUrl)")
            return
        }
        [string] $html = $this.TidyUpText($bytes)
        $document = $this.ParseHTML($html)
        $db = $this.ExtractPackageDatabase($document)
        if ($db) {
            $this.Database = $db
            $this.WriteDatabaseInfoToLog()
            $this.SaveDatabase()
        }
    }

    hidden [string] TidyUpText([byte[]] $bytes) {

        Function Recode($src, $dst, $text, [switch] $BOM, [switch] $AsBytes) {
            if (-not $AsBytes) {
                [byte[]] $bytes = $src.GetBytes($text)
            } else {
                [byte[]] $bytes = $text
            }
            [byte[]] $conv = [System.Text.Encoding]::Convert($src, $dst, $bytes)
            if ($BOM) {
                [System.Collections.Generic.List[byte]] $buffer = [System.Collections.Generic.List[byte]]::new()
                $buffer.AddRange($conv)
                switch ($dst)
                {
                    ([System.Text.Encoding]::UTF8) {
                        # 0xEF,0xBB,0xBF
                        $buffer.Insert(0, 0xBF)
                        $buffer.Insert(0, 0xBB)
                        $buffer.Insert(0, 0xEF)
                    }
                    ([System.Text.Encoding]::Unicode) {
                        # 0xFE,0xFF big
                        # 0xFF,0xFE little
                        $buffer.Insert(0, 0xFE)
                        $buffer.Insert(0, 0xFF)
                    }
                }
                $conv = $buffer.ToArray()
            }
            if ($AsBytes) {
                return ,$conv
            }
            [string] $res = $dst.GetString($conv, 0, $conv.Length)
            return ,$res
        }

        $text = [System.Text.Encoding]::UTF8.GetString($bytes)

        [regex]::Replace($text, '&[a-z]+;|&#x?\d+;', {
            param($entity)
                switch ($entity) {
                    { $_ -in @('&nbsp;', '&#160;')} {
                        return ' '
                    }
                    Default {
                        return $entity
                    }
                }
            })

        $text = [System.Net.WebUtility]::HtmlDecode($text)

        $CharHyphen = [char] 0x2011
        $text = $text.Replace($CharHyphen, '-')

        $text = Recode ([System.Text.Encoding]::UTF8) ([System.Text.Encoding]::ASCII) $text

        return ,$text
    }

    hidden [object] ParseHTML([string] $text) {
        $document = New-Object mshtml.HTMLDocumentClass
        $document.IHTMLDocument2_write($text)
        $document.Close()
        return $document
    }

    hidden [hashtable] ExtractPackageDatabase([object] $document) {

        Function ToTitleCase($text) {
            return (Get-Culture).TextInfo.ToTitleCase($text)
        }

        Function DecryptUrl($data) {
            $base = 'https://download.lfd.uci.edu/pythonlibs/'

            $data = [System.Web.HttpUtility]::HtmlDecode($data)
            $parser = [regex] 'dl\(\[(?<Numbers>[^\]]+)],\s*"(?<Code>[^"]+)"'
            $groups = $parser.Match($data).Groups
            ($ml, $mi) = ($groups['Numbers'] -split ','), $groups['Code'].Value

            $mi = $mi -replace '&lt;','<'
            $mi = $mi -replace '&#62;','>'
            $mi = $mi -replace '&#38;','&'

            $ot = [System.Text.StringBuilder]::new()
            for ($j = 0; $j -lt $mi.Length; $j++) {
                [void]$ot.Append([char][int] ($ml[ ([int][char]$mi[$j]) - 47]) )
            }

            return "${base}$($ot.ToString())"
        }

        Function GuessNameMatch($a, $b) {
            if ($a -eq $b) {
                return $true
            }

            $a = $a.ToLower()
            $b = $b.ToLower()

            $a = $a -replace '\.|-|_|\d+',''
            $b = $b -replace '\.|-|_|\d+',''

            if ($a.Contains($b) -or $b.Contains($a)) {
                return $true
            }

            $a = $a -replace '^py(thon)?',''
            $b = $b -replace '^py(thon)?',''

            if ($a.Contains($b) -or $b.Contains($a)) {
                return $true
            }
        }

        $db = @{
            'description'='Laboratory for Fluorescence Dynamics, University of California, Irvine';
            'maintainer'='Christoph Gohlke';
            'originTimestamp'=($document.body.getElementsByClassName('date').item().innerText);
            'databaseTimestamp'=(Get-Date -Format f);
            'url'=[Irvine]::IndexUrl;
            'items'=@{};
        }

        $packagesFound = 0
        $wheelsFound = 0
        $wheelsSkipped = 0
        $nameMismatches = 0
        $previouslyUnknown = 0

        $links = $document.IHTMLDocument3_getElementsByTagName('a')
        foreach ($link in $links) {
            if ($link.onclick -match 'javascript:dl') {
                # $title = [System.Web.HttpUtility]::HtmlDecode($link.title)  # link on-hover tooltip: size and date
                $wheel = $link.innerHtml  # wheel filename
                $strong = $link.parentElement.parentElement.parentElement.children[1]  # <strong> with the package name
                $info = $strong.children[0]
                $url = $info.href  # package home page
                $displayName = $strong.innerText  # actual package name
                $name = $null

                if (-not [string]::IsNullOrWhitespace($displayName)) {
                    $displayName = $displayName.Trim()
                    $name = $displayName.ToLower()
                } else {
                    $this.WriteLog.Invoke("Can't parse name for '$wheel'.")
                }

                $miscSection = $displayName -eq 'Misc'

                $wheelInfo = $this.GetPackageInfoFromWheelName.Invoke($wheel)
                if (($wheelInfo -eq $null) -or [string]::IsNullOrWhiteSpace($wheelInfo.Distribution)) {
                    $this.WriteLog.Invoke("Can't parse wheel file name: $wheel. Skipped.")
                    ++$wheelsSkipped
                    continue
                }

                $ignoreDescription = $false
                $description = ''

                if ([string]::IsNullOrWhiteSpace($displayName)) {
                    $displayName = $wheelInfo.Distribution  # take from wheel info
                    $name = $displayName.ToLower()
                } elseif(-not (GuessNameMatch $displayName $wheelInfo.Distribution)) {
                    if (-not $miscSection) {  # muzzle for Misc
                        $this.WriteLog.Invoke("Names don't match: package '$displayName' != '$($wheelInfo.Distribution)' wheel filename")
                        ++$nameMismatches
                    }
                    $ignoreDescription = $true
                    $displayName = $wheelInfo.Distribution
                    $name = $displayName
                }

                if (-not $ignoreDescription) {
                    # assemble a package description from pieces
                    $descriptionFragmentElement = $strong.nextSibling
                    $description = [System.Text.StringBuilder]::new()
                    while ($descriptionFragmentElement.nodeName -ne 'UL') {  # until we meet the <UL> list of wheels
                        [void]$description.Append($descriptionFragmentElement.textContent)
                        $descriptionFragmentElement = $descriptionFragmentElement.nextSibling
                    }
                    $description = ToTitleCase($description.ToString().Trim() -replace '^,\s*','')

                    if ([string]::IsNullOrWhiteSpace($description) -and (-not $miscSection)) {
                        $this.WriteLog.Invoke("Can't get description for '$wheel'")
                    }
                }

                if (-not $db.items."$name") {
                    $db.items."$name" = @{
                        'displayName'=$displayName;
                        'description'=$description;
                        'url'=$url;
                        'wheels'=([System.Collections.ArrayList]::new());
                    }
                    ++$packagesFound

                    if (-not ($this.TestPackageInIndex.Invoke($name))) {
                        $this.WriteLog.Invoke("NEW package: '$name'")
                        ++$previouslyUnknown
                    }
                }

                [void] $db.items."$name".wheels.Add(@{
                    'filename'=$wheel;
                    # 'title'=$title;
                    'url'=(DecryptUrl $link.onclick);
                })

                ++$wheelsFound
            }
        }

        if ($packagesFound -eq 0) {
            $this.WriteLog.Invoke("Couldn't extract a single package! Index's not rebuilt.")
            return $null
        } else {
            $this.WriteLog.Invoke("Done with $wheelsFound wheels okay, $packagesFound packages, skipped $wheelsSkipped wheels, $nameMismatches name issues resolved.")
            if ($previouslyUnknown) {
                $this.WriteLog.Invoke("Found $previouslyUnknown previously unknown packages.")
            }
        }

        return $db
    }

}

Function NewPluginInstance() {
    return [Irvine]::new()
}

Export-ModuleMember -Function NewPluginInstance
