# a little mess in here, work in progress

$env:PYTHONIOENCODING="utf-8"
$env:LC_CTYPE="utf-8"

[Void][Reflection.Assembly]::LoadWithPartialName("System.Collections")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Collections.Hashtable")
[Void][Reflection.Assembly]::LoadWithPartialName("System.IO.FileStream")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Web")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Web.HttpUtility")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

#$index = New-Object System.Collections.Hashtable
$index = New-Object System.Collections.Generic.HashSet[String]
$r = [regex] '^(?<Name>.*?)\s*\((?<Version>.*?)\)\s+-\s+(?<Description>.*?)$'
$count = 0
$tags_url = 'https://pypi.python.org/pypi?%3Aaction=list_classifiers'

# The correct way to get the whole list of packages
$all_list_html = 'https://pypi.python.org/simple/'

#$webClient = New-Object System.Net.WebClient
#$tags = $webClient.DownloadString($tags_url)

Function Load-KnownPackageIndex {
    $fs = New-Object System.IO.FileStream "$PSScriptRoot\known-packages.bin", ([System.IO.FileMode]::Open)
    $bf = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
    $index = $bf.Deserialize($fs)
    $fs.Close()
    return $index
}

#$index = Load-KnownPackageIndex

$lines = Get-Content "$PSScriptRoot\index.txt" | Where {$_ -notmatch '^\s+$'} | foreach { $_.ToString().Trim() } 

foreach ($l in $lines) {
#    Write-Host "Getting $l"

    #if ($index.ContainsKey($l)) {
        #continue
    #}
    #$index[$l] = @{Version=''; Description=''}
    $index.Add($l) | Out-Null
    $count += 1
}

#[System.Management.Automation.PSSerializer]::Serialize($index) > known-packages.xml

$fs = New-Object System.IO.FileStream 'known-packages.bin', ([System.IO.FileMode]::Create)
$bf = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
$bf.Serialize($fs, $index)
$fs.Close()

Write-Host "Acquired $count records. Total $($index.Count) in index."
