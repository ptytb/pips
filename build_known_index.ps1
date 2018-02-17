$env:PYTHONIOENCODING="utf-8"
$env:LC_CTYPE="utf-8"

[Void][Reflection.Assembly]::LoadWithPartialName("System.Collections")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Collections.Hashtable")
[Void][Reflection.Assembly]::LoadWithPartialName("System.IO.FileStream")


$index = New-Object System.Collections.Hashtable
$r = [regex] '^(?<Name>.*?)\s*\((?<Version>.*?)\)\s+-\s+(?<Description>.*?)$'
$count = 0

# search for a..z as long as `pip search *' does not work; still not the whole list, unfortunately

foreach ($l in [char]'a'..[char]'z') {
    $letter = [char] $l
    Write-Host "Getting $letter"
    $output = & pip search $letter
    
    foreach ($line in $output) {
        $m = $r.Match($line)
        if ([String]::IsNullOrEmpty($m.Groups['Name'].Value)) {
            continue
        }
        if ($index.ContainsKey($m.Groups['Name'].Value)) {
            continue
        }
        $index[$m.Groups['Name'].Value] = @{Version=$m.Groups['Version'].Value; Description=$m.Groups['Description'].Value}
        $count += 1
    }
}

#[System.Management.Automation.PSSerializer]::Serialize($index) > known-packages.xml

$fs = New-Object System.IO.FileStream 'known-packages.bin', ([System.IO.FileMode]::CreateNew)
$bf = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
$bf.Serialize($fs, $index)
$fs.Close()
