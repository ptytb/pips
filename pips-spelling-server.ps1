Import-Module $PSScriptRoot\BK-tree\bktree

$bktree = [BKTree]::new()
$bktree.LoadArrays("$PSScriptRoot\known-packages-bktree.bin")

'Loaded dict.'

$pipe = new-object System.IO.Pipes.NamedPipeServerStream("\\.\pipe\pips_spelling_server");

'Created server side of "\\.\pipe\pips_spelling_server"'

$pipe.WaitForConnection(); 
 
$sr = new-object System.IO.StreamReader($pipe); 
$sw = new-object System.IO.StreamWriter($pipe); 

while (($text = $sr.ReadLine()) -ne $null) 
{
  	Write-Host $text
	$request = $text | ConvertFrom-Json
	$candidates = $bktree.SearchFast($request.Request, $request.Distance)
	$json = @{ 'Request'=$request.Request; 'Candidates'=$candidates; } | ConvertTo-Json -Depth 5 -Compress
	$sw.WriteLine($json)
	$sw.Flush()
	$text = $null
}; 
 
$sr.Dispose();
$pipe.Dispose();
