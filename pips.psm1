$PSDefaultParameterValues['*:Encoding'] = 'UTF8'


$pips_pipe_instance = [System.IO.Directory]::GetFiles("\\.\\pipe\\") | Where-Object { $_ -match 'pips_spelling_server'}
if ($pips_pipe_instance) {
    [System.Windows.Forms.MessageBox]::Show(
        "There's another pips instance running, exiting.",
        "pips",
        [System.Windows.Forms.MessageBoxButtons]::OK)
    ${function:global:Start-Main} = { Exit }
    Return
}


$global:FRAMEWORK_VERSION = [version]([Runtime.InteropServices.RuntimeInformation]::FrameworkDescription -replace '^.[^\d.]*','')

[Void][Reflection.Assembly]::LoadWithPartialName("System")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Drawing.Size")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Drawing.Point")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.MessageBox")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.FontStyle")
[Void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms.VisualStyles')
[Void][Reflection.Assembly]::LoadWithPartialName("System.Text")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Text.RegularExpressions")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Collections")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Collections.ArrayList")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Collections.Generic")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Collections.Generic.HashSet")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Web")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Web.HttpUtility")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Net")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Net.WebClient")

[Void][Reflection.Assembly]::LoadWithPartialName("System.Management")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Management.Automation")


[PSObject].Assembly.GetType("System.Management.Automation.TypeAccelerators")::add(
    'MaybeColor', [Nullable[System.Drawing.Color]])


Function global:MakeEvent([hashtable] $properties) {
    [EventArgs] $EventArgs = [EventArgs]::Empty
    foreach ($p in $properties.GetEnumerator()) {
        Add-Member -InputObject $EventArgs -MemberType 'NoteProperty' -Name $p.Key -Value $p.Value -Force
    }
    return ,$EventArgs
}

Function global:HasUpperChars([string] $text) {
    return [char[]] $text | ForEach-Object `
        { $HasUpperChars = $false } `
        { $HasUpperChars = $HasUpperChars -or [char]::IsUpper($_) } `
        { $HasUpperChars }
}

Function global:Sort-Versions {
    param($MaybeVersions, [switch] $Descending)
    [version] $none = [version]::new()
    [version] $version = [version]::new()
    $predicate = {
        $success = [version]::TryParse(($_ -replace '[^\d.]',''), [ref] $version)
        if ($success) { [version] $version } else { $none }
    }.GetNewClosure()
    return $MaybeVersions | Sort-Object -Property @{ Expression=$predicate; Descending=$Descending } @args
}

Function global:Recode($src, $dst, $text, [switch] $BOM, [switch] $AsBytes) {
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

Function global:ApplyAsync([object] $Queue, [Func[object, object]] $Function, [delegate] $Finally) {
    $delegate = New-RunspacedDelegate([Action[System.Threading.Tasks.Task[Tuple[object, delegate, Func[object, object], delegate]]]] {
        param($task)
        $queue, $delegate, $function = $task.Result.Item1, $task.Result.Item2, $task.Result.Item3
        if ($queue.Count -gt 0) {
            $element = $queue.Dequeue()

            $taskFromFunction = $function.Invoke(@{
                Element=$element;
                ApplyAsyncContext=$task.Result;
                ApplyAsyncContextType=($task.Result.GetType().ToString())
            })

            $null = $taskFromFunction.ContinueWith($delegate, [System.Threading.Tasks.TaskScheduler]::FromCurrentSynchronizationContext())
        } else {
            $finally = $task.Result.Item4
            $null = $task.ContinueWith($finally, [System.Threading.Tasks.TaskScheduler]::FromCurrentSynchronizationContext())
        }
    })

    $coldStart = [System.Threading.Tasks.TaskCompletionSource[Tuple[object, delegate, Func[object, object], delegate]]]::new();
    $null = $coldStart.Task.ContinueWith($delegate, [System.Threading.Tasks.TaskScheduler]::FromCurrentSynchronizationContext())
    $coldStart.SetResult([Tuple[object, delegate, Func[object, object], delegate]]::new($Queue, $delegate, $Function, $Finally))
    return $coldStart.Task
}


$global:WM_USER = [int] 0x0400
$global:WM_SETREDRAW = [int] 0x0B
$global:EM_SETEVENTMASK = [int] ($WM_USER + 69);
$global:WM_CHAR = [int] 0x0102
$global:WM_SCROLL = [int] 276
$global:WM_VSCROLL = [int] 277
$global:VK_BACKSPACE = [int] 0x08
$global:SB_LINEUP = [int] 0x00
$global:SB_LINEDOWN = [int] 0x01
$global:SB_LINELEFT = [int] 0x00
$global:SB_LINERIGHT = [int] 0x01
$global:SB_PAGEUP = [int] 0x02
$global:SB_PAGEDOWN = [int] 0x03
$global:SB_PAGETOP = [int] 0x06
$global:SB_PAGEBOTTOM = [int] 0x07
$MemberDefinition='
[DllImport("User32.dll")]public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, int lParam);
'
$API = Add-Type -MemberDefinition $MemberDefinition -Name 'WinAPI_SendMessage' -PassThru
${global:SendMessage} = $API::SendMessage


$null = Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
public class RichTextBoxEx : System.Windows.Forms.RichTextBox
{
    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    static extern IntPtr LoadLibrary(string lpFileName);

    protected override CreateParams CreateParams
    {
        get
        {
            CreateParams prams = base.CreateParams;
            if (LoadLibrary("msftedit.dll") != IntPtr.Zero)
            {
                prams.ClassName = "RICHEDIT50W";
            }
            return prams;
        }
    }
}
'@ -ReferencedAssemblies 'System.Windows.Forms.dll'
$global:RichTextBox_t = [RichTextBoxEx]


Add-Type -Name TerminateGracefully -Namespace Console -MemberDefinition @'
// https://stackoverflow.com/questions/813086/can-i-send-a-ctrl-c-sigint-to-an-application-on-windows
// https://docs.microsoft.com/en-us/windows/console/generateconsolectrlevent

delegate bool ConsoleCtrlDelegate(CtrlTypes CtrlType);

// Enumerated type for the control messages sent to the handler routine
public enum CtrlTypes : uint
{
  CTRL_C_EVENT = 0,
  CTRL_BREAK_EVENT,
  CTRL_CLOSE_EVENT,
  CTRL_LOGOFF_EVENT = 5,
  CTRL_SHUTDOWN_EVENT
}

[DllImport("kernel32.dll", SetLastError = true)]
static extern bool AllocConsole();

[DllImport("kernel32.dll", SetLastError = true)]
static extern int AttachConsole(uint dwProcessId);

[DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
static extern bool FreeConsole();

[DllImport("kernel32.dll")]
static extern bool SetConsoleCtrlHandler(ConsoleCtrlDelegate HandlerRoutine, bool Add);

[DllImport("kernel32.dll")]
[return: MarshalAs(UnmanagedType.Bool)]
private static extern bool GenerateConsoleCtrlEvent(CtrlTypes dwCtrlEvent, uint dwProcessGroupId);

public static void StopProgram(int pid, CtrlTypes signal)
{
  if (AttachConsole((uint) pid) == 0) {
    AllocConsole();
    AttachConsole((uint) pid);
  }

  SetConsoleCtrlHandler(null, true);
  GenerateConsoleCtrlEvent(signal, 0);
  FreeConsole();
  System.Diagnostics.Process.GetProcessById(pid).WaitForExit(2000);
  SetConsoleCtrlHandler(null, false);
}
'@

Function global:TryTerminateGracefully([System.Diagnostics.Process] $process) {
    [Console.TerminateGracefully]::StopProgram($process.Id, [Console.TerminateGracefully+CtrlTypes]::CTRL_C_EVENT)
    if ($process.HasExited) {
        return
    }
    [Console.TerminateGracefully]::StopProgram($process.Id, [Console.TerminateGracefully+CtrlTypes]::CTRL_BREAK_EVENT)
    if ($process.HasExited) {
        return
    }
    try {
        $process.StandardInput.Close()
    } catch { }
    $process.Kill()
}


Import-Module -Global .\PSRunspacedDelegate\PSRunspacedDelegate


$startServer = New-RunspacedDelegate ( [Func[Object]] {
    Write-Host Start server
    Start-Process -WindowStyle Hidden -FilePath powershell -ArgumentList "-ExecutionPolicy Bypass $PSScriptRoot\pips-spelling-server.ps1"
    Write-Host Server started.
});
$task = [System.Threading.Tasks.Task[Object]]::new($startServer);
$continuation = New-RunspacedDelegate ( [Action[System.Threading.Tasks.Task[Object]]] {
    Write-Host Connecting.

    $pipe = $null
    while (-not $pipe -or -not $pipe.IsConnected) {
        $pipe = new-object System.IO.Pipes.NamedPipeClientStream("\\.\pipe\pips_spelling_server");
        if ($pipe) {
            $milliseconds = 250
            try {
                $pipe.Connect($milliseconds);
            } catch {
            }
        } else {
            Write-Host reconnecting
        }
    }
    Write-Host connected!
    $Global:sw = new-object System.IO.StreamWriter($pipe);
    $Global:sr = new-object System.IO.StreamReader($pipe);
    $Global:sw.AutoFlush = $false
    [bool] $Global:SuggestionsWorking = $false
});

$task.ContinueWith($continuation);
$task.Start()


$Global:FuncRPCSpellCheck = {
    param([string] $text, [int] $distance)

    if ([String]::IsNullOrEmpty($text) -or ($text.IndexOfAny('=@\/:') -ne -1)) {
        return
    }

    if ($Global:SuggestionsWorking) {
        return
    } else {
        $Global:SuggestionsWorking = $true
    }

    $text = $text.ToLower()
    $request = @{ 'Request'=$text; 'Distance'=$distance; } | ConvertTo-Json -Depth 5 -Compress

	$tw = $Global:sw.WriteLineAsync($request);
	$continuation1 = New-RunspacedDelegate ( [Action[System.Threading.Tasks.Task]] {
    	$tf = $Global:sw.FlushAsync();
    	$continuation2 = New-RunspacedDelegate ( [Action[System.Threading.Tasks.Task]] {
        	$tr = $Global:sr.ReadLineAsync()
            $null = $tr.ContinueWith($global:FuncRPCSpellCheck_Callback);
    	});
        [void]$tf.ContinueWith($continuation2);
	});
    [void]$tw.ContinueWith($continuation1);
}


Function global:Get-Bin($command, [switch] $All) {
    $arguments = @{
        All=$All
    }
    $commands = Get-Command @arguments -ErrorAction SilentlyContinue -CommandType Application -Name $command
    if ($commands) {
        $commands = $commands | Select-Object -ExpandProperty Source
        if ($all) {
            $found = @($commands)
        } else {
            $found = $commands
        }
    } else {
        if ($all) {
            $found = @()
        } else {
            $found = $null
        }
    }
    return ,$found
}

Function global:Guess-EnvPath ($path, $fileName, [switch] $directory, [switch] $Executable) {
    $subdirs = @('\'; '\Scripts\'; '\.venv\Scripts\'; '\.venv\'; '\env\Scripts\'; '\env\'; '\bin\')
    foreach ($tryPath in $subdirs) {
        $target = "${path}${tryPath}${fileName}"
        if ($directory) {
            if (Exists-Directory $target) {
                return $target
            }
        } else {
            if ($Executable) {
                $executableExtensions = @('exe'; 'bat'; 'cmd'; 'ps1')
                foreach ($tryExtension in $executableExtensions) {
                    $target = "${path}${tryPath}${fileName}.${tryExtension}"
                    if (Exists-File $target) {
                        return $target
                    }
                }
            } else {
                if (Exists-File $target) {
                    return $target
                }
            }
        }
    }
    return $null
}

Function global:Get-CurrentInterpreter($item, [switch] $Executable) {
    if (-not [string]::IsNullOrEmpty($item)) {
        $item = $Global:interpretersComboBox.SelectedItem."$item"
        if ($Executable) {
            $item = Get-ExistingFilePathOrNull $item
        }
        return $item
    } else {
        return $Global:interpretersComboBox.SelectedItem
    }
}

Function global:Delete-CurrentInterpreter() {
    if (-not (Get-CurrentInterpreter 'User')) {
        WriteLog 'Can only delete venv which was added manually with env:Open or env:Create.'
        return
    }

    $path = Get-CurrentInterpreter 'Path'
    $failedToRemove = $false

    $response = ([System.Windows.Forms.MessageBox]::Show(
        "Remove venv path completely?`n`n${path}`n`n" +
        "Yes = Remove directory`nNo = Keep directory, forget entry`nCancel = Do nothing",
        "Removing venv", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel))

    if ($response -eq 'Yes') {
        $stats = Get-ChildItem -Recurse -Path $path | Measure-Object
        try {
            Remove-Item -Path $path -Recurse -Force
        } catch { }
        if (Exists-Directory $path) {
            WriteLog "Cannot delete '$path'"
            $failedToRemove = $true
        } else {
            WriteLog "Removed $($stats.Count) items."
        }
    }

    if ((($response -eq 'No') -or ($response -eq 'Yes')) -and -not $failedToRemove) {
        $interpreters.Remove($interpretersComboBox.SelectedItem)
        $Script:trackDuplicateInterpreters.Remove($path)
        $interpretersComboBox.DataSource = $null
        $interpretersComboBox.DataSource = $interpreters
        WriteLog "Removed venv '${path}' from list."

        $interpretersComboBox.SelectedIndex = 0
        WriteLog "Switching to '$(Get-CurrentInterpreter 'Path')'"
    }
}

Function global:Get-PipsSetting($name, [switch] $AsArgs, [switch] $First) {
    switch ($name)
    {
        "CondaChannels" {
            $channels = $global:settings.condaChannels
            if (-not $channels -or $channels.Length -eq 0) {
                $channels = @('anaconda'; 'defaults'; 'conda-forge')
            }
            if ($First) {
                $channels = @($channels[0])
            }
            if ($AsArgs) {
                $channels = "-c $($channels -join ' -c ')"
            }
            return $channels
        }

        Default { throw "No such setting: $name" }
    }
}

Function global:Exists-File($path, [string] $Mask) {
    if ($Mask) {
        $candidates = $null
        try {
            $candidates = Get-ChildItem -Path $path -Filter $Mask -Depth 0 -File -ErrorAction Stop | Select-Object -First 1
        } catch { }
        return ($candidates -ne $null)
    }
    return [System.IO.File]::Exists($path)
}

$global:invalidPathCharacters = [System.IO.Path]::GetInvalidPathChars()
Function global:Exists-Directory($path) {
    if ([string]::IsNullOrWhiteSpace($path)) {
        return $false
    }
    if ($path.IndexOfAny($global:invalidPathCharacters) -ne -1) {
        return $false
    }
    return [System.IO.Directory]::Exists($path)
}

Function Get-ExistingFilePathOrNull($path) {
    if (Exists-File $path) {
        return $path
    } else {
        return $null
    }
}

Function Get-ExistingPathOrNull($path) {
    if (Exists-Directory $path) {
        return $path
    } else {
        return $null
    }
}

$pypi_url = 'https://pypi.python.org/pypi/'
$anaconda_url = 'https://anaconda.org/search?q='
$peps_url = 'https://www.python.org/dev/peps/'
$github_search_url = 'https://api.github.com/search/repositories?q={0}+language:python&sort=stars&order=desc'
$github_url = 'https://github.com'
$python_releases = 'https://www.python.org/downloads/windows/'

$lastWidgetLeft = 5
$lastWidgetTop = 5
$widgetLineHeight = 23
$global:dataGridView = $null
$global:inputFilter = $null
$global:actionsModel = $null
$global:dataModel = $null  # [DataRow] keeps actual rows for the table of packages
$isolatedCheckBox = $null
$header = ("Select", "Package", "Installed", "Latest", "Type", "Status")
$csv_header = ("Package", "Installed", "Latest", "Type", "Status")
$search_columns = ("Select", "Package", "Installed", "Description", "Type", "Status")
$formLoaded = $false
$outdatedOnly = $true
$interpreters = $null
$global:autoCompleteIndex = $null
$Global:interpretersComboBox = $null

enum InstallAutoCompleteMode {
  Name;
  Version;
  Directory;
  GitTag;
  WheelFile;
  None
}

enum AppMode {
    Idle;
    Working
}

[AppMode] $global:APP_MODE = [AppMode]::Idle

$global:packageTypes = [System.Collections.ArrayList]::new()
$global:packageTypes.AddRange(@('pip', 'conda', 'git', 'wheel', 'https'))


$iconBase64_DownArrow = @'
iVBORw0KGgoAAAANSUhEUgAAAAsAAAALCAYAAACprHcmAAAABGdBTUEAALGPC/xhBQAAAAlwSFlz
AAAOvAAADrwBlbxySQAAABp0RVh0U29mdHdhcmUAUGFpbnQuTkVUIHYzLjUuMTAw9HKhAAAAP0lE
QVQoU42KQQ4AIAzC9v9Po+Pi3MB4aAKFAPCNlA4pHVI6TtjRMc4s7ZRcey0U5sitC0pxTIZ4IaVD
Sg1iAai9ScU7YisTAAAAAElFTkSuQmCC
'@

$iconBase64_Snakes = @'
AAABAAEAEBAAAAAAAABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAQAQAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYAAAAaAAAAIQAAACEAAAAYAAAABQAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAODg4CPw8PDD/Pz8+P/////7+/v44uLizVtbW1AA
AAAGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD9/f24s/X//13j//9I2f//SdP//6Dm
///i4uLPAAAAHAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////9mXt//9c6P//UeD/
/9X2//8/0P///Pz8+QAAACgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYAAAAZAAAAIf////9i7f//
Ye3//1rn//9P3v//RNb///////8AAAA7AAAAIAAAABcAAAAFAAAAAN/f3yPw8PDC/Pz8+P//////
////Yu3//2Lt//++9/////////////////////////v7+/ji4uLNWlpaUAAAAAb9/f25zLql/4hq
Rv99ZEX//////2Pt//9i7f//Yu3//2Dr//9W4///S9v//0DS//85y///neT//+Li4s8AAAAd////
9qF6S/+PbEH/hGdD//Hu6v+b9P//Y+3//2Lt//9i7f//Xur//1Ti//9J2f//PtH//0DM///8/Pz5
AAAAKf////+kdj3/mXE//41rQv+jjnT/8O7q/////////////////+39//+F7f//UuD//0fY//88
z////////wAAACr////1soJH/6J1Pf+Wb0D/i2pC/4BlRP98Y0X/fGNF/31kRv+snIn/7f3//1vn
//9P3///UNn///v7+/gAAAAh////ttzCov+tfD//oHQ+/5RuQP+JaUL/fmRF/3xjRf98Y0X/fWVH
//////9h7P//XOb//6zv///r6+vEAAAACwAAAAD///+4////9//////////////////////JvrH/
fGNF/3xjRf////////////7+/vX6+vq4v7+/JQAAAAAAAAAAAAAAAAAAAAAAAAAA/////5tyP/+Q
bEH/hWdD/31jRf98Y0X//////wAAACkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//
//WmeUD/6N/T/45rQf+CZkT/f2ZJ//v7+/gAAAAhAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAD///+22MCj/6h9Sf+XcED/k3NN/8S3qP/r6+vEAAAACwAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAP///7j////3//////7+/vX6+vq5v7+/JQAAAAAAAAAAAAAAAAAAAAAA
AAAA+B8AAPAPAADwDwAA8A8AAIABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAQAA8A8AAPAP
AADwDwAA+B8AAA==
'@


$FuncSetWebClientWorkaround = {
    Function Set-UseUnsafeHeaderParsing {
        param(
            [Parameter(Mandatory,ParameterSetName='Enable')]
            [switch]$Enable,

            [Parameter(Mandatory,ParameterSetName='Disable')]
            [switch]$Disable
        )

        $ShouldEnable = $PSCmdlet.ParameterSetName -eq 'Enable'

        $netAssembly = [Reflection.Assembly]::GetAssembly([System.Net.Configuration.SettingsSection])

        if ($netAssembly)
        {
            $bindingFlags = [Reflection.BindingFlags] 'Static,GetProperty,NonPublic'
            $settingsType = $netAssembly.GetType('System.Net.Configuration.SettingsSectionInternal')

            $instance = $settingsType.InvokeMember('Section', $bindingFlags, $null, $null, @())

            if ($instance)
            {
                $bindingFlags = 'NonPublic','Instance'
                $useUnsafeHeaderParsingField = $settingsType.GetField('useUnsafeHeaderParsing', $bindingFlags)

                if ($useUnsafeHeaderParsingField)
                {
                  $useUnsafeHeaderParsingField.SetValue($instance, $ShouldEnable)
                }
            }
        }
    }

    [Net.ServicePointManager]::SecurityProtocol = (
        [Net.SecurityProtocolType]::Tls12 -bor `
        [Net.SecurityProtocolType]::Tls11 -bor `
        [Net.SecurityProtocolType]::Tls   -bor `
        [Net.SecurityProtocolType]::Ssl3)

    [Net.ServicePointManager]::ServerCertificateValidationCallback = New-RunspacedDelegate (
        [System.Net.Security.RemoteCertificateValidationCallback]{ $true })

    [Net.ServicePointManager]::Expect100Continue = $true

    Set-UseUnsafeHeaderParsing -Enable
}

& $FuncSetWebClientWorkaround
Function global:Download-String($url, $ContinueWith = $null) {
    try {
        $wc = [System.Net.WebClient]::new()
        $wc.Headers['User-Agent'] = "Mozilla/5.0 (compatible; MSIE 6.0;)"
        $wc.Headers['Accept'] = '*/*'
        $wc.Headers['Accept-Encoding'] = 'identity'
        $wc.Headers['Accept-Language'] = 'en'
        $wc.Encoding = [System.Text.Encoding]::UTF8

        if ($ContinueWith) {
            $delegate = New-RunspacedDelegate ([System.Net.DownloadStringCompletedEventHandler] {
                param([object] $sender, [System.Net.DownloadStringCompletedEventArgs] $e)
                [string] $result = $null

                if ((-not $e.Cancelled) -and ($e.Error -eq $null)) {
                    $result = $e.Result
                }

                if ($result -eq $null) {
                    throw "Failed to Connect to website $url"
                    return
                }

                $null = $ContinueWith.Invoke($result)
            }.GetNewClosure())
            $wc.Add_DownloadStringCompleted($delegate)

            $result = $wc.DownloadStringAsync($url)
        } else {
            $result = $wc.DownloadString($url)
        }
    } catch {
        Write-Error "$url`n$_`n"
        $result = $null
    }
    return $result
}

Function Download-Data($url) {
    try {
        $wc = New-Object System.Net.WebClient
        $wc.Headers["User-Agent"] = "Mozilla/5.0 (compatible; MSIE 6.0;)"
        $wc.Encoding = [System.Text.Encoding]::UTF8
        $result = $wc.DownloadData($url)
    } catch {
        Write-Error "$url`n$_`n"
        $result = $null
    }
    return $result
}

Function Convert-Base64ToBMP($base64Text) {
    $iconStream = [System.IO.MemoryStream][System.Convert]::FromBase64String($base64Text)
    $iconBmp = [System.Drawing.Bitmap][System.Drawing.Image]::FromStream($iconStream)
    return $iconBmp
}

Function Convert-Base64ToICO($base64Text) {
    $iconStream = [System.IO.MemoryStream][System.Convert]::FromBase64String($base64Text)
    $icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Image]::FromStream($iconStream)).GetHIcon())
    return $icon
}


Function global:EnsureColor($color) {
    if (($color -is [string]) -and (-not [string]::IsNullOrWhiteSpace($color))) {
        $color = [System.Drawing.Color]::FromKnownColor($color)
    } elseif ($color -isnot [System.Drawing.Color]) {
        $color = $null
    }
    return [MaybeColor] $color
}

$global:_WritePipLogBacklog = [System.Collections.Generic.List[hashtable]]::new()

Function global:WriteLog {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromRemainingArguments=$true, Position=1)]
        [AllowEmptyCollection()]
        [AllowNull()]
        [object[]] $Lines = @(),
        [switch] $UpdateLastLine,
        [switch] $NoNewline,
        [Parameter(Mandatory=$false)] [AllowNull()] $Background = $null,
        [Parameter(Mandatory=$false)] [AllowNull()] $Foreground = $null
    )

    $arguments = @{
        Lines=$Lines;
        UpdateLastLine=([bool] $PSBoundParameters['UpdateLastLine']);
        NoNewline=([bool] $PSBoundParameters['NoNewline']);
        Background=(EnsureColor $Background);
        Foreground=(EnsureColor $Foreground);
    }

    [void] $global:_WritePipLogBacklog.Add($arguments)
}

Function Add-TopWidget($widget, $span=1) {
    $widget.Location = New-Object Drawing.Point $lastWidgetLeft,$lastWidgetTop
    $widget.size = New-Object Drawing.Point ($span*100-5),$widgetLineHeight
    $Script:form.Controls.Add($widget)
    $Script:lastWidgetLeft = $lastWidgetLeft + ($span*100)
}

Function Add-HorizontalSpacer() {
    $Script:lastWidgetLeft = $lastWidgetLeft + 100
}

Function NewLine-TopLayout() {
    $Script:lastWidgetTop  = $Script:lastWidgetTop + $widgetLineHeight + 5
    $Script:lastWidgetLeft = 5
}

Function Add-Button ($name, $handler, [switch] $AsyncHandler) {
    $button = New-Object Windows.Forms.Button
    $button.Text = $name
    if ($AsyncHandler) {
        $button.Add_Click({
            $widgetStateTransition = WidgetStateTransitionForCommandButton $button
            $task = $handler.Invoke($args)
            $task.ContinueWith(([WidgetStateTransition]::ReverseAllAsync()),
                $widgetStateTransition,
                [System.Threading.Tasks.TaskScheduler]::FromCurrentSynchronizationContext())
        }.GetNewClosure())
    } else {
        $button.Add_Click([EventHandler] $handler)
    }
    Add-TopWidget $button
    return $button
}

Function Add-ButtonMenu ($text, $tools, $onclick) {
    $form = $script:form  # to be captured by $handler's closure

    $handler = {
        param($button)
        $menuStrip = New-Object System.Windows.Forms.ContextMenuStrip

        # Firstly, list non-persistent (more context-related) menu items
        foreach ($tool in $tools) {
            if ($tool.Persistent -or $tool.IsSeparator) {
                continue
            }
            if ($tool.Contains('IsAccessible') -and -not $tool.IsAccessible.Invoke()) {
                continue
            }
            $menuStrip.Items.Add($tool.MenuText)
        }

        if ($menuStrip.Items.Count -gt 0) {
            $menuStrip.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
        }

        foreach ($tool in $tools) {
            if ($tool.IsSeparator) {
                $menuStrip.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
            } elseif ($tool.Persistent) {
                $item = [System.Windows.Forms.ToolStripMenuItem]::new($tool.MenuText)
                $item.Enabled = (-not $tool.Contains('IsAccessible')) -or $tool.IsAccessible.Invoke()
                $menuStrip.Items.Add($item)
            }
        }

        $tools = $Script:tools
        $onclick = $Script:onclick
        $menuStrip.add_ItemClicked({
            foreach ($tool in $tools) {
                if ($tool.MenuText -eq $_.ClickedItem) {
                    $menuStrip.Hide()
                    [void] $onclick.Invoke( @($tool) )
                }
            }
        }.GetNewClosure())

        $point = New-Object System.Drawing.Point ($button.Location.X, $button.Bottom)
        $menuStrip.Show($Script:form.PointToScreen($point))
    }.GetNewClosure()

    $button = Add-Button $text $handler
    $button.ImageAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $button.Image = Convert-Base64ToBMP $iconBase64_DownArrow

    return $button
}

Function Add-Label ($name) {
    $label = New-Object Windows.Forms.Label
    $label.Text = $name
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    Add-TopWidget $label
    return $label
}

Function Add-Input ($handler) {
    $input = New-Object Windows.Forms.TextBox
    $input.Add_TextChanged([EventHandler] $handler)
    Add-TopWidget $input
    return $input
}


Function Add-Buttons {
    $global:WIDGET_GROUP_COMMAND_BUTTONS = @(
        Add-Button "Check Updates" { Get-PythonPackages } ;
        Add-Button "List Installed" { Get-PythonPackages($false) } ;
        Add-Button "Sel All Visible" { Select-VisiblePipPackages($true) } ;
        Add-Button "Select None" { Select-PipPackages($false) } ;
        Add-Button "Check Deps" { Check-PipDependencies } ;
        Add-Button "Execute" { Execute-PipAction } -AsyncHandler ;
    )
}

Function global:Get-PyDoc($request) {
    $requestNormalized = $request -replace '-','_'
    $output = & (Get-CurrentInterpreter 'PythonExe') -m pydoc $requestNormalized

    if ("$output".StartsWith('No Python documentation found')) {
        $output = & (Get-CurrentInterpreter 'PythonExe') -m pydoc ($requestNormalized).ToLower()
    }

    # $output = Recode ([Text.Encoding]::UTF8) ([Text.Encoding]::Unicode) $output

    return $output
}

Function Get-PythonBuiltinPackages() {
    $builtinLibs = New-Object System.Collections.ArrayList
    $path = Get-CurrentInterpreter 'Path'
    $libs = "${path}\Lib"
    $ignore = [regex] '^__'
    $filter = [regex] '\.py.?$'

    $trackDuplicates = New-Object System.Collections.Generic.HashSet[String]

    foreach ($item in dir $libs) {
        if ($item -is [System.IO.DirectoryInfo]) {
            $packageName = "$item"
        } elseif ($item -is [System.IO.FileInfo]) {
            $packageName = "$item" -replace $filter,''
        }
        if (($packageName -cmatch $ignore) -or ($trackDuplicates.Contains($packageName))) {
            continue
        }
        $null = $trackDuplicates.Add("$packageName")
        $null = $builtinLibs.Add([PSCustomObject] @{Package=$packageName; Type='builtin'})
    }

    $getBuiltinsScript = "import sys; print(','.join(sys.builtin_module_names))"
    $sys_builtin_module_names = & (Get-CurrentInterpreter 'PythonExe') -c $getBuiltinsScript
    $modules = $sys_builtin_module_names.Split(',')
    foreach ($builtinModule in $modules) {
        if ($trackDuplicates.Contains("$builtinModule")) {
            continue
        }
        $null = $builtinLibs.Add([PSCustomObject] @{Package=$builtinModule; Type='builtin'})
    }

    return ,$builtinLibs
}

Function Get-PythonOtherPackages {
    $otherLibs = New-Object System.Collections.ArrayList
    $path = Get-CurrentInterpreter 'Path'
    $libs = "${path}\Lib\site-packages"
    $ignore = [regex] '\.dist-info$|\.egg-info$|\.egg$|^__pycache__$'
    $filter = [regex] '\.py.?$'

    if (Exists-Directory $libs) {
        foreach ($item in dir $libs) {
            if ($item -is [System.IO.DirectoryInfo]) {
                if (-not (Exists-File "$libs\$item\__init__.py")) {
                    continue
                }
                $packageName = "$item"
            } elseif ($item -is [System.IO.FileInfo]) {
                if ($packageName -notmatch $filter) {
                    continue
                }
                $packageName = "$item" -replace $filter,''
            }
            if (($packageName -match $ignore) `
                -or ((Test-PackageInList $packageName) -ne -1)`
                -or ((Test-PackageInList ($packageName -replace '_','-')) -ne -1)) {
                continue
            }
            $null = $otherLibs.Add([PSCustomObject] @{Package=$packageName; Type='other'})
        }
    }

    return ,$otherLibs
}

Function Get-CondaPackagesHelper([bool] $outdatedOnly) {
    $conda_exe = Get-CurrentInterpreter 'CondaExe' -Executable
    $arguments = New-Object System.Collections.ArrayList

    if ($outdatedOnly) {
        $null = $arguments.Add('search')
        $null = $arguments.Add('--outdated')
        $null = $arguments.Add((Get-PipsSetting 'CondaChannels' -AsArgs -First))
    } else {
        $null = $arguments.Add('list')
        $null = $arguments.Add('--no-pip')
        $null = $arguments.Add('--show-channel-urls')
    }

    $null = $arguments.Add('--json')
    $null = $arguments.Add('--prefix')
    $null = $arguments.Add((Get-CurrentInterpreter 'Path'))

    $items = & $conda_exe $arguments | ConvertFrom-Json
    return ,$items
}

Function Get-CondaPackages([bool] $outdatedOnly) {
    $condaPackages = New-Object System.Collections.ArrayList
    $items = Get-CondaPackagesHelper $false
    $installed = @{}

    foreach ($item in $items) {
        if (-not $outdatedOnly) {
            $null = $condaPackages.Add([PSCustomObject] @{Type='conda'; Package=$item.name; 'Installed'=$item.version})
        } else {
            $null = $installed.Add($item.name, $item.version)
        }
    }

    if ($outdatedOnly) {
        $items = Get-CondaPackagesHelper $outdatedOnly
        $items.PSObject.Properties | ForEach-Object {
            $name = $_.Name
            $archives = $_.Value

            if (($archives.Count -eq 0) -or (-not $installed.Contains($name))) {
                return
            }

            foreach ($archive in $archives) {
                [version] $version_installed = [version]::new()
                [version] $version_updated = [version]::new()

                $tryVIn = [version]::TryParse(($installed[$name] -replace '[^\d.]',''), [ref] $version_installed)
                $tryVUpd = [version]::TryParse(($archive.version -replace '[^\d.]',''), [ref] $version_updated)

                if (-not($tryVIn -and $tryVUpd -and ($version_installed -ge $version_updated))) {
                    $null = $condaPackages.Add([PSCustomObject] @{Type='conda'; Package=$name;
                        'Latest'=$archive.version;
                        'Installed'=$installed[$name]; })
                }

            }
        }
    }

    return ,$condaPackages
}

$actionCommands = @{
    pip=@{
        info          = { return (& (Get-CurrentInterpreter 'PythonExe') -m pip  show              $args 2>&1) };
        documentation = { $null = (Show-DocView $pkg).Show(); return ''    };
        files         = { return (& (Get-CurrentInterpreter 'PythonExe') -m pip  show    --files   $args 2>&1) };
        update        = { return (& (Get-CurrentInterpreter 'PythonExe') -m pip  install -U        $args 2>&1) };
        install       = { return (& (Get-CurrentInterpreter 'PythonExe') -m pip  install           $args 2>&1) };
        install_dry   = { return 'Not supported on pip'                        };
        install_nodep = { return (& (Get-CurrentInterpreter 'PythonExe') -m pip  install --no-deps $args 2>&1) };
        download      = { return (& (Get-CurrentInterpreter 'PythonExe') -m pip  download          $args 2>&1) };
        uninstall     = { return (& (Get-CurrentInterpreter 'PythonExe') -m pip  uninstall --yes   $args 2>&1) };
        reverseDependencies = {
            param($pkg)
            $di = Get-PipDistributionInfo
            $packages = $di | Get-Member -Type NoteProperty | Select-Object -ExpandProperty Name
            foreach ($package in $packages) {
                    $deps = $di."$package".deps | Get-Member -Type NoteProperty | Select-Object -ExpandProperty Name
                    if ($pkg -in $deps) {
                        "$package"
                    }
            }
        };
    };
    conda=@{
        info          = { return (& (Get-CurrentInterpreter 'CondaExe') list --prefix (Get-CurrentInterpreter 'Path') -v --json $args 2>&1) };
        documentation = { $null = (Show-DocView $pkg).Show(); return ''    };
        files         = {
            $path = "$(Get-CurrentInterpreter 'Path')\conda-meta"
            $query = "$args*.json"
            $file = Get-ChildItem -Path $path $query
            $json = Get-Content "$path\$($file.Name)" | ConvertFrom-Json
            return $json.files
        };
        update        = { return (& (Get-CurrentInterpreter 'CondaExe') update --prefix (Get-CurrentInterpreter 'Path') --yes -q $args 2>&1) };
        install       = { return (& (Get-CurrentInterpreter 'CondaExe') install (Get-PipsSetting 'CondaChannels' -AsArgs -First) --prefix (Get-CurrentInterpreter 'Path') --yes -q --no-shortcuts $args 2>&1) };
        install_dry   = { return (& (Get-CurrentInterpreter 'CondaExe') install (Get-PipsSetting 'CondaChannels' -AsArgs -First) --prefix (Get-CurrentInterpreter 'Path') --dry-run $args 2>&1) };
        install_nodep = { return (& (Get-CurrentInterpreter 'CondaExe') install (Get-PipsSetting 'CondaChannels' -AsArgs -First) --prefix (Get-CurrentInterpreter 'Path') --yes -q --no-shortcuts --no-deps --no-update-dependencies   $args 2>&1) };
        download      = { return 'Not supported on conda' };
        uninstall     = { return (& (Get-CurrentInterpreter 'CondaExe') uninstall --prefix (Get-CurrentInterpreter 'Path') --yes $args 2>&1) };
        reverseDependencies = { return (& (Get-CurrentInterpreter 'CondaExe') search --json --reverse-dependency $args 2>&1) | ConvertFrom-Json | Get-Member -Type NoteProperty | Select-Object -ExpandProperty Name };
    };
}
$actionCommands.wheel   = $actionCommands.pip
$actionCommands.sdist   = $actionCommands.pip
$actionCommands.builtin = $actionCommands.pip
$actionCommands.other   = $actionCommands.pip
$actionCommands.git     = $actionCommands.pip
$actionCommands.https   = $actionCommands.pip

Function Copy-AsRequirementsTxt($list) {
    $requirements = New-Object System.Text.StringBuilder
    foreach ($item in $list) {
        $null = $requirements.AppendLine("$($item.Package)==$($item.Installed)")
    }
    Set-Clipboard $requirements.ToString()
    WriteLog "Copied $($list.Count) items to clipboard."
}

$actionItemCount = 0
Function Make-PipActionItem($name, $code, $validator, $takesList = $false) {
    $action = New-Object psobject -Property @{Name=$name; TakesList=$takesList; Id=(++$Script:actionItemCount);}
    $action | Add-Member -MemberType ScriptMethod -Name ToString -Value { "$($this.Name) [F$($this.Id)]" } -Force
    $action | Add-Member -MemberType ScriptMethod -Name Execute  -Value $code -Force
    $action | Add-Member -MemberType ScriptMethod -Name Validate -Value $validator -Force
    return $action
}

Function Add-ComboBoxActions {
    $actionsModel = New-Object System.Collections.ArrayList
    $Add = { param($a) $null = $actionsModel.Add($a) }

    & $Add (Make-PipActionItem 'Show Info' `
        { param($pkg,$type); $actionCommands[$type].info.Invoke($pkg) } `
        { param($pkg,$out); $out -match $pkg } )

    & $Add (Make-PipActionItem 'Documentation' `
        { param($pkg,$type); $actionCommands[$type].documentation.Invoke($pkg) } `
        { param($pkg,$out); $out -match '.*' } )

    & $Add (Make-PipActionItem 'List Files' `
        {
            param($pkg,$type);
            if ($type -eq 'other') {
                Get-ChildItem -Recurse "$(Get-CurrentInterpreter "SitePackagesDir")\$pkg" | ForEach-Object { $_.FullName }
            } else {
                $actionCommands[$type].files.Invoke($pkg)
            }
        } `
        { param($pkg,$out); $out -match $pkg } )

    & $Add (Make-PipActionItem 'Update' `
        { param($pkg,$type); $actionCommands[$type].update.Invoke($pkg) } `
        { param($pkg,$out); $out -match "Successfully installed (?:[^\s]+\s+)*$pkg" } )

    & $Add (Make-PipActionItem 'Install (Dry Run)' `
        { param($pkg,$type); $actionCommands[$type].install_dry.Invoke($pkg) } `
        { param($pkg,$out); $out -match "Successfully installed (?:[^\s]+\s+)*$pkg" } )

    & $Add (Make-PipActionItem 'Install (No Deps)' `
        { param($pkg,$type); $actionCommands[$type].install_nodep.Invoke($pkg) } `
        { param($pkg,$out); $out -match "Successfully installed (?:[^\s]+\s+)*$pkg" } )

    & $Add (Make-PipActionItem 'Install' {
            param($pkg,$type,$version)
            $git_url = Validate-GitLink $pkg
            if ($git_url) {
                $pkg = $git_url
            }
            if ((-not [string]::IsNullOrEmpty($version)) -and ($type -notin @('git', 'wheel'))) {  # as git version is a timestamp; wheel version is in the file name
                $pkg = "$pkg==$version"
            }
            $actionCommands[$type].install.Invoke($pkg) } `
        { param($pkg,$out); ($out -match "Successfully installed (?:[^\s]+\s+)*$pkg") } )

    & $Add (Make-PipActionItem 'Download' `
        { param($pkg,$type); $actionCommands[$type].download.Invoke($pkg) } `
        { param($pkg,$out); $out -match 'Successfully downloaded ' } )

    & $Add (Make-PipActionItem 'Uninstall' `
        { param($pkg,$type); $actionCommands[$type].uninstall.Invoke($pkg) } `
        { param($pkg,$out); $out -match ('Successfully uninstalled ' + $pkg) } )

    & $Add (Make-PipActionItem 'As requirements.txt' `
        { param($list); Copy-AsRequirementsTxt($list) } `
        { param($pkg,$out); $out -match '.*' } `
        $true )  # Yes, take a whole list of packages

    $PIPTREE_LEGEND = "
Tree legend:
* = Extra package
x = Package doesn't exist in index
∞ = Dependency loop found
"

    & $Add (Make-PipActionItem 'Dependency tree' `
        { param($list); $Script:PIPTREE_LEGEND; Get-DependencyAsciiGraph $list; WriteLog "`n" }.GetNewClosure() `
        { param($pkg,$out); $out -match '.*' } )

    & $Add (Make-PipActionItem 'Reverse dependencies' `
        { param($pkg,$type); $actionCommands[$type].reverseDependencies.Invoke($pkg) } `
        { param($pkg,$out); $out -match '.*' } )

    $global:actionsModel = $actionsModel

    $actionListComboBox = New-Object System.Windows.Forms.ComboBox
    $actionListComboBox.DataSource = $actionsModel
    $actionListComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $Script:actionListComboBox = $actionListComboBox
    Add-TopWidget $actionListComboBox 1.25

    return $actionListComboBox
}


$trackDuplicateInterpreters = New-Object System.Collections.Generic.HashSet[String]


Function Create-EnvironmentVariablesDataTable {
    $environmentVariables = New-Object System.Data.DataTable
    foreach ($column in @('Enabled', 'Variable', 'Value')) {
        $type = if ($column -ne 'Enabled') { ([string]) } else { ([bool]) }
        $column = New-Object System.Data.DataColumn $column,$type
        $environmentVariables.Columns.Add($column)
    }
    return ,$environmentVariables
}

Function Get-InterpreterRecord($path, $items, $user = $false) {
    if ($trackDuplicateInterpreters.Contains($path)) {
        return
    }

    $python = Guess-EnvPath $path 'python' -Executable
    if (-not $python) {
        return
    }
    $versionString = & $python --version 2>&1
    $version = [regex]::Match($versionString, '\s+(\d+\.\d+)').Groups[1]
    $arch = (Test-is64Bit $python).FileType

    $action = New-Object psobject -Property @{
        Path                 = $path;
        Version              = "$version";
        Arch                 = $arch;
        Bits                 = @{"x64"="64"; "x86"="32";}[$arch];
        PythonExe            = $python;
        PipExe               = Guess-EnvPath $path 'pip' -Executable;
        CondaExe             = Guess-EnvPath $path 'conda' -Executable;
        VirtualenvExe        = Guess-EnvPath $path 'virtualenv' -Executable;
        VenvActivate         = Guess-EnvPath $path 'activate' -Executable;
        PipenvExe            = Guess-EnvPath $path 'pipenv' -Executable;
        RequirementsTxt      = Guess-EnvPath $path 'requirements.txt';
        Pipfile              = Guess-EnvPath $path 'Pipfile';
        PipfileLock          = Guess-EnvPath $path 'Pipfile.lock';
        SitePackagesDir      = Guess-EnvPath $path 'Lib\site-packages' -directory;
        User                 = $user;
        EnvironmentVariables = $null;
    }
    $action | Add-Member -MemberType ScriptMethod -Name ToString -Value {
        "{2} [{0}] {1}" -f $this.Arch, $this.PythonExe, $this.Version
    } -Force

    $null = $items.Add($action)
    $null = $trackDuplicateInterpreters.Add($path)

    if ($Global:interpretersComboBox) {
        $interpretersComboBox.DataSource = $null
        $interpretersComboBox.DataSource = $interpreters
    }
}

Function Find-Interpreters {
    $items = New-Object System.Collections.ArrayList

    $pythons_in_path = Get-Bin 'python' -All
    foreach ($path in $pythons_in_path) {
        Get-InterpreterRecord (Split-Path -Parent $path) $items
    }

    $pythons_in_system_drive_root = @(Get-ChildItem $env:SystemDrive\Python* | ForEach-Object { $_.FullName })
    foreach ($path in $pythons_in_system_drive_root) {
        Get-InterpreterRecord $path $items
    }

    $local_user_pythons = "$env:LOCALAPPDATA\Programs\Python"
    if (Exists-Directory $local_user_pythons) {
        foreach ($d in dir $local_user_pythons) {
            if ($d -is [System.IO.DirectoryInfo]) {
                Get-InterpreterRecord (${d}.FullName) $items
            }
        }
    }

    # search registry as defined here: https://www.python.org/dev/peps/pep-0514/
    if (Test-Path HKCU:Software\Python) {
        $reg_pythons_u = Get-ChildItem HKCU:Software\Python | ForEach-Object { Get-ChildItem HKCU:$_ } | ForEach-Object { (Get-ItemProperty -Path HKCU:$_\InstallPath) } | Where-Object { $_ } | ForEach-Object { $_.'(default)'  }
    }
    if (Test-Path HKLM:Software\Python) {
        $reg_pythons_m = Get-ChildItem HKLM:Software\Python | ForEach-Object { Get-ChildItem  HKLM:$_ } | ForEach-Object { (Get-ItemProperty -Path HKLM:$_\InstallPath).'(default)' }
    }
    if (Test-Path HKLM:Software\Wow6432Node\Python) {
        $reg_WoW_pythons = Get-ChildItem HKLM:Software\Wow6432Node\Python | ForEach-Object { Get-ChildItem  HKLM:$_ } | ForEach-Object { (Get-ItemProperty -Path HKLM:$_\InstallPath).'(default)' }
    }
    foreach ($d in @($reg_pythons_u; $reg_pythons_m; $reg_WoW_pythons)) {
        Get-InterpreterRecord ($d -replace '\\$','') $items
    }

    return ,$items  # keep comma to prevent conversion to an @() array
}


Function Add-ComboBoxInterpreters {
    $interpreters = Find-Interpreters
    [System.Collections.ArrayList] $interpreters = $interpreters | Sort-Object -Property @{Expression="Version"; Descending=$True}

    foreach ($interpreter in $Global:settings.envs) {
        if ($interpreter.User) {
            $interpreter | Add-Member -MemberType ScriptMethod -Name ToString -Value {
                "{2} [{0}] {1}" -f $this.Arch, $this.PythonExe, $this.Version
            } -Force
            [void]$interpreters.Add($interpreter)
        }
    }

    $Script:interpreters = $interpreters
    $interpretersComboBox = New-Object System.Windows.Forms.ComboBox
    $interpretersComboBox.DataSource = $interpreters
    $interpretersComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $Global:interpretersComboBox = $interpretersComboBox
    Add-TopWidget $interpretersComboBox 4
    return $interpretersComboBox
}

Function Add-CheckBox($text, $code) {
    $checkBox = New-Object System.Windows.Forms.CheckBox
    $checkBox.Text = $text
    $checkBox.Add_Click($code)
    Add-TopWidget $checkBox 0.75
    return ($checkBox)
}

Function Toggle-VirtualEnv ($state) {
    if (-not $Script:formLoaded) {
        return
    }

    Function Guess-EnvPath ($fileName) {
        $paths = @('.\env\Scripts\'; '.\Scripts\'; '.env\')
        foreach ($tryPath in $paths) {
            if (Test-Path ($tryPath + $fileName)) {
                return ($tryPath + $fileName)
            }
        }
        return $null
    }

    $pipEnvActivate = Guess-EnvPath 'activate.ps1'
    $pipEnvDeactivate = Guess-EnvPath 'deactivate.bat'

    if ($pipEnvActivate -eq $null -or $pipEnvDeactivate -eq $null) {
        WriteLog 'virtualenv not found. Run me from where "pip -m virtualenv env" command has been executed.'
        return
    }

    if ($state) {
        $env:_OLD_VIRTUAL_PROMPT = "$env:PROMPT"
        $env:_OLD_VIRTUAL_PYTHONHOME = "$env:PYTHONHOME"
        $env:_OLD_VIRTUAL_PATH = "$env:PATH"

        WriteLog ('Activating: ' + $pipEnvActivate)
        . $pipEnvActivate
    }
    else {
        WriteLog ('Deactivating: "' + $pipEnvDeactivate + '" and unsetting environment')

        &$pipEnvDeactivate

        $env:VIRTUAL_ENV = ''

        if ($env:_OLD_VIRTUAL_PROMPT) {
            $env:PROMPT = "$env:_OLD_VIRTUAL_PROMPT"
            $env:_OLD_VIRTUAL_PROMPT = ''
        }

        if ($env:_OLD_VIRTUAL_PYTHONHOME) {
            $env:PYTHONHOME = "$env:_OLD_VIRTUAL_PYTHONHOME"
            $env:_OLD_VIRTUAL_PYTHONHOME = ''
        }

        if ($env:_OLD_VIRTUAL_PATH) {
            $env:PATH = "$env:_OLD_VIRTUAL_PATH"
            $env:_OLD_VIRTUAL_PATH = ''
        }
    }

    #WriteLog "PROMPT=" $env:PROMPT
    #WriteLog "PYTHONHOME=" $env:PYTHONHOME
    #WriteLog "PATH=" $env:PATH
}

Function ConvertFrom-RegexGroupsToObject($groups) {
    $gitLinkInfo = New-Object PSCustomObject

    foreach ($group in $groups) {
        if (-not [string]::IsNullOrEmpty($group.Value) -and $group.Name -ne "0") {
            $gitLinkInfo | Add-Member -MemberType NoteProperty -Name "$($group.Name)" -Value "$($group.Value)"
        }
    }

    return $gitLinkInfo
}

Function global:Validate-GitLink ($url, [switch] $AsObject) {
    $r = [regex] '^(?<Prefix>\w+\+)?(?<Protocol>\w+)://(?<Host>[^/]+)/(?<User>[^/]+)/(?<Repo>[^/@#]+)(?:@(?<Hash>[^#]+))?(?:#.*)?$'
    $s = [regex] '^(?<Name>[^/]+)/(?<Repo>[^/]+)$'
    $f = [regex] '^(?<Prefix>\w+\+)?file:///(?<Path>.+)$'

    $m_file = $f.Match($url)
    $g_file = $m_file.Groups
    if ($g_file.Count -gt 1) {
        if (Exists-Directory "$($g_file['Path'])/.git") {
            if ($AsObject) {
                return ConvertFrom-RegexGroupsToObject $g_file
            } else {
                return "git+file:///$($g_file['Path'] -replace '\\','/')"
            }
        } else {
            return $null
        }
    }
    if (Exists-Directory "$url/.git") {
        if ($AsObject) {
            return [PSCustomObject]@{'Path' = $url;}
        } else {
            return "git+file:///$($url -replace '\\','/')"
        }
    }

    $m_short = $s.Match($url)
    $g_short = $m_short.Groups
    if ($g_short.Count -gt 1) {
        $url = "$github_url/$($g_short['Name'])/$($g_short['Repo'])"
    }

    $m = $r.Match($url)
    $g = $m.Groups

    if (($g['Prefix'].Value -ne 'git+') -and -not ($g['Protocol'].Value -in @('git', 'ssh', 'https'))) {
        return $null
    }

    $hash = if ($g['Hash'].Value) { "@$($g['Hash'].Value)" } else { [string]::Empty }

    if ($AsObject) {
        return ConvertFrom-RegexGroupsToObject $g
    }

    return "git+$($g['Protocol'])://$($g['Host'])/$($g['User'])/$($g['Repo'])$Hash#egg=$($g['Repo'])"
}

Function global:Set-PackageListEditable ($enable) {
    @('Package'; 'Installed'; 'Latest'; 'Type'; 'Status') | ForEach-Object {
        if ($global:dataModel.Columns.Contains($_)) {
            $global:dataModel.Columns[$_].ReadOnly = -not $enable
        }
    }
}

Function global:Prepare-PackageAutoCompletion {

    if ($global:autoCompleteIndex -ne $null) {
        return
    }

    Import-Module .\BK-tree\bktree
    $bktree = [BKTree]::new()
    $bktree.LoadArrays('known-packages-bktree.bin')

    $global:autoCompleteIndex = New-Object System.Windows.Forms.AutoCompleteStringCollection

    [int] $maxLength = 0
    foreach ($item in $bktree.index_id_to_name.Values) {
        [void] $autoCompleteIndex.Add($item)
    }
    # only needed bktree here to load package names
    Remove-Variable bktree -ErrorAction SilentlyContinue
    Remove-Module bktree -ErrorAction SilentlyContinue

    foreach ($plugin in $global:plugins) {
        [void] $autoCompleteIndex.AddRange($plugin.GetAllPackageNames())
    }
}

Function Generate-FormInstall {
    Prepare-PackageAutoCompletion

    $form = New-Object System.Windows.Forms.Form
    $form.KeyPreview = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Type package names and hit Enter after each. Ctrl-Enter to force.

source:name==version | github_user/project@tag | C:\git\repo@tag
"
    $label.Location = New-Object Drawing.Point 7,7
    $label.Size = New-Object Drawing.Point 360,40
    $form.Controls.Add($label)

    $cb = New-Object System.Windows.Forms.TextBox
    # $cb = New-Object System.Windows.Forms.ComboBox

    $form | Add-Member -MemberType NoteProperty -Name 'currentInstallAutoCompleteMode' -Value 'None'

    $autoCompleteIndex = $global:autoCompleteIndex
    $FuncGuessAutoCompleteMode = {
        $text = $cb.Text
        $n = $text.LastIndexOfAny('\/')
        $possibleDirectoryPath = $text.Substring(0, $n + 1)
        $pathExists = (Exists-Directory $possibleDirectoryPath)

        if ($text.Contains('==') -and -not $pathExists) {
            return [InstallAutoCompleteMode]::Version
        } elseif ($text.Contains('@') -and ($pathExists -or (Validate-GitLink ($text -replace '@.*$','')))) {
            return [InstallAutoCompleteMode]::GitTag
        } elseif ($pathExists) {
            if (Exists-File $possibleDirectoryPath -Mask '*.whl') {
                return [InstallAutoCompleteMode]::WheelFile
            } else {
                return [InstallAutoCompleteMode]::Directory
            }
        } else {
            return [InstallAutoCompleteMode]::Name
        }
    }.GetNewClosure()


    $FuncAfterDownload = ([EventHandler] {
        param($Sender, $EventArgs)
        # https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.control.invoke?view=netframework-4.7.2

        $cb = $Sender
        ($packageName, $completions_format, $releases) = (
            $EventArgs.PackageName,
            $EventArgs.CompletionsFormat,
            $EventArgs.Items)

        $autoCompletePackageVersion = [System.Windows.Forms.AutoCompleteStringCollection]::new()

        foreach ($release in $releases) {
            $entry = $completions_format -f @($packageName, $release)
            $null = $autoCompletePackageVersion.Add($entry)
        }

        $cb.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::None
        $cb.AutoCompleteCustomSource = $autoCompletePackageVersion
        $cb.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource
        $cb.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::Suggest

        $null = $SendMessage.Invoke($cb.Handle,  $WM_CHAR, 0x20, 0)
        $null = $SendMessage.Invoke($cb.Handle,  $WM_CHAR, $VK_BACKSPACE, 0)
    })

    $FuncSetAutoCompleteMode = [EventHandler] {
        param($Sender, $EventArgs)

        [InstallAutoCompleteMode] $mode = $EventArgs.Mode

        # Write-Host "ACTIVATING MODE !!! $mode"

        $cb = $Sender
        $FuncAfterDownload = $script:FuncAfterDownload
        $text = $cb.Text
        $n = $text.LastIndexOfAny('\/')
        $possibleDirectoryPath = $text.Substring(0, $n + 1)

        $cb.AutoCompleteCustomSource = $null
        $cb.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::None
        $cb.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::None

        switch ($mode)
        {
            ([InstallAutoCompleteMode]::Name) {
                # Write-Host '<###### 1>=' $cb.Handle
                $cb.AutoCompleteCustomSource = $global:autoCompleteIndex
                $cb.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::Suggest
                $cb.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource
                # Write-Host '<###### 2>=' $cb.Handle
            }

            ([InstallAutoCompleteMode]::Directory) {
                $cb.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
                $cb.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::FileSystemDirectories
            }

            ([InstallAutoCompleteMode]::Version) {
                $packageName = $text -replace '==.*',''
                $completions_format = '{0}=={1}'

                # if (-not $global:PyPiPackBeginageJsonCache.ContainsKey($packageName)) {
                #     $null = Download-PythonPackageDetails $packageName
                #     $releases = $global:PyPiPackageJsonCache[$packageName].'releases' | Get-Member -Type Properties | ForEach-Object { $_.Name }
                #     return ,$releases
                # }

                $jsonUrl = "https://pypi.python.org/pypi/$packageName/json"
                $null = Download-String $jsonUrl -ContinueWith {
                    param($json)

                    $info = ConvertFrom-Json -InputObject $json
                    $releases = @($info.'releases'.PSObject.Properties | Select-Object -ExpandProperty Name)
                    # $releases = Sort-Versions $releases

                    $EventArgs = MakeEvent @{
                        PackageName=$packageName;
                        CompletionsFormat=$completions_format;
                        Items=$releases;
                    }
                    $null = $cb.BeginInvoke($FuncAfterDownload, ($cb, $EventArgs))
                }.GetNewClosure()
            }

            ([InstallAutoCompleteMode]::GitTag) {
                $packageName = $text -replace '@.*',''
                $completions_format = "{0}@{1}"

                $gitLinkInfo = Validate-GitLink $packageName -AsObject
                $null = Get-GithubRepoTags $gitLinkInfo -ContinueWith {
                    param($releases)
                    $EventArgs = MakeEvent @{
                        PackageName=$packageName;
                        CompletionsFormat=$completions_format;
                        Items=$releases;
                    }
                    $null = $cb.BeginInvoke($FuncAfterDownload, ($cb, $EventArgs))
                }.GetNewClosure()
            }

            ([InstallAutoCompleteMode]::WheelFile) {
                # Write-Host 'PP='$possibleDirectoryPath
                $autoCompleteWheels = [System.Windows.Forms.AutoCompleteStringCollection]::new()
                $wheelFiles = Get-ChildItem -Path $possibleDirectoryPath -Filter '*.whl' -File -Depth 0
                foreach ($wheel in $wheelFiles) {
                    # Write-Host 'WHEEL=' $($wheel.Name)
                    [void] $autoCompleteWheels.Add("$possibleDirectoryPath$($wheel.Name)")
                }

                $cb.AutoCompleteCustomSource = $autoCompleteWheels
                $cb.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
                $cb.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource

                $null = $SendMessage.Invoke($cb.Handle,  $WM_CHAR, 0x20, 0)
                $null = $SendMessage.Invoke($cb.Handle,  $WM_CHAR, $VK_BACKSPACE, 0)
            }

            Default { throw "Wrong completion mode: $mode $($mode.GetType())" }
        }

    }.GetNewClosure()

    $FuncCleanupToolTip = {
        $hint = $global:hint
        if ($hint) {
            $hint.Dispose()
            $global:hint = $null
            return $true
        }
        return $false
    }.GetNewClosure()

    $FuncShowToolTip = {
        param($title, $text)
        $null = & $FuncCleanupToolTip
        $hint = New-Object System.Windows.Forms.ToolTip
        $global:hint = $hint
        $hint.IsBalloon = $true
        $hint.ToolTipTitle = $title
        $hint.ToolTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
        $hint.Show([string]::Empty, $cb, 0);
        $hint.Show($text, $cb, 0, $cb.Height);
    }.GetNewClosure()

    $FuncAddInstallSource = {
        # TODO: split this into two functions: detect type and add

        param($package, [bool] $force)

        $link = Validate-GitLink $package
        if ($link) {
            if (-not $force -and (Test-PackageInList $link) -ne -1) {
                & $FuncShowToolTip "$package" "Repository '$link' is already in the list"
                return $false
            }
            $row = $global:dataModel.NewRow()
            $row.Package = $link
            $row.Type = 'git'
            $row.Status = 'Pending'
            $row.Select = $true
            $global:dataModel.Rows.InsertAt($row, 0)
            return $true
        }

        if ($package -match '^https://') {
            $row = $global:dataModel.NewRow()
            $row.Package = $package
            $row.Type = 'https'
            $row.Status = 'Pending'
            $row.Select = $true
            $global:dataModel.Rows.InsertAt($row, 0)
            return $true
        }

        if (Exists-File $package -and $package.EndsWith('.whl')) {
            $row = $global:dataModel.NewRow()
            $row.Package = $package
            $row.Type = 'wheel'
            $row.Status = 'Pending'
            $row.Select = $true
            $global:dataModel.Rows.InsertAt($row, 0)
            return $true
        }

        $version = [string]::Empty
        $type = [string]::Empty
        $package_with_version = [regex] '^(?:(?<Type>[a-z_]+):)?(?<Name>[^=]+)(?:==(?<Version>[^=]+))?$'
        $pv_match = $package_with_version.Match($package)
        $pv_group = $pv_match.Groups
        if ($pv_group.Count -gt 1) {
            ($type, $package, $version) = ($pv_group['Type'].Value, $pv_group['Name'].Value, $pv_group['Version'].Value)
        }

        if (-not $force -and -not ($autoCompleteIndex.Contains($package))) {
            return $false
        }

        $nAlreadyInList = Test-PackageInList $package
        if ($nAlreadyInList -ne -1) {
            $oldRow = $global:dataModel.Rows[$nAlreadyInList]

            if (($global:dataModel.Columns.Contains('Latest')) -and
                (-not ([string]::IsNullOrEmpty($oldRow.Latest)))) {
                $oldVersion = $oldRow.Latest
            } elseif (($global:dataModel.Columns.Contains('Installed')) -and
                (-not ([string]::IsNullOrEmpty($oldRow.Installed)))) {
                $oldVersion = $oldRow.Installed
            } else {
                $oldVersion = [string]::Empty
            }

            #Write-Host "old='$oldVersion', new='$version'"

            $IsDifferentVersion = $version -ne $oldVersion
            $IsDifferentType = ((-not [string]::IsNullOrEmpty($type)) -and ($type -ne $oldRow.Type)) -or ((-not [string]::IsNullOrEmpty($oldRow.Type)) -and [string]::IsNullOrEmpty($type))
        } else {
            $IsDifferentVersion = $true
            $IsDifferentType = $true
        }

        if ($nAlreadyInList -ne -1) {
            if ((-not $IsDifferentVersion) -and (-not $IsDifferentType)) {
                & $FuncShowToolTip "$package" "Package '$package' is already in the list"
                return $false
            } else {
                $row = $oldRow
                Set-PackageListEditable $true
            }
        } else {
            $row = $global:dataModel.NewRow()
        }

        $row.Select = $true
        $row.Package = $package

        # opinionated behavior but seems to be conspicuously right
        if ((-not [string]::IsNullOrWhiteSpace($version)) -or [string]::IsNullOrWhiteSpace($type)) {
            $row.Installed = $version
        }

        if ($IsDifferentType) {
            if (-not [string]::IsNullOrWhiteSpace($type)) {
                if (-not ($type -in $global:packageTypes)) {
                    & $FuncShowToolTip "${type}:$package" "Wrong source type '$type'.`n`nSupported types: $($global:packageTypes -join ', ')"
                    return $false
                }
                $row.Type = $type
            } elseif ([string]::IsNullOrWhiteSpace($row.Type)) {
                $row.Type = 'pip'
            }
        }

        $row.Status = 'Pending'

        if ($nAlreadyInList -eq -1) {
            $global:dataModel.Rows.InsertAt($row, 0)
        } else {
            Set-PackageListEditable $false
        }

        return $true
    }

    $cbRef = [ref] $cb

    $global:FuncRPCSpellCheck_Callback = New-RunspacedDelegate ([Action[System.Threading.Tasks.Task[string]]] {
        param([System.Threading.Tasks.Task[string]] $task)

        if (-not [string]::IsNullOrEmpty($task.Result)) {
            $result = $task.Result | ConvertFrom-Json
        } else {
            return
        }

        # $candidates = $result.Candidates
        # if ($candidates.Count -le 10) {
        #     $candidatesToolTipText = "$($candidates -join "`n")"
        # } else {
        #     $candidatesToolTipText = "$(($candidates | Select-Object -First 10) -join "`n")`n...`n`nfull list in the log"
        # }

        # & $FuncShowToolTip "$text" "Packages with similar names found in the index.`n`nDid you mean:`n`n$candidatesToolTipText"


        $null = WriteLog "Suggestions for '$($result.Request)': $($result.Candidates)" -UpdateLastLine -Background ([System.Drawing.Color]::LightSkyBlue)

        $global:SuggestionsWorking = $false

        $cb = $cbRef.Value

        if (($cb -ne $null) -and ($cb.Text -ne $result.Request)) {
            $null = & $FuncRPCSpellCheck $cb.Text 1
        }
    }.GetNewClosure())

    $cb.add_TextChanged({
        param($cb)
        $text = $cb.Text

        $guessedCompletionMode = & $FuncGuessAutoCompleteMode
        # Write-Host 'MODE GUESS=' $guessedCompletionMode
        if ($guessedCompletionMode -ne $form.currentInstallAutoCompleteMode) {
            # Write-Host 'MODE WAS=' $form.currentInstallAutoCompleteMode 'CHANGE TO=' $guessedCompletionMode

            $form.currentInstallAutoCompleteMode = $guessedCompletionMode

            $EventArgs = MakeEvent @{
                Mode=$guessedCompletionMode;
            }
            $null = $cb.BeginInvoke($FuncSetAutoCompleteMode, ($cb, $EventArgs))
        }

        if ($guessedCompletionMode -eq [InstallAutoCompleteMode]::Name) {
            $null = & $FuncRPCSpellCheck $text 1
        }

    }.GetNewClosure());

    $cb.Location = New-Object Drawing.Point 7,60
    $cb.Size = New-Object Drawing.Point 330,32
    $form.Controls.Add($cb)

    $form.add_KeyDown({
        # [void] (& $FuncCleanupToolTip)

        if ($_.KeyCode -eq 'Escape') {
            if (($cb.Text.Length -gt 0) -or (& $FuncCleanupToolTip)) {
                $cb.Text = [string]::Empty
            } else {
                $form.Close()
            }
        }

        if ($_.KeyCode -eq 'Enter') {
            $text = $cb.Text.Trim()

            if ([string]::IsNullOrWhiteSpace($text)) {
                return
            }

            if (-not (Test-KeyPress -Keys ShiftKey)) {
                $force = Test-KeyPress -Keys ControlKey
                $okay = & $FuncAddInstallSource $text.ToLower() $force
                if ($okay) {
                   $cb.Text = [string]::Empty
                   return
                }
            } else {
                WriteLog -UpdateLastLine `
                    "Searching for similar package names. Wait for ~5 sec..."
                $null = & $FuncRPCSpellCheck $cb.Text 2
            }
        }
    }.GetNewClosure())

    $form.Add_Closing({
        $cbRef.Value = $null
    }.GetNewClosure())

    $install = New-Object System.Windows.Forms.Button
    $install.Text = "Install"
    $install.Location = New-Object Drawing.Point 140,90
    $install.Size = New-Object Drawing.Point 70,24
    $install.add_Click({
        $okay = & $FuncAddInstallSource $cb.Text
        if ($okay) {
            $cb.Text = [string]::Empty
        }
        if ($cb.Text -ne [string]::Empty) {
            return
        }
        Select-PipAction 'Install'
        Execute-PipAction
        $form.Close()
    }.GetNewClosure())
    $form.Controls.Add($install)

    $form.Size = New-Object Drawing.Point 365,160
    $form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Hide
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.Text = 'Install packages | [Shift+Enter] fuzzy name search'
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false
    $form.Icon = $Script:form.Icon

    $null = $form.ShowDialog()
}

Function global:Request-UserString($message, $title, $default, $completionItems = $null, [ref] $ControlKeysState) {
    $Form                            = New-Object system.Windows.Forms.Form
    $Form.ClientSize                 = '421,247'
    $Form.text                       = $title
    $Form.TopMost                    = $false
    $Form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Hide
    $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $Form.MinimizeBox = $false
    $Form.MaximizeBox = $false
    $Form.Icon = $Script:form.Icon

    $TextBox1                        = New-Object system.Windows.Forms.TextBox
    $TextBox1.multiline              = $false
    $TextBox1.width                  = 379
    $TextBox1.height                 = 20
    $TextBox1.location               = New-Object System.Drawing.Point(21,165)

    $ButtonCancel                    = New-Object system.Windows.Forms.Button
    $ButtonCancel.text               = "Cancel"
    $ButtonCancel.width              = 60
    $ButtonCancel.height             = 30
    $ButtonCancel.location           = New-Object System.Drawing.Point(140,200)
    $ButtonCancel.DialogResult       = [System.Windows.Forms.DialogResult]::Cancel

    $ButtonOK                        = New-Object system.Windows.Forms.Button
    $ButtonOK.text                   = "OK"
    $ButtonOK.width                  = 60
    $ButtonOK.height                 = 30
    $ButtonOK.location               = New-Object System.Drawing.Point(225,200)
    $ButtonOK.DialogResult           = [System.Windows.Forms.DialogResult]::OK

    $Label1                          = New-Object system.Windows.Forms.Label
    $Label1.text                     = $message
    $Label1.AutoSize                 = $false
    $Label1.width                    = 371
    $Label1.height                   = 123
    $Label1.location                 = New-Object System.Drawing.Point(25,29)

    $Form.controls.AddRange( @($TextBox1, $ButtonCancel, $ButtonOK, $Label1) )
    $Form.AcceptButton = $ButtonOK
    $Form.CancelButton = $ButtonCancel

    $TextBox1.Text = $default
    if ($completionItems) {
        $autoCompleteStrings = New-Object System.Windows.Forms.AutoCompleteStringCollection
        $autoCompleteStrings.AddRange($completionItems)
        $TextBox1.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::Suggest
        $TextBox1.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource
        $TextBox1.AutoCompleteCustomSource = $autoCompleteStrings
    }

    $result = $Form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {

        if ($ControlKeysState) {
            $ControlKeysState.Value = @{
                ShiftKey=(Test-KeyPress -Keys ShiftKey);
                ControlKey=(Test-KeyPress -Keys ControlKey);
            }
        }

        return $TextBox1.Text
    } else {
        return $null
    }
}

Function Generate-FormSearch {
    $pluginNames = ($global:plugins | ForEach-Object { $_.GetPluginName() }) -join ', '
    $message = "Enter keywords to search with PyPi, Conda, Github and plugins: $pluginNames`n`nChecked items will be kept in the search list"
    $title = "pip, conda, github search"
    $default = ""
    $input = Request-UserString $message $title $default
    if (-not $input) {
        return
    }

    WriteLog "Searching for $input"
    WriteLog 'Double click or [Ctrl+Enter] a table row to open a package home page in browser'
    $stats = Get-SearchResults $input
    WriteLog "Found $($stats.Total) packages: $($stats.PipCount) pip, $($stats.CondaCount) conda, $($stats.GithubCount) github, $($stats.PluginCount) from plugins. Total $($global:dataModel.Rows.Count) packages in list."
    WriteLog
}

Function Init-PackageGridViewProperties() {
    $dataGridView.MultiSelect = $false
    $dataGridView.SelectionMode = [System.Windows.Forms.SelectionMode]::One
    $dataGridView.ColumnHeadersVisible = $true
    $dataGridView.RowHeadersVisible = $false
    $dataGridView.ReadOnly = $false
    $dataGridView.AllowUserToResizeRows = $false
    $dataGridView.AllowUserToResizeColumns = $false
    $dataGridView.VirtualMode = $false
    $dataGridView.AutoGenerateColumns = $true
    $dataGridView.AllowUserToAddRows = $false
    $dataGridView.AllowUserToDeleteRows = $false
    $dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
}

Function Init-PackageUpdateColumns($dataTable) {
    $dataTable.Columns.Clear()
    foreach ($c in $header) {
        if ($c -eq "Select") {
            $column = New-Object System.Data.DataColumn $c,([bool])
        } else {
            $column = New-Object System.Data.DataColumn $c,([string])
            $column.ReadOnly = $true
        }
        $dataTable.Columns.Add($column)
    }
}

Function Init-PackageSearchColumns($dataTable) {
    $dataTable.Columns.Clear()
    foreach ($c in $search_columns) {
        if ($c -eq "Select") {
            $column = New-Object System.Data.DataColumn $c,([bool])
        } else {
            $column = New-Object System.Data.DataColumn $c,([string])
            $column.ReadOnly = $true
        }
        $dataTable.Columns.Add($column)
    }
}

Function global:Highlight-PythonPackages {
    if (-not $outdatedOnly) {
        $dataGridView.BeginInit()
        foreach ($row in $dataGridView.Rows) {
            if ($row.DataBoundItem.Row.Type -eq 'builtin') {
                $row.DefaultCellStyle.BackColor = [Drawing.Color]::LightGreen
            } elseif ($row.DataBoundItem.Row.Type -eq 'other') {
                $row.DefaultCellStyle.BackColor = [Drawing.Color]::LightPink
            }
        }
        $dataGridView.EndInit()
    }
}

Function global:Open-LinkInBrowser($url) {
    if ((-not [string]::IsNullOrWhiteSpace($url)) -and ($url -match '^https?://')) {
        $url = [System.Uri]::EscapeUriString($url)
        Start-Process -FilePath $url
    }
}

$Global:jobCounter = 0
$Global:jobTimer = New-Object System.Windows.Forms.Timer
$Global:jobTimer.Interval = 250
$Global:jobTimer.add_Tick({ $null | Out-Null })  # this hack forces processing of two different event loops: window & PS object events
[int] $Global:jobSemaphore = 0
Function Run-SubProcessWithCallback($code, $callback, $params) {
    $Global:jobCounter++
    $n = $Global:jobCounter

    if ($Global:jobSemaphore -eq 0) {
        $Global:jobTimer.Start()
        # Write-Host 'run timer'
        }
    $Global:jobSemaphore++;

    Register-EngineEvent -SourceIdentifier "Custom.RaisedEvent$n" -Action $callback

    $codeString = $code.ToString()

    $job = Start-Job {
            param($codeString, $params, $n)
            $code = [scriptblock]::Create($codeString)
            $result = $code.Invoke($params)
            Register-EngineEvent "Custom.RaisedEvent$n" -Forward
            New-Event "Custom.RaisedEvent$n" -EventArguments @{Result=$result; Id=$n; Params=$params}
    } -Name "Job$n" -ArgumentList $codeString, $params, $n

    $null = Register-ObjectEvent $job -EventName StateChanged -SourceIdentifier "JobEnd$n" -MessageData @{Id=$n} `
        -Action {
            if($sender.State -eq 'Completed')  {
                $n = $event.MessageData.Id
                Unregister-Event -SourceIdentifier "JobEnd$n"
                Unregister-Event -SourceIdentifier "Custom.RaisedEvent$n"
                Get-Job -Name "Job$n" | Wait-Job | Remove-Job
                Get-Job -Name "JobEnd$n" | Wait-Job | Remove-Job
                Get-Job -Name "Custom.RaisedEvent$n" | Wait-Job | Remove-Job

                $Global:jobSemaphore--;
                if ($Global:jobSemaphore -eq 0) {
                    $Global:jobTimer.Stop()
                    # Write-Host 'stop timer'
                }
                # Write-Host "*** Cleaned up $n sem=$Global:jobSemaphore"
            }
        }
}

$Global:PyPiPackageJsonCache = New-Object 'System.Collections.Generic.Dictionary[string,PSCustomObject]'

Function global:Format-PythonPackageToolTip($info) {
    $name = "$($info.info.name)`n`n"
    $tt = "Summary: $($info.info.summary)`nAuthor: $($info.info.author)`nRelease: $($info.info.version)`n"
    $ti = "License: $($info.info.license)`nHome Page: $($info.info.home_page)`n"
    $lr = ($info.releases."$($info.info.version)")[0]  # The latest release

    $downloadStats = $info.releases |
        ForEach-Object { $_.PSObject.Properties.Value.downloads } |
        Measure-Object -Sum

    $tr = "Release Uploaded: $($lr.upload_time -replace 'T',' ')`n"
    $stats_rd = "Release Downloads: $($lr.downloads)`n"
    $stats_td = "Total Downloads: $($downloadStats.Sum.ToString("#,##0"))`n"
    $stats_tr = "Total Releases: $($downloadStats.Count)`n"
    $deps = "`nRequires Python: $($info.info.requires_python)`nRequires libraries:`n`t$($info.info.requires_dist -join "`n`t")`n"
    $tags = "`n$($info.info.classifiers -join "`n")"
    return "${name}${tt}${ti}${tr}${stats_rd}${stats_td}${stats_tr}${deps}${tags}"
}

Function global:Download-PythonPackageDetails ($packageName) {
    $pypi_json_url = 'https://pypi.python.org/pypi/{0}/json'
    $json = Download-String($pypi_json_url -f $packageName)
    if (-not [string]::IsNullOrEmpty($json)) {
        $info = $json | ConvertFrom-Json
        if ($info -ne $null) {
               $null = $Global:PyPiPackageJsonCache.Add($packageName, $info)
        }
    }
}

Function global:Update-PythonPackageDetails {
    $_ = $args[0]
    $viewRow = $dataGridView.Rows[$_.RowIndex]
    $rowItem = $viewRow.DataBoundItem
    $cells = $dataGridView.Rows[$_.RowIndex].Cells
    $packageName = $rowItem.Row.Package

    if ((-not [String]::IsNullOrEmpty($cells['Package'].ToolTipText)) -or (-not($rowItem.Row.Type -in @('pip', 'wheel', 'sdist')))) {
        return
    }

    if ($Global:PyPiPackageJsonCache.ContainsKey($packageName)) {
        $info = $Global:PyPiPackageJsonCache[$packageName]
        $cells['Package'].ToolTipText = Format-PythonPackageToolTip $info
        return
    }

    if (-not (Test-KeyPress -Keys ShiftKey)) {
        return
    }

    $global:dataModel.Columns['Status'].ReadOnly = $false
    $cells['Status'].Value = 'Fetching...'
    $global:dataModel.Columns['Status'].ReadOnly = $true
    $viewRow.DefaultCellStyle.BackColor = [Drawing.Color]::Gray

    Run-SubProcessWithCallback ({
        # Worker: Separate process
        param($params)
        [Void][Reflection.Assembly]::LoadWithPartialName("System.Net")
        [Void][Reflection.Assembly]::LoadWithPartialName("System.Net.WebClient")
        $code = [scriptblock]::Create($params.FuncNetWorkaround)
        [void] $code.Invoke( @() )
        $packageName = $params.PackageName
        $pypi_json_url = 'https://pypi.python.org/pypi/{0}/json'
        $jsonUrl = [String]::Format($pypi_json_url, $packageName)
        $webClient = New-Object System.Net.WebClient
        $json = $webClient.DownloadString($jsonUrl)
        $info = $json | ConvertFrom-Json
        return @($info)
    }) ({
        # Callback: Access over $Global in here
        $message = $event.SourceArgs
        $okay = $message.Result -ne $null
        if ($okay) {
            $Global:PyPiPackageJsonCache.Add($message.Params.PackageName, $message.Result)
        }
        $viewRow = $Global:dataGridView.Rows[$message.Params.RowIndex]
        $viewRow.DefaultCellStyle.BackColor = [Drawing.Color]::Empty
        $row = $viewRow.DataBoundItem.Row
        $global:dataModel.Columns['Status'].ReadOnly = $false
        $row.Status = if ($okay) { 'OK' } else { 'Failed' }
        $global:dataModel.Columns['Status'].ReadOnly = $true
    }) @{PackageName=$packageName; RowIndex=$_.RowIndex; FuncNetWorkaround=$FuncSetWebClientWorkaround.ToString();}
}

Function global:Request-FolderPathFromUser($text = [string]::Empty) {
    $selectFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $selectFolderDialog.Description = $text
    $null = $selectFolderDialog.ShowDialog()
    $path = $selectFolderDialog.SelectedPath
    $path = Get-ExistingPathOrNull $path
    return $path
}

Function global:Set-ActiveInterpreterWithPath($path) {
    for ($i = 0; $i -lt $Script:interpreters.Count; $i++) {
        if ($Script:interpreters[$i].Path -eq $path) {
            $Global:interpretersComboBox.SelectedIndex = $i
            WriteLog "Switching to env '$path'"
            break
        }
    }
}

Function global:Run-Elevated ($scriptblock, $argsList) {
    <#
        .PARAMETER argList
        Arguments for the Script Block. Must not contain single quotes, can contain spaces
    #>
    Start-Process -Verb RunAs -FilePath powershell -WindowStyle Hidden -Wait -ArgumentList `
        "-Command Invoke-Command {$scriptBlock} -ArgumentList $($argsList -replace '^|$','''' -join ',')"
}

Function Add-CreateEnvButtonMenu {
    $tools = @(
        @{
            MenuText = 'with virtualenv';
            Code = {
                param($path)
                $output = & (Get-CurrentInterpreter 'VirtualenvExe') --python="$(Get-CurrentInterpreter 'PythonExe')" $path 2>&1
                return $output
            };
            IsAccessible = { [bool] (Get-CurrentInterpreter 'VirtualenvExe' -Executable) };
        };
        @{
            MenuText = 'with pipenv';
            Code = {
                param($path)
                $env:PIPENV_VENV_IN_PROJECT = 1
                Set-Location -Path $path
                $output = & (Get-CurrentInterpreter 'PipenvExe') --python "$(Get-CurrentInterpreter 'Version')" install 2>&1
                return $output
            };
            IsAccessible = { [bool] (Get-CurrentInterpreter 'PipenvExe' -Executable) };
        };
        @{
            MenuText = 'with conda';
            Code = {
                param($path)
                $version = Get-CurrentInterpreter 'Version'
                $output = & (Get-CurrentInterpreter 'CondaExe') create -y -q --prefix $path python=$version 2>&1
                return $output
            };
            IsAccessible = { [bool] (Get-CurrentInterpreter 'CondaExe' -Executable) };
        };
        @{
            NoTargetPath = $true;
            MenuText = '(tool required) Install virtualenv';
            Code = {
                param($path)
                $output = & (Get-CurrentInterpreter 'PythonExe') -m pip install virtualenv 2>&1
                return $output
            };
            IsAccessible = { -not [bool] (Get-CurrentInterpreter 'VirtualenvExe') };
        };
        @{
            NoTargetPath = $true;
            MenuText = '(tool required) Install pipenv';
            Code = {
                param($path)
                $output = & (Get-CurrentInterpreter 'PythonExe') -m pip install pipenv 2>&1
                return $output
            };
            IsAccessible = { -not [bool] (Get-CurrentInterpreter 'PipenvExe') };
        };
        @{
            NoTargetPath = $true;
            MenuText = '(tool required) Install conda';
            Code = {
                param($path)
                # menuinst, cytoolz are required by conda to run
                $menuinst = Validate-GitLink "https://github.com/ContinuumIO/menuinst@1.4.8"
                $output_0 = & (Get-CurrentInterpreter 'PythonExe') -m pip install $menuinst 2>&1
                $output_1 = & (Get-CurrentInterpreter 'PythonExe') -m pip install cytoolz conda 2>&1
                $CondaExe = Guess-EnvPath (Get-CurrentInterpreter 'Path') 'conda' -Executable

                # conda needs a little caress to run together with pip
                $path = (Get-CurrentInterpreter 'Path')
                $stub = "$path\Lib\site-packages\conda\cli\pip_warning.py"
                $main = "$path\Lib\site-packages\conda\cli\main.py"
                Move-Item $stub "${stub}_"
                Copy-Item $main $stub
                # New-Item -Path $stub -ItemType SymbolicLink -Value $main

                $record = Get-CurrentInterpreter
                $record.CondaExe = $CondaExe
                return @($output_0, $output_1) | ForEach-Object { $_ }
            };
            IsAccessible = { -not [bool] (Get-CurrentInterpreter 'CondaExe' -Executable) };
        };
        @{
            Persistent = $true;
            MenuText = 'Custom... (TODO)';
            Code = {  };
        }
    )

    $FuncUpdateInterpreters = {
        param($path)
        Get-InterpreterRecord $path $interpreters -user $true
        Set-ActiveInterpreterWithPath $path

        $ruleName = "pip env $path"
        $pythonExe = (Get-CurrentInterpreter 'PythonExe')
        $firewallUserResponse = ([System.Windows.Forms.MessageBox]::Show(
            "Create a firewall rule for the new environment?`n`nRule name: '$ruleName'`nPath: '$pythonExe'`nAllow outgoing connections`n`n" +
            "You can edit the rule by running wf.msc",
            "Configure firewall", [System.Windows.Forms.MessageBoxButtons]::YesNo))
        if ($firewallUserResponse -eq 'Yes') {
            Run-Elevated ({
                param($path, $exe)
                New-NetFirewallRule `
                    -DisplayName "$path" `
                    -Program "$exe" `
                    -Direction Outbound `
                    -Action Allow
            }) @($ruleName, $pythonExe)

            $rule = Get-NetFirewallRule -DisplayName "$ruleName" -ErrorAction SilentlyContinue
            if ($rule) {
                WriteLog "Firewall rule '$ruleName' was successfully created."
            } else {
                WriteLog "Error while creating firewall rule '$ruleName'."
            }
        }
    }

    $FuncGetPythonInfo = {
        return (Get-CurrentInterpreter 'Version')
    }

    $menuclick = {
        param($tool)

        if (-not $tool.NoTargetPath) {
            $path = Request-FolderPathFromUser `
                "New python environment with active version $($FuncGetPythonInfo.Invoke()) will be created"
            if ($path -eq $null) { return }
        }
        WriteLog "$($tool.MenuText), please wait..."
        $output = $tool.Code.Invoke( @($path) )
        WriteLog ($output -join "`n")

        [void] $FuncUpdateInterpreters.Invoke($path)
    }.GetNewClosure()

    $buttonEnvCreate = Add-ButtonMenu 'env: Create' $tools $menuclick
    return $buttonEnvCreate
}

Function Add-EnvToolButtonMenu {
    $menu = @(
        @{
            Persistent = $true;
            MenuText = 'Python REPL';
            Code = { Start-Process -FilePath (Get-CurrentInterpreter 'PythonExe') -WorkingDirectory (Get-CurrentInterpreter 'Path') };
        };
        @{
            MenuText = 'Shell with Virtualenv Activated';
            Code = { Start-Process -FilePath cmd.exe -WorkingDirectory (Get-CurrentInterpreter 'Path') -ArgumentList "/K $(Get-CurrentInterpreter 'VenvActivate')" };
            IsAccessible = { (Get-CurrentInterpreter 'VenvActivate') };
        };
        @{
            Persistent = $true;
            MenuText = 'Open IDLE'
            Code = { Start-Process -FilePath (Get-CurrentInterpreter 'PythonExe') -WorkingDirectory (Get-CurrentInterpreter 'Path') -ArgumentList '-m idlelib.idle' -WindowStyle Hidden };
        };
        @{
            MenuText = 'pipenv shell'
            Code = { Start-Process -FilePath (Get-Bin 'pipenv.exe') -WorkingDirectory (Get-CurrentInterpreter 'Path') -ArgumentList 'shell' };
            IsAccessible = { [bool] (Get-Bin 'pipenv.exe') -and [bool] (Get-CurrentInterpreter 'Pipfile') };
        };
        @{
            Persistent = $true;
            MenuText = 'Environment variables...';
            Code = {
                $interpreter = Get-CurrentInterpreter
                if ($interpreter.User) {
                    Generate-FormEnvironmentVariables $interpreter
                } else {
                    $path = $interpreter.Path
                    WriteLog "WARNING: Editing global variables for global interpreter $path"
                    & rundll32 sysdm.cpl,EditEnvironmentVariables
                }
            };
        };
        @{
            Persistent = $true;
            MenuText = 'Open containing directory'
            Code = { Start-Process -FilePath 'explorer.exe' -ArgumentList "$(Get-CurrentInterpreter 'Path')" };
        };
        @{
            Persistent = $false;
            MenuText = 'Remove environment...'
            Code = {
                Delete-CurrentInterpreter
            };
            IsAccessible = { [bool] (Get-CurrentInterpreter 'User') };
        };
    )

    $menuclick = {
        param($item)
        $output = $item.Code.Invoke()
    }

    $buttonEnvTools = Add-ButtonMenu 'env: Tools' $menu $menuclick
    return $buttonEnvTools
}

Function global:Get-PyDocTopics() {
    $pythonExe = Get-CurrentInterpreter 'PythonExe'
    if ([string]::IsNullOrEmpty($pythonExe)) {
        WriteLog 'No python executable found.'
        return
    }
    $pyCode = "import pydoc; print('\n'.join(pydoc.Helper.topics.keys()))"
    $output = & $pythonExe -c $pyCode 2>&1
    if ($output) {
        return $output.Split("`n", [System.StringSplitOptions]::RemoveEmptyEntries)
    } else {
        return
    }
}

Function global:Get-PyDocApropos($request) {
    $pythonExe = Get-CurrentInterpreter 'PythonExe'
    if ([string]::IsNullOrEmpty($pythonExe)) {
        WriteLog 'No python executable found.'
        return
    }
    $request = $request -replace '''',''
    $pyCode = "import pydoc; pydoc.apropos('$request')"
    $output = & $pythonExe -c $pyCode
    if ($output) {
        return $output.Split("`n", [System.StringSplitOptions]::RemoveEmptyEntries)
    } else {
        return
    }
}

Function Add-ToolsButtonMenu {
    $menu = @(
        @{
            Persistent = $true;
            MenuText = 'View PyDoc for...';
            Code = {
                $message = "Enter requset in format: package.subpackage.Name

or symbol like 'print' or '%'

or topic like 'FORMATTING'

or keyword like 'elif'
"
                $title = "View PyDoc"
                $default = ""
                $topics = Get-PyDocTopics
                $input = Request-UserString $message $title $default $topics
                if (-not $input) {
                    return
                }
                (Show-DocView $input).Show()
            };
        };
        @{
            Persistent = $true;
            MenuText = 'Search through PyDocs (apropos) ...'
            Code = {
                $message = "Enter keywords to search

All module docs will be scanned.

Some packages may generate garbage or show windows, don't panic.
"
                $title = "PyDoc apropos"
                $default = ""
                $input = Request-UserString $message $title $default
                if (-not $input) {
                    return
                }
                WriteLog "Searching apropos for $input"
                $apropos = Get-PyDocApropos $input
                if ($apropos -and $apropos.Count -gt 0) {
                    WriteLog "Found $($apropos.Count) topics"
                    $docView = Show-DocView -SetContent ($apropos -join "`n") -Highlight $Script:pyRegexNameChain -NoDefaultHighlighting
                    $docView.Show()
                } else {
                    WriteLog 'Nothing found.'
                }
            };
        };
        @{
            Persistent=$true;
            MenuText = 'Set as default Python';
            Code = {
                $version = Get-CurrentInterpreter 'Version'
                $arch = Get-CurrentInterpreter 'Arch'
                $bits = Get-CurrentInterpreter 'Bits'
                $fileLines = New-Object System.Collections.ArrayList
                $fileLines.Add("[defaults]")
                $fileLines.Add("python=$version -$bits")
                $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
                [System.IO.File]::WriteAllLines(
                    "$($env:LOCALAPPDATA)\py.ini",
                    $fileLines,
                    $Utf8NoBomEncoding)
                WriteLog "$bits bit Python $version - is now default, use py.exe to launch"
            };
        };
        @{
            Persistent=$true;
            MenuText = 'Edit conda channels';
            Code = {
                $channels = $global:settings.condaChannels -join ' '
                $list = Request-UserString "Enter conda channels, separated with space.`n`nLeave empty to use defaults.`n`nPopular channels: conda-forge anaconda defaults bioconda`n`nThe first channel is for installing, the others are for searching only." 'Edit conda channels' $channels
                if ($list -eq $null) {
                    return
                }
                if (-not [string]::IsNullOrWhiteSpace($list)) {
                    $global:settings."condaChannels" = $list -split '\s+'
                } else {
                    $global:settings."condaChannels" = $null
                }
            };
        };
        @{
            Persistent=$true;
            IsAccessible = { $global:APP_MODE -eq [AppMode]::Idle };
            MenuText = 'Show pip cache info';
            Code = {
                $paths = [System.Collections.Generic.List[string]]::new()

                foreach ($type in @('pip', 'wheel')) {
                    $getCacheFolderScript_b10 = "from pip.utils.appdirs import user_cache_dir; print(user_cache_dir('$type'))"
                    $getCacheFolderScript_a10 = "from pip._internal.utils.appdirs import user_cache_dir; print(user_cache_dir('$type'))"
                    $cacheFolder = & (Get-CurrentInterpreter 'PythonExe') -c $getCacheFolderScript_b10
                    if ([string]::IsNullOrWhiteSpace($cacheFolder)) {
                        $cacheFolder = & (Get-CurrentInterpreter 'PythonExe') -c $getCacheFolderScript_a10
                    }
                    if ([string]::IsNullOrWhiteSpace($cacheFolder)) {
                        WriteLog "Could not determine $type cache location."
                        continue
                    }
                    [void] $paths.Add($cacheFolder)
                }

                foreach ($plugin in $global:plugins) {
                    [void] $paths.Add($plugin.GetCachePath())
                }

                foreach ($cacheFolder in $paths) {
                    $stats = Get-ChildItem -Recurse $cacheFolder | Measure-Object -Property Length -Sum
                    if ($stats) {
                        WriteLog ("`nCache at {0}`nFiles: {1}`nSize: {2} MB" -f $cacheFolder,$stats.Count,[math]::Round($stats.Sum / 1048576, 2))
                    }
                }
            };
        };
        @{
            Persistent=$true;
            MenuText = 'Clear log';
            Code = {
                ClearLog
            };
        };
        @{
            Persistent=$true;
            IsAccessible = { $global:APP_MODE -eq [AppMode]::Idle };
            MenuText = 'Clear package list';
            Code = {
                Clear-Rows
            };
        };
        @{
            Persistent=$true;
            MenuText = "Report a bug";
            Code = {
                $title = [System.Web.HttpUtility]::UrlEncode("Something went wrong")
                $powerShellInfo = $PSVersionTable | Format-Table -HideTableHeaders -AutoSize | Out-String
                $pipsPath = $PSScriptRoot
                $pipsGitPath = [IO.Path]::combine($pipsPath, '.git')
                $pipsHash = & (Get-Bin 'git') --git-dir=$pipsGitPath --work-tree=$pipsPath rev-parse --short HEAD 2>&1
                $hostInfo = "``````$powerShellInfo```````n.NET $FRAMEWORK_VERSION`npips $pipsHash`n`nDescribe your issue here"
                $issue = [System.Web.HttpUtility]::UrlPathEncode($hostInfo)
                Open-LinkInBrowser "https://github.com/ptytb/pips/issues/new?title=$title&body=$hostInfo"
            };
        };
    )

    $menuArray = [System.Collections.ArrayList]::new()
    [void] $menuArray.AddRange($menu)
    if ($global:plugins.Count -gt 0) {
        [void] $menuArray.Add(@{ 'IsSeparator'=$true; })
    }
    foreach ($plugin in $global:plugins) {
        [void] $menuArray.AddRange($plugin.GetToolMenuCommands())
    }
    $menu = $menuArray

    $menuclick = {
        param($item)
        $output = $item.Code.Invoke()
    }

    $toolsButton = Add-ButtonMenu 'Tools' $menu $menuclick
    return $toolsButton
}

Function global:Show-CurrentPackageInBrowser() {
    $view_row = $dataGridView.CurrentRow
    if ($view_row) {
        $row = $view_row.DataBoundItem.Row
        $packageName = $row.Package
        $urlName = [System.Web.HttpUtility]::UrlEncode($packageName)

        foreach ($plugin in $global:plugins) {
            $packageHomepageFromPlugin = $plugin.GetPackageHomepage($packageName, $row.Type)
            if ($packageHomepageFromPlugin) {
                break
            }
        }

        if ($packageHomepageFromPlugin) {
            Open-LinkInBrowser "$packageHomepageFromPlugin"
            return
        } elseif ($row.Type -eq 'conda') {
            $url = $anaconda_url
        } elseif ($row.Type -eq 'git') {
            $gitLinkInfo = Validate-GitLink $packageName -AsObject
            $url = "$github_url/"
            $urlName = "$($gitLinkInfo.User)/$($gitLinkInfo.Repo)"
        } else {
            $url = $pypi_url
        }
        Open-LinkInBrowser "${url}${urlName}"
    }
}

Function Generate-Form {
    $form = New-Object Windows.Forms.Form
    $form.Text = "pips - python package browser"
    $form.Size = New-Object Drawing.Point 1125, 840
    $form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Show
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $form.StartPosition = 'CenterScreen'
    $form.Topmost = $false
    $form.KeyPreview = $true
    $form.Icon = Convert-Base64ToICO $iconBase64_Snakes
    $Script:form = $form

    $null = Add-Buttons

    $actionListComboBox = Add-ComboBoxActions

    $group = New-Object System.Windows.Forms.Panel
    $group.Location = New-Object System.Drawing.Point 502,2
    $group.Size = New-Object System.Drawing.Size 226,28
    $group.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $form.Controls.Add($group)

    $Script:isolatedCheckBox = Add-CheckBox 'isolated' { Toggle-VirtualEnv $Script:isolatedCheckBox.Checked }
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.SetToolTip($isolatedCheckBox, "Ignore environmental variables, user configuration and global packages.`n`n--isolated`n--local")

    $global:WIDGET_GROUP_INSTALL_BUTTONS = @(
        Add-Button "Search..." { Generate-FormSearch } ;
        Add-Button "Install..." { Generate-FormInstall }
    )
    $null = Add-ToolsButtonMenu

    $null = NewLine-TopLayout

    $null = Add-Label "Filter results:"

    $global:inputFilter = Add-Input {  # TextChanged Handler here
        param($input)

        $selectedRow = $null

        if ($dataGridView.CurrentRow) {
            # Keep selection while filter is being changed
            $selectedRow = $dataGridView.CurrentRow.DataBoundItem.Row
        }

        $searchText = $input.Text -replace "'","''"
        $searchText = $searchText -replace '\[|\]',''

        Function Create-SearchSubQuery($column, $searchText, $junction) {
            return "$column LIKE '%" + ( ($searchText -split '\s+' | where { -not [String]::IsNullOrEmpty($_) }) -join  "%' $junction $column LIKE '%" ) + "%'"
        }

        switch ($searchMethodComboBox.Text) {
            'Whole Phrase' {
                $subQueryPackage     = "Package LIKE '%{0}%'" -f $searchText
                $subQueryDescription = "Description LIKE '%{0}%'" -f $searchText
            }
            'Exact Match' {
                $subQueryPackage     = "Package LIKE '{0}'" -f $searchText
                $subQueryDescription = "Description LIKE '{0}'" -f $searchText
            }
            'Any Word' {
                $subQueryPackage     = Create-SearchSubQuery 'Package'     $searchText 'OR'
                $subQueryDescription = Create-SearchSubQuery 'Description' $searchText 'OR'
            }
            'All Words' {
                $subQueryPackage     = Create-SearchSubQuery 'Package'     $searchText 'AND'
                $subQueryDescription = Create-SearchSubQuery 'Description' $searchText 'AND'
            }
        }
        $isInstallMode = $global:dataModel.Columns.Contains('Description')
        if ($isInstallMode) {
            $query = "($subQueryPackage) OR ($subQueryDescription)"
        } else {
            $query = "$subQueryPackage"
        }

        #Write-Host $query

        if ($searchText.Length -gt 0) {
            $global:dataModel.DefaultView.RowFilter = $query
        } else {
            $global:dataModel.DefaultView.RowFilter = $null
        }

        if ($selectedRow) {
            Set-SelectedRow $selectedRow
        }

        Highlight-PythonPackages
    }.GetNewClosure()

    $global:inputFilter = $global:inputFilter
    $toolTipFilter = New-Object System.Windows.Forms.ToolTip
    $toolTipFilter.SetToolTip($global:inputFilter, "Esc to clear")

    $searchMethodComboBox = New-Object System.Windows.Forms.ComboBox
    $searchMethods = New-Object System.Collections.ArrayList
    $null = $searchMethods.Add('Whole Phrase')
    $null = $searchMethods.Add('Any Word')
    $null = $searchMethods.Add('All Words')
    $null = $searchMethods.Add('Exact Match')
    $searchMethodComboBox.DataSource = $searchMethods
    $searchMethodComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $global:searchMethodComboBox = $searchMethodComboBox
    Add-TopWidget($searchMethodComboBox)
    $searchMethodComboBox.add_SelectionChangeCommitted({
        $flt = $global:inputFilter
        $t = $flt.Text
        $flt.Text = [String]::Empty
        $flt.Text = $t
        })

    $labelInterp   = Add-Label "Active Interpreter:"
    $toolTipInterp = New-Object System.Windows.Forms.ToolTip
    $toolTipInterp.SetToolTip($labelInterp, "Ctrl+C to copy selected path")

    $interpretersComboBox = Add-ComboBoxInterpreters
    $Global:interpretersComboBox = $interpretersComboBox
    $buttonEnvOpen = Add-Button "env: Open..." {
        $path = Request-FolderPathFromUser ("Choose a folder with python environment, created by either Virtualenv or pipenv`n`n" +
            "Typically it contains dirs: Include, Lib, Scripts")
        if ($path) {
            $oldCount = $interpreters.Count
            Get-InterpreterRecord $path $interpreters -user $true
            if (($interpreters.Count -gt $oldCount) -or ($trackDuplicateInterpreters.Contains($path))) {
                if ($interpreters.Count -gt $oldCount) {
                    WriteLog "Added virtual environment location: $path"
                }
                Set-ActiveInterpreterWithPath $path
            } else {
                WriteLog "No python found in $path"
            }
        }
    }

    $buttonEnvCreate = Add-CreateEnvButtonMenu
    $buttonEnvTools = Add-EnvToolButtonMenu

    $global:WIDGET_GROUP_ENV_BUTTONS = @($buttonEnvOpen, $buttonEnvCreate, $buttonEnvTools)

    $interpreters = $Script:interpreters
    $trackDuplicateInterpreters = $Script:trackDuplicateInterpreters
    $form.add_KeyDown({
        if ($Global:interpretersComboBox.Focused) {
            if (($_.KeyCode -eq 'C') -and $_.Control) {
                $python_exe = Get-CurrentInterpreter 'PythonExe'
                Set-Clipboard $python_exe
                WriteLog "Copied to clipboard: $python_exe"
            }
            if ($_.KeyCode -eq 'Delete') {
                Delete-CurrentInterpreter
            }
        }
    }.GetNewClosure())

    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $Script:dataGridView = $dataGridView
    $Global:dataGridView = $dataGridView

    $dataGridView.Location = New-Object Drawing.Point 7,($Script:lastWidgetTop + $Script:widgetLineHeight)
    $dataGridView.Size = New-Object Drawing.Point 800,450
    $dataGridView.ShowCellToolTips = $false
    $dataGridView.Add_Sorted({ Highlight-PythonPackages })

    $dataGridToolTip = New-Object System.Windows.Forms.ToolTip

    $dataGridView.Add_CellMouseEnter({
        if (($_.RowIndex -gt -1)) {
            Update-PythonPackageDetails $_
            $text = $dataGridView.Rows[$_.RowIndex].Cells['Package'].ToolTipText
            $dataGridToolTip.RemoveAll()
            if (-not [string]::IsNullOrEmpty($text)) {
                $dataGridToolTip.InitialDelay = 50
                $dataGridToolTip.ReshowDelay = 10
                $dataGridToolTip.AutoPopDelay = [Int16]::MaxValue
                $dataGridToolTip.ShowAlways = $true
                $dataGridToolTip.SetToolTip($dataGridView, $text)
            }
        }
    }.GetNewClosure())

    $form.add_KeyDown({
        if ($_.KeyCode -eq 'Escape') {
            $dataGridToolTip.Hide($dataGridView)
            $dataGridToolTip.InitialDelay = [Int16]::MaxValue

            if ($global:inputFilter.Focused) {
                if ([string]::IsNullOrEmpty($global:inputFilter.Text)) {
                    $dataGridView.Focus()
                } else {
                    $global:inputFilter.Text = [String]::Empty
                }
            } else {
                $global:inputFilter.Focus()
            }
        }
        if ($_.KeyCode -eq 'Return') {
            if ($global:inputFilter.Focused) {
                $_.Handled = $true
                $dataGridView.Focus()
            }
        }
        if ($global:inputFilter.Focused -and ($_.KeyCode -in @('Up', 'Down'))) {
            $Script:dataGridView.Focus()
            $_.Handled = $false
        }
    }.GetNewClosure())

    Init-PackageGridViewProperties

    $dataModel = New-Object System.Data.DataTable
    $dataGridView.DataSource = $dataModel
    $global:dataModel = $dataModel
    Init-PackageUpdateColumns $dataModel

    $form.Controls.Add($dataGridView)

    $logView = $global:RichTextBox_t::new()
    $logView.Location = New-Object Drawing.Point 7,520
    $logView.Size = New-Object Drawing.Point 800,270
    $logView.ReadOnly = $true
    $logView.Multiline = $true
    $logView.Font = New-Object System.Drawing.Font("Consolas", 11)
    $form.Controls.Add($logView)

    $logView.add_HandleCreated({
        param($Sender)
        $null = [SearchDialogHook]::new($Sender)
        $null = [TextBoxNavigationHook]::new($Sender)

        $logView = $Sender

        $FuncWriteLog = {
            param(
                [object[]] $Lines,
                [bool] $UpdateLastLine,
                [bool] $NoNewline,
                [MaybeColor] $Background,
                [MaybeColor] $Foreground
            )

            $null = $SendMessage.Invoke($logView.Handle, $WM_SETREDRAW, 0, 0)
            $eventMask = $SendMessage.Invoke($logView.Handle, $EM_SETEVENTMASK, 0, 0)

            if ($UpdateLastLine) {
                $text = ($Lines -join ' ') -replace "`r|`n",''

                $cr = $lastLineCharIndex = $logView.Find("`r", 0, $logView.TextLength,
                    [System.Windows.Forms.RichTextBoxFinds]::Reverse)

                $lf = $lastLineCharIndex = $logView.Find("`n", 0, $logView.TextLength,
                    [System.Windows.Forms.RichTextBoxFinds]::Reverse)

                $lastLineCharIndex = [Math]::Max($cr, $lf)

                if ($lastLineCharIndex -eq -1) {
                    $lastLineCharIndex = 0
                } else {
                    ++$lastLineCharIndex
                }

                $lastLineLength = $logView.TextLength - $lastLineCharIndex
                $logView.Select($lastLineCharIndex, $lastLineLength);
                $logView.SelectedText = $text

                $logFrom = $lastLineCharIndex
                $logTo = $logView.TextLength
            } else {
                $logFrom = $logView.TextLength

                foreach ($obj in $Lines) {
                    $logView.AppendText("$obj")
                }

                $logTo = $logView.TextLength

                if (-not $NoNewline) {
                    $logView.AppendText("`n")
                }
            }

            if (($Background -ne $null) -or ($Foreground -ne $null)) {
                $logView.Select($logFrom, $logTo)
                if ($Background -ne $null) {
                    $logView.SelectionBackColor = $Background
                }
                if ($Foreground -ne $null) {
                    $logView.SelectionColor = $Foreground
                }
            }

            $textLength = $logView.TextLength
            $logView.Select($textLength, $textLength)

            $null = $SendMessage.Invoke($logView.Handle, $WM_SETREDRAW, 1, 0)
            $null = $SendMessage.Invoke($logView.Handle, $EM_SETEVENTMASK, 0, $eventMask)
            $null = $SendMessage.Invoke($logView.Handle, $WM_VSCROLL, $SB_PAGEBOTTOM, 0)
        }.GetNewClosure()

        $WritePipLogDelegate = [EventHandler] {
            param($Sender, $EventArgs)
            $Arguments = $EventArgs.Arguments
            $null = & $FuncWriteLog @Arguments
        }.GetNewClosure()

        ${function:global:WriteLog} = {
            [CmdletBinding()]
            param(
                [Parameter(ValueFromRemainingArguments=$true, Position=1)]
                [AllowEmptyCollection()]
                [AllowNull()]
                [object[]] $Lines = @(),
                [switch] $UpdateLastLine,
                [switch] $NoNewline,
                [Parameter(Mandatory=$false)] [AllowNull()] $Background = $null,
                [Parameter(Mandatory=$false)] [AllowNull()] $Foreground = $null
            )

            $arguments = @{
                Lines=$Lines;
                UpdateLastLine=([bool] $PSBoundParameters['UpdateLastLine']);
                NoNewline=([bool] $PSBoundParameters['NoNewline']);
                Background=(EnsureColor $Background);
                Foreground=(EnsureColor $Foreground);
            }

            if ($logView.InvokeRequired) {
                $EventArgs = MakeEvent @{
                    Arguments=$arguments
                }
                $null = $logView.Invoke($WritePipLogDelegate, ($logView, $EventArgs))
            } else {
                $null = & $FuncWriteLog @arguments
            }
        }.GetNewClosure()

        ${function:global:ClearLog} = {
            $logView.Clear()
        }.GetNewClosure()

        ${function:global:GetLogLength} = {
            return $logView.TextLength
        }.GetNewClosure()

        foreach ($arguments in $global:_WritePipLogBacklog) {
            WriteLog @arguments
        }
        Remove-Variable -Scope Global _WritePipLogBacklog
    })

    $logView.Add_LinkClicked({
        param($Sender, $EventArgs)
        Open-LinkInBrowser $EventArgs.LinkText
    })

    $FuncHighlightLogFragment = {
        if ($global:dataModel.Rows.Count -eq 0) {
            return
        }

        $viewRow = $Script:dataGridView.CurrentRow
        if (-not $viewRow) {
            return
        }
        $row = $viewRow.DataBoundItem.Row

        $logView.SelectAll()
        $logView.SelectionBackColor = $logView.BackColor

        if (Get-Member -inputobject $row -name "LogTo" -Membertype Properties) {
            $logView.Select($row.LogFrom, $row.LogTo)
            $logView.SelectionBackColor = [Drawing.Color]::Yellow
            $logView.ScrollToCaret()
        }
    }.GetNewClosure()

    $dataGridView.Add_CellMouseClick({ & $FuncHighlightLogFragment }.GetNewClosure())
    $dataGridView.Add_SelectionChanged({ & $FuncHighlightLogFragment }.GetNewClosure())

    $dataGridView.Add_CellMouseDoubleClick({
            if (($_.RowIndex -gt -1) -and ($_.ColumnIndex -gt 0)) {
                Show-CurrentPackageInBrowser
            }
        }.GetNewClosure())
    $form.Add_Load({ $Script:formLoaded = $true })

    $lastWidgetTop = $Script:lastWidgetTop

    $FuncResizeForm = {
        $dataGridView.Width = $form.ClientSize.Width - 15
        $dataGridView.Height = $form.ClientSize.Height / 2
        $logView.Top = $dataGridView.Bottom + 15
        $logView.Width = $form.ClientSize.Width - 15
        $logView.Height = $form.ClientSize.Height - $dataGridView.Bottom - $lastWidgetTop
    }.GetNewClosure()

    $null = & $FuncResizeForm
    $form.Add_Resize({ & $FuncResizeForm }.GetNewClosure())
    $form.Add_Shown({
        WriteLog "`n" 'Hold Shift and hover the rows to fetch the detailed package info form PyPi'
        $form.BringToFront()
    })

    $FunctionalKeys = (1..12 | ForEach-Object { "F$_" })

    $form.add_KeyUp({
        if ($_.KeyCode -in $FunctionalKeys) {  # Handle F1..F12 functional keys
            $n = [int]("$($_.KeyCode)" -replace 'F','') - 1
            if ($n -lt $Script:actionListComboBox.DataSource.Count) {
                $Script:actionListComboBox.SelectedIndex = $n
            }
        }

        if ($Script:dataGridView.Focused -and $Script:dataGridView.RowCount -gt 0) {
            if ($_.KeyCode -eq 'Home') {
                Set-SelectedNRow 0
                $_.Handled = $true
            }

            if ($_.KeyCode -eq 'End') {
                Set-SelectedNRow ($Script:dataGridView.RowCount - 1)
                $_.Handled = $true
            }

            if ($_.KeyCode -eq 'Return' -and $_.Control) {
                Show-CurrentPackageInBrowser
                $_.Handled = $true
                return
            }

            if ($_.KeyCode -eq 'Return' -and $_.Shift) {
                $_.Handled = $true
                Execute-PipAction
                return
            }

            if ($_.KeyCode -in @('Space', 'Return')) {
                $oldSelect = $Script:dataGridView.CurrentRow.DataBoundItem.Row.Select
                $Script:dataGridView.CurrentRow.DataBoundItem.Row.Select = -not $oldSelect
                $_.Handled = $true
            }
        }
    }.GetNewClosure())

    return ,$form
}

Function Generate-FormEnvironmentVariables($interpreterRecord) {
    $Form                            = New-Object system.Windows.Forms.Form
    $Form.ClientSize                 = '659,653'
    $Form.Text                       = "Environment Variables"
    $Form.TopMost                    = $false
    $Form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Hide
    $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $Form.MinimizeBox = $false
    $Form.MaximizeBox = $false
    $Form.KeyPreview = $true
    $Form.Icon = $Script:form.Icon

    $VariablesGroup                  = New-Object system.Windows.Forms.Groupbox
    $VariablesGroup.height           = 322
    $VariablesGroup.width            = 624
    $VariablesGroup.text             = "Environment Variables for $($interpreterRecord."Path")"
    $VariablesGroup.location         = New-Object System.Drawing.Point(16,24)

    $New                             = New-Object system.Windows.Forms.Button
    $New.text                        = "New"
    $New.width                       = 94
    $New.height                      = 25
    $New.location                    = New-Object System.Drawing.Point(409,280)

    $Delete                          = New-Object system.Windows.Forms.Button
    $Delete.text                     = "Delete"
    $Delete.width                    = 95
    $Delete.height                   = 25
    $Delete.location                 = New-Object System.Drawing.Point(515,280)

    $DataGridView1                   = New-Object System.Windows.Forms.DataGridView
    $DataGridView1.width             = 592
    $DataGridView1.height            = 238
    $DataGridView1.location          = New-Object System.Drawing.Point(16,20)
    $DataGridView1.ColumnHeadersVisible = $true
    $DataGridView1.RowHeadersVisible = $false
    $DataGridView1.AutoGenerateColumns = $true
    $DataGridView1.AllowUserToAddRows = $false
    $DataGridView1.AllowUserToDeleteRows = $true
    $DataGridView1.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $DataGridView1.DataSource        = Create-EnvironmentVariablesDataTable

    $StartupScriptGroup              = New-Object system.Windows.Forms.Groupbox
    $StartupScriptGroup.height       = 222
    $StartupScriptGroup.width        = 624
    $StartupScriptGroup.text         = "Activate Script additional commands at startup (Windows batch)"
    $StartupScriptGroup.location     = New-Object System.Drawing.Point(16,360)

    $TextBox1                        = New-Object System.Windows.Forms.TextBox
    $TextBox1.Multiline              = $true
    $TextBox1.WordWrap               = $false
    $TextBox1.AcceptsReturn          = $true
    $TextBox1.AcceptsTab             = $true
    $TextBox1.ScrollBars             = 'Both'
    $TextBox1.width                  = 592
    $TextBox1.height                 = 185
    $TextBox1.Text                   = 'REM Run these commands when virtualenv being activated'
    $TextBox1.location               = New-Object System.Drawing.Point(16,20)

    $ButtonCancel                    = New-Object system.Windows.Forms.Button
    $ButtonCancel.text               = "Cancel"
    $ButtonCancel.width              = 96
    $ButtonCancel.height             = 26
    $ButtonCancel.location           = New-Object System.Drawing.Point(531,612)
    $ButtonCancel.DialogResult       = [System.Windows.Forms.DialogResult]::Cancel

    $ButtonOk                         = New-Object system.Windows.Forms.Button
    $ButtonOk.text                    = "OK"
    $ButtonOk.width                   = 94
    $ButtonOk.height                  = 26
    $ButtonOk.location                = New-Object System.Drawing.Point(425,612)
    $ButtonOK.DialogResult            = [System.Windows.Forms.DialogResult]::OK

    $ButtonWrite                      = New-Object system.Windows.Forms.Button
    $ButtonWrite.text                 = "Save to Activate Script"
    $ButtonWrite.width                = 151
    $ButtonWrite.height               = 26
    $ButtonWrite.location             = New-Object System.Drawing.Point(19,612)

    $ButtonRestore                    = New-Object system.Windows.Forms.Button
    $ButtonRestore.text               = "Restore Activate Script"
    $ButtonRestore.width              = 151
    $ButtonRestore.height             = 26
    $ButtonRestore.location           = New-Object System.Drawing.Point(180,612)

    $Form.controls.AddRange(@($VariablesGroup,$StartupScriptGroup,$ButtonOk,
                             $ButtonCancel,$ButtonWrite,$ButtonRestore))
    $VariablesGroup.Controls.AddRange(@($DataGridView1,$New,$Delete))
    $StartupScriptGroup.Controls.AddRange(@($TextBox1))
    $Form.AcceptButton = $ButtonOK
    $Form.CancelButton = $ButtonCancel

    # Convert form's DataTable back to '$interpreterRecord.EnvironmentVariables'
    $FuncDataTableToEnvVars = {
        $newVars = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt $DataGridView1.Rows.Count; $i++) {
            $row = $DataGridView1.Rows[$i].DataBoundItem.Row
            $value = if([string]::IsNullOrEmpty($row.Value)) { $null } else { $row.Value }
            [void]$newVars.Add((New-Object PSObject -Property @{
                'Variable'=$row.Variable;
                'Value'=$value;
                'Enabled'=[bool]$row.Enabled;
            }))
        }
        $interpreterRecord | Add-Member -Force NoteProperty `
            -Name 'EnvironmentVariables' `
            -Value $newVars

        [string[]] $StartupScript = @($TextBox1.Lines)
        $interpreterRecord | Add-Member -Force NoteProperty `
            -Name 'StartupScript' `
            -Value $StartupScript
    }.GetNewClosure()

    $FuncUpdateActivateScript = {
        param($WritePipsSection = $true)

        & $FuncDataTableToEnvVars

        $lines = Get-Content $interpreterRecord.VenvActivate
        $newLines = New-Object System.Collections.ArrayList
        $skipRegion = $false
        $sectionBegin = "REM PIPS VARS BEGIN"
        $sectionEnd   = "REM PIPS VARS END"

        $newLines.Add("@echo off")

        if ($WritePipsSection) {
            $newLines.Add($sectionBegin)
            $interpreterRecord.EnvironmentVariables |
                Where-Object { $_.Enabled } |
                ForEach-Object { $newLines.Add("set `"$($_.Variable)=$($_.Value)`"") }
            foreach ($line in $interpreterRecord.StartupScript) {
                $newLines.Add($line)
            }
            $newLines.Add($sectionEnd)
        }

        foreach ($line in $lines) {
            if ($line -match $sectionBegin) {
                $skipRegion = $true
            } elseif ($line -match $sectionEnd) {
                $skipRegion = $false
            } elseif (-not $skipRegion -and -not ($line -match '^@echo off')) {
                $newLines.Add($line)
            }
        }

        $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
        [System.IO.File]::WriteAllLines(
            $interpreterRecord.VenvActivate,
            $newLines,
            $Utf8NoBomEncoding)
    }.GetNewClosure()

    $New.Add_Click({
            $row  = $DataGridView1.DataSource.NewRow()
            $row.Enabled = $true
            $DataGridView1.DataSource.Rows.Add($row)
        }.GetNewClosure())

    $Delete.Add_Click({
            if ($DataGridView1.CurrentRow -ne $null) {
                $DataGridView1.CurrentRow.DataBoundItem.Delete()
            }
        }.GetNewClosure())

    $ButtonWrite.Add_Click({
        & $FuncUpdateActivateScript
        }.GetNewClosure())

    $ButtonRestore.Add_Click({
        & $FuncUpdateActivateScript -WritePipsSection $false
        }.GetNewClosure())

    # $Form.add_KeyDown({
    #         if ($_.KeyCode -eq 'Return' -and $TextBox1.Focused) {
    #             $_.SuppressKeyPress = $true
    #             $_.Handled = $true
    #         }
    #     }.GetNewClosure())

    # Convert '$interpreterRecord.EnvironmentVariables' object to DataTable
    $interpreterRecord.EnvironmentVariables |
        Sort-Object -Property Variable |
        ForEach-Object {
            $row  = $DataGridView1.DataSource.NewRow()
            $row.Variable = $_.Variable
            $row.Value    = $_.Value
            $row.Enabled  = $_.Enabled
            $DataGridView1.DataSource.Rows.Add($row)
        }

    if ($interpreterRecord.PSObject.Properties.Name -contains 'StartupScript') {
        $TextBox1.Text = $interpreterRecord.StartupScript -join [System.Environment]::NewLine
    }

    $result = $Form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        & $FuncDataTableToEnvVars
    }
}


$fontBold      = New-Object System.Drawing.Font("Consolas", 11, [System.Drawing.FontStyle]::Bold)
$fontUnderline = New-Object System.Drawing.Font("Consolas", 11, [System.Drawing.FontStyle]::Underline)

$ro_none       = [System.Text.RegularExpressions.RegexOptions]::None
$ro_compiled   = [System.Text.RegularExpressions.RegexOptions]::Compiled
$ro_multiline  = [System.Text.RegularExpressions.RegexOptions]::Multiline
$ro_singleline = [System.Text.RegularExpressions.RegexOptions]::Singleline
$ro_ecma       = [System.Text.RegularExpressions.RegexOptions]::ECMAScript
$regexDefaultOptions = $ro_compiled

$pydocSections = @('NAME', 'DESCRIPTION', 'PACKAGE CONTENTS', 'CLASSES', 'FUNCTIONS', 'DATA', 'VERSION', 'AUTHOR', 'FILE', 'MODULE DOCS', 'SUBMODULES', 'CREDITS', 'DATE', 'MODULE REFERENCE')
$pydocKeywords = @('False', 'None', 'True', 'and', 'as', 'assert', 'break', 'class', 'continue', 'def', 'del', 'elif', 'else', 'except', 'finally', 'for', 'from', 'global', 'if', 'import', 'in', 'is', 'lambda', 'nonlocal', 'not', 'or', 'pass', 'raise', 'return', 'try', 'while', 'with', 'yield')
$pydocSpecial  = @('self', 'async', 'await', 'dict', 'list', 'tuple', 'set', 'float', 'int', 'bool', 'str', 'type', 'map', 'filter', 'bytes', 'bytearray', 'frozenset')
$pydocSpecialU = @('builtins', 'main', 'all', 'abs', 'add', 'and', 'call', 'class', 'cmp', 'coerce', 'complex', 'contains', 'del', 'delattr', 'delete', 'delitem', 'delslice', 'dict', 'div', 'divmod', 'eq', 'float', 'floordiv', 'ge', 'get', 'getattr', 'getattribute', 'getitem', 'getslice', 'gt', 'hash', 'hex', 'iadd', 'iand', 'idiv', 'ifloordiv', 'ilshift', 'imod', 'imul', 'index', 'init', 'instancecheck', 'int', 'invert', 'ior', 'ipow', 'irshift', 'isub', 'iter', 'itruediv', 'ixor', 'le', 'len', 'long', 'lshift', 'lt', 'metaclass', 'mod', 'mro', 'mul', 'ne', 'neg', 'new', 'nonzero', 'oct', 'or', 'pos', 'pow', 'radd', 'rand', 'rcmp', 'rdiv', 'rdivmod', 'repr', 'reversed', 'rfloordiv', 'rlshift', 'rmod', 'rmul', 'ror', 'rpow', 'rrshift', 'rshift', 'rsub', 'rtruediv', 'rxor', 'set', 'setattr', 'setitem', 'setslice', 'slots', 'str', 'sub', 'subclasscheck', 'truediv', 'unicode', 'weakref', 'xor')

Function Compile-Regex($pattern, $options = $regexDefaultOptions) {
    return New-Object System.Text.RegularExpressions.Regex($pattern, $options)
}

$pyRegexStrQuote = Compile-Regex @"
('(?:[^\n'\\]|\\.)*?'|"(?:[^\n"\\]|\\.)*?")
"@

$pyRegexNumber    = Compile-Regex '([-+]?[0-9]*\.?[0-9]+(?:[eE][-+]?[0-9]+)?)'
$pyRegexSection   = Compile-Regex ("^\s*(" + (($pydocSections | foreach { "${_}"     }) -join '|') + ")\s*$") ($ro_compiled + $ro_multiline)
$pyRegexKeyword   = Compile-Regex ("\b(" + (($pydocKeywords | foreach { "${_}"     }) -join '|') + ")\b")
$pyRegexSpecial   = Compile-Regex ("\b(" + (($pydocSpecial  | foreach { "${_}"     }) -join '|') + ")\b")
$pyRegexSpecialU  = Compile-Regex ("\b(" + (($pydocSpecialU | foreach { "__${_}__" }) -join '|') + ")\b")
$pyRegexPEP       = Compile-Regex '\b(PEP[ -]?(?:\d+))'
$pyRegexSubPkgs   = Compile-Regex 'PACKAGE CONTENTS\n((?:\s+[^\n]+\n)+)'
$pyRegexNameChain = Compile-Regex '^\s*(\w+(?:\.\w+)*)' ($ro_compiled + $ro_multiline)

$DrawingColor = New-Object Drawing.Color  # Workaround for PowerShell, which parses script classes before loading types

class DocView {

    [System.Windows.Forms.Form]        $formDoc
    [System.Windows.Forms.RichTextBox] $docView
    [int]           $modifiedTextLengthDelta  # We've found all matches with immutual Text, but will be changing actual document
    [String]        $packageName

    DocView($content, $packageName, $NoDefaultHighlighting = $false) {
        $self = $this
        $this.packageName = $packageName
        $this.modifiedTextLengthDelta = 0

        $this.formDoc = New-Object Windows.Forms.Form
        $this.formDoc.Text = "PyDoc for [$packageName] *** Click/Enter goto $packageName.WORD | Esc back | Ctrl+Wheel font | Space scroll"
        $this.formDoc.Size = New-Object Drawing.Point 830, 840
        $this.formDoc.Topmost = $false
        $this.formDoc.KeyPreview = $true
        $this.formDoc.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent

        $this.formDoc.Icon = $Script:form.Icon

        $this.docView = $global:RichTextBox_t::new()
        $this.docView.Location = New-Object Drawing.Point 7,7
        $this.docView.Size = New-Object Drawing.Point 800,810
        $this.docView.ReadOnly = $true
        $this.docView.Multiline = $true
        $this.docView.Font = New-Object System.Drawing.Font("Consolas", 11)
        $this.docView.WordWrap = $false
        $this.docView.DetectUrls = $true
        $this.docView.AllowDrop = $false
        $this.docView.RichTextShortcutsEnabled = $false
        $this.docView.Text = $content
        $this.formDoc.Controls.Add($this.docView)

        $this.docView.add_HandleCreated({
            param($Sender)
            $null = [SearchDialogHook]::new($Sender)
            $null = [TextBoxNavigationHook]::new($Sender)
        })

        $jumpToWord = ({
                param($clickedIndex)

                for ($charIndex = $clickedIndex; $charIndex -gt 0; $charIndex--) {
                    if ($self.docView.Text[$charIndex] -notmatch '\w|\.') {
                        $begin = $charIndex + 1
                        break
                    }
                }

                for ($charIndex = $clickedIndex; $charIndex -lt $self.docView.Text.Length; $charIndex++) {
                    if ($self.docView.Text[$charIndex] -notmatch '\w|\.') {
                        $end = $charIndex
                        break
                    }
                }

                $selectedLength = $end - $begin
                if ($selectedLength -gt 0) {
                    $clickedWord = $self.docView.Text.Substring($begin, $selectedLength)

                    $childViewer = Show-DocView "$($self.PackageName).${clickedWord}"
                    $childViewer.formDoc.add_Shown({
                            $childViewer.SetSize($self.GetSize())
                            $childViewer.SetLocation($self.GetLocation())
                        })
                    $childViewer.Show()
                }
            }.GetNewClosure())

        $this.docView.add_KeyDown({
            if ($_.KeyCode -eq 'Escape') {
                $self.formDoc.Close()
            }
            if ($_.KeyCode -eq 'Enter') {
                $charIndex = $self.docView.SelectionStart
                [void] $jumpToWord.Invoke($charIndex)
            }
        }.GetNewClosure())

        $this.docView.add_LinkClicked({
                Open-LinkInBrowser $_.LinkText
            }.GetNewClosure())

        $this.docView.Add_MouseClick({
                $clickedIndex = $self.docView.GetCharIndexFromPosition($_.Location)
                [void] $jumpToWord.Invoke($clickedIndex)
            }.GetNewClosure())

        $this.Resize()
        $this.formDoc.Add_Resize({
                $self.Resize()
            }.GetNewClosure())

        if (-not $NoDefaultHighlighting) {
            $this.Highlight_PEPLinks()
            $this.Highlight_PyDocSyntax()
        }
    }

    [void] Show() {
        $this.docView.Select(0, 0)
        $this.docView.ScrollToCaret()
        $this.formDoc.ShowDialog()
    }

    [void] Resize() {
        $this.docView.Width  = $this.formDoc.ClientSize.Width  - 15
        $this.docView.Height = $this.formDoc.ClientSize.Height - 15
    }

    [System.Drawing.Size] GetSize() {
        return $this.formDoc.Size
    }

    [System.Drawing.Point] GetLocation() {
        return $this.formDoc.DesktopLocation
    }

    [void] SetSize($size) {
        $this.formDoc.Size = $size
    }

    [void] SetLocation($location) {
        $this.formDoc.DesktopLocation = $location
    }

    [void] Alter_MatchingFragments($pattern, $selectionAlteringCode) {
        $this.modifiedTextLengthDelta = 0
        $matches = $pattern.Matches($this.docView.Text)

        if ($matches.Count -eq 0) {
            return
        }

        foreach ($match in $matches.Groups) {
            if ($match.Name -eq 0) {
                continue
            }
            [void] $this.docView.Select($match.Index + $this.modifiedTextLengthDelta, $match.Length)
            [void] $selectionAlteringCode.Invoke($match.Index + $this.modifiedTextLengthDelta, $match.Length, $match.Value)
        }
    }

    [void] Highlight_Text($pattern, $foreground, $useBold, $useUnderline) {
        $this.Alter_MatchingFragments($pattern, {
            $this.docView.SelectionColor = $foreground
            if ($useBold) {
                $this.docView.SelectionFont = $fontBold
            }
            if ($useUnderline) {
                $this.docView.SelectionFont = $fontUnderline
            }
        })
    }

    [void] Highlight_PyDocSyntax() {
        $this.Highlight_Text($Script:pyRegexStrQuote,  ($Script:DrawingColor::DarkGreen),   $false, $false)
        $this.Highlight_Text($Script:pyRegexNumber,    ($Script:DrawingColor::DarkMagenta), $false, $false)
        $this.Highlight_Text($Script:pyRegexSection,   ($Script:DrawingColor::DarkCyan),    $true,  $false)
        $this.Highlight_Text($Script:pyRegexKeyword,   ($Script:DrawingColor::DarkRed),     $true,  $false)
        $this.Highlight_Text($Script:pyRegexSpecial,   ($Script:DrawingColor::DarkOrange),  $true,  $false)
        $this.Highlight_Text($Script:pyRegexSpecialU,  ($Script:DrawingColor::DarkOrange),  $true,  $false)
    }

    [void] Highlight_PEPLinks() {
        # PEP Links
        $this.Alter_MatchingFragments($Script:pyRegexPEP, {
            param($index, $length, $match)

            $parts = "$match".Split(' -', [System.StringSplitOptions]::RemoveEmptyEntries)
            if ($parts.Count -ne 2) {
                return
            }

            ($PEP, $number) = $parts
            $pep_url = "${peps_url}pep-$(([int] $number).ToString('0000'))"

            $this.modifiedTextLengthDelta -= $this.docView.TextLength
            $this.docView.SelectedText = "$match [$pep_url] "  # keep the space to prevent link merging
            $this.modifiedTextLengthDelta += $this.docView.TextLength
        })

        # Subpackage Links
        $this.Alter_MatchingFragments($Script:pyRegexSubPkgs, {
            param($index, $length, $match)

            $startLine = $this.docView.GetLineFromCharIndex($index)
            for ($n = $startLine; $n -lt $this.docView.Lines.Count ; $n++) {
                $rawLine = $this.docView.Lines[$n]
                $line = $rawLine.TrimStart()
                $indentWidth = $rawLine.Length - $line.Length
                $line = $line.TrimEnd()
                if ($line.Length -eq 0) {
                    break
                }
                $index = $this.docView.GetFirstCharIndexFromLine($n)
                $this.docView.Select($index + $indentWidth, $line.Length)
                $this.docView.SelectionColor = $Script:DrawingColor::Navy
                $this.docView.SelectionFont  = $fontUnderline
            }
        })
    }

}

class SearchDialogHook {

    [string] $_query = $null
    [bool] $_reverse = $false
    [bool] $_caseSensitive = $false
    [System.Collections.Generic.List[string]] $_history

    SearchDialogHook([System.Windows.Forms.RichTextBox] $richTextBox) {
        $self = $this
        $self._history = [System.Collections.Generic.List[string]]::new()

        $richTextBox.add_KeyDown({
            param([System.Windows.Forms.RichTextBox] $Sender)

            switch ($_.KeyCode) {
                ([System.Windows.Forms.Keys]::OemQuestion) {
                        $controlKeysState = $null
                        $query = Request-UserString @"
Enter text for searching

Hold Shift to search backwards

You can hit (N)ext or (P)revious to skim over the matches
"@ 'Search text' $self._query $self._history ([ref] $controlKeysState)

                        if (-not $query) {
                            return
                        }

                        $reverse = $controlKeysState.ShiftKey
                        $self._query = $query
                        $self._reverse = $reverse
                        $self._history.Add($query)
                        $self._caseSensitive = HasUpperChars $query

                        $self.Search($Sender, $reverse)
                }

                'N' {
                    $self.Search($Sender, $self._reverse)
                }

                'P' {
                    $self.Search($Sender, -not ($self._reverse))
                }
            }
        }.GetNewClosure())

    }

    Search($Sender, [bool] $reverse) {
        $query = $this._query

        $searchOptions = [System.Windows.Forms.RichTextBoxFinds]::None

        if ($this._caseSensitive) {
            $searchOptions = $searchOptions -bor ([System.Windows.Forms.RichTextBoxFinds]::MatchCase)
        }

        if ($reverse) {
            $searchOptions = $searchOptions -bor ([System.Windows.Forms.RichTextBoxFinds]::Reverse)
        }

        if ($query) {
            $offset = if ($reverse) { -1 } else { 1 }

            if ($reverse) {
                $start = 0
                $end = $Sender.SelectionStart + $offset
            } else {
                $start = $Sender.SelectionStart + $offset
                $end = $Sender.TextLength - 1
            }

            Function InRange { param($_) ($_ -ge 0) -and ($_ -le ($Sender.TextLength - 1)) }

            if (($start -le $end) -and (InRange($start)) -and (InRange($end))) {
                [int] $found = $Sender.Find($query, $start, $end, $searchOptions)
            }
        }
    }

}

class TextBoxNavigationHook {

    TextBoxNavigationHook([System.Windows.Forms.RichTextBox] $richTextBox) {
        $richTextBox.add_KeyDown({
            param($Sender)
            if ($_.KeyCode -eq 'H') {
                $null = $SendMessage.Invoke($Sender.Handle, $WM_SCROLL, $SB_LINELEFT, 0)
                $_.Handled = $true
            }
            if ($_.KeyCode -eq 'J') {
                $null = $SendMessage.Invoke($Sender.Handle, $WM_VSCROLL, $SB_LINEDOWN, 0)
                $_.Handled = $true
            }
            if ($_.KeyCode -eq 'K') {
                $null = $SendMessage.Invoke($Sender.Handle, $WM_VSCROLL, $SB_LINEUP, 0)
                $_.Handled = $true
            }
            if ($_.KeyCode -eq 'L') {
                $null = $SendMessage.Invoke($Sender.Handle, $WM_SCROLL, $SB_LINERIGHT, 0)
                $_.Handled = $true
            }
            if (($_.KeyCode -eq 'G') -and (-not $_.Shift)) {
                $null = $SendMessage.Invoke($Sender.Handle, $WM_VSCROLL, $SB_PAGETOP, 0)
                $richTextBox.Select(0, 0)
                $_.Handled = $true
            }
            if (($_.KeyCode -eq 'G') -and $_.Shift) {
                $null = $SendMessage.Invoke($Sender.Handle, $WM_VSCROLL, $SB_PAGEBOTTOM, 0)
                $textLength = $richTextBox.TextLength
                $richTextBox.Select($textLength, $textLength)
                $_.Handled = $true
            }
            if ($_.KeyCode -eq 'Space') {
                $null = $SendMessage.Invoke($Sender.Handle, $WM_VSCROLL, $SB_PAGEDOWN, 0)
                $_.Handled = $true
            }
            if (($_.KeyCode -eq 'F') -and $_.Control) {
                $null = $SendMessage.Invoke($Sender.Handle, $WM_VSCROLL, $SB_PAGEDOWN, 0)
                $_.Handled = $true
            }
            if (($_.KeyCode -eq 'B') -and $_.Control) {
                $null = $SendMessage.Invoke($Sender.Handle, $WM_VSCROLL, $SB_PAGEUP, 0)
                $_.Handled = $true
            }
            if (($_.KeyCode -eq 'OemMinus') -and $_.Control) {
                if ($richTextBox.ZoomFactor -gt 0.1) {
                    $richTextBox.ZoomFactor -= $richTextBox.ZoomFactor * 0.1
                }
                $_.Handled = $true
            }
            if (($_.KeyCode -eq 'Oemplus') -and $_.Control) {
                if ($richTextBox.ZoomFactor -lt 10.0) {
                    $richTextBox.ZoomFactor += $richTextBox.ZoomFactor * 0.1
                }
                $_.Handled = $true
            }
            if (($_.KeyCode -eq 'D8') -and $_.Control) {
                $richTextBox.ZoomFactor = 1.0
                $_.Handled = $true
            }
        }.GetNewClosure())
    }

}


class ProcessWithPipedIO {
    hidden [System.Diagnostics.Process] $_process
    hidden [System.Threading.Tasks.TaskCompletionSource[int]] $_taskCompletionSource  # Keeps the exit code of a process
    hidden [System.Collections.Generic.List[string]] $_processOutput
    hidden [System.Collections.Generic.List[string]] $_processError
    hidden [System.Windows.Forms.Timer] $_timer

    ProcessWithPipedIO($Command, $Arguments) {
        $startInfo = [System.Diagnostics.ProcessStartInfo]::new()

        $startInfo.FileName = $Command
        $startInfo.Arguments = $Arguments

        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true

        $startInfo.RedirectStandardInput = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true

        $process = [System.Diagnostics.Process]::new()
        $process.StartInfo = $startInfo
        $process.EnableRaisingEvents = $true

        $self = $this

        $taskCompletionSource = [System.Threading.Tasks.TaskCompletionSource[int]]::new()

        $ExitedCallback = New-RunspacedDelegate ([EventHandler] {
            param($Sender, $EventArgs)
            if ($self._timer -eq $null) {
                $null = $self._taskCompletionSource.TrySetResult($self._process.ExitCode)
            }
        }.GetNewClosure())

        $process.add_Exited($ExitedCallback)

        $this._process = $process
        $this._taskCompletionSource = $taskCompletionSource
    }

    [System.Threading.Tasks.Task[int]] Start() {
        try {
            $started = $this._process.Start()

            if (-not $started) {
                throw [Exception]::new("Failed to start process $($this._process.StartInfo.FileName)")
            }

            $this._process.BeginOutputReadLine()
            $this._process.BeginErrorReadLine()

        } catch {
            $this._taskCompletionSource.SetException($_.Exception)
        }

        return $this._taskCompletionSource.Task
    }

    [System.Threading.Tasks.Task[int]] StartWithLogging([bool] $LogOutput, [bool] $LogErrors) {
        $self = $this

        if ($LogOutput) {
            $this._processOutput = [System.Collections.Generic.List[string]]::new()

            $OutputCallback = New-RunspacedDelegate ([System.Diagnostics.DataReceivedEventHandler] {
                param($Sender, $EventArgs)
                $null = $self._processOutput.Add($EventArgs.Data)
            }.GetNewClosure())

            $this._process.add_OutputDataReceived($OutputCallback)
        }

        if ($LogErrors) {
            $this._processError = [System.Collections.Generic.List[string]]::new()

            $ErrorCallback = New-RunspacedDelegate ([System.Diagnostics.DataReceivedEventHandler] {
                param($Sender, $EventArgs)
                $null = $self._processError.Add($EventArgs.Data)
            }.GetNewClosure())

            $this._process.add_ErrorDataReceived($ErrorCallback)
        }

        $null = $this.Start()

        $delegate = New-RunspacedDelegate ([EventHandler] {
            param($Sender)
            $count = $self.FlushBuffersToLog()

            if ($self._process.HasExited -and ($count -eq 0)) {
                $Sender.Stop()
                $null = $self._taskCompletionSource.TrySetResult($self._process.ExitCode)
                $self._timer = $null
            }
        }.GetNewClosure())

        $this._timer = [System.Windows.Forms.Timer]::new()
        $this._timer.Interval = 75
        $this._timer.add_Tick($delegate)
        $this._timer.Start()

        return $this._taskCompletionSource.Task
    }

    Kill() {
        $this._process.CancelOutputRead()
        $this._process.CancelErrorRead()
        TryTerminateGracefully($this._process)
    }

    hidden [int] _FlushContainerToLog($container, $color) {
        $count = 0
        if ($container) {
            $count = $container.Count
            if ($count -gt 0) {
                $lines = $container -join [Environment]::NewLine
                $container.Clear()
                WriteLog $lines -Background $color
            }
        }
        return $count
    }

    hidden [int] FlushBuffersToLog () {
        $outLines = $this._FlushContainerToLog($this._processOutput, 'LightBlue')
        $errLines = $this._FlushContainerToLog($this._processError, 'LightSalmon')
        return $outLines + $errLines
    }
}


class WidgetStateTransition {

    hidden [System.Collections.Generic.Stack[hashtable]] $_states
    hidden [System.Collections.Generic.List[System.Windows.Forms.Control]] $_controls
    hidden [System.Collections.Generic.HashSet[Action]] $_actions

    WidgetStateTransition () {
        $this._states = [System.Collections.Generic.Stack[hashtable]]::new()
        $this._controls = [System.Collections.Generic.List[System.Windows.Forms.Control]]::new()
    }

    [WidgetStateTransition] Add([System.Windows.Forms.Control] $control) {
        $null = $this._controls.Add($control)
        return $this
    }

    [WidgetStateTransition] AddRange([System.Windows.Forms.Control[]] $controls) {
        $null = $this._controls.AddRange($controls)
        return $this
    }

    [WidgetStateTransition] Transform([hashtable] $properties) {
        $states = @{}
        foreach ($control in $this._controls) {
            $widgetState = @{}
            foreach ($property in $properties.GetEnumerator()) {
                # Write-Host $property.Key '=' $property.Value ' of ' $property.Value.GetType()
                if ($property.Value -isnot [ScriptBlock]) {
                    $null = $widgetState.Add($property.Key, $control."$($property.Key)")
                    $control."$($property.Key)" = $property.Value
                } else {
                    $handlers = [WidgetStateTransition]::SpliceEventHandlers($control, $property.Key, $property.Value)
                    $null = $widgetState.Add($property.Key, $handlers)
                }
            }
            $states[$control] = $widgetState
        }
        $this._controls.Clear()
        $this._states.Push($states)
        return $this
    }

    [WidgetStateTransition] Reverse() {
        $states = $this._states.Pop()
        foreach ($widgetState in $states.GetEnumerator()) {

            switch ($widgetState.Key) {

                { $_ -is [System.Windows.Forms.Control] } {
                    $control, $properties = $widgetState.Key, $widgetState.Value
                    foreach ($property in $properties.GetEnumerator()) {
                        # Write-Host 'REV ' $property.Key '=' $property.Value ' of ' $property.Value.GetType()
                        if ($property.Value -isnot [delegate[]]) {
                            $control."$($property.Key)" = $property.Value
                        } else {
                            $null = [WidgetStateTransition]::SpliceEventHandlers($control, $property.Key, $property.Value)
                        }
                    }
                    break
                }

                { $_ -is [string] } {
                    $name, $value = $widgetState.Key, $widgetState.Value
                    Set-Variable -Name $name -Value $value -Scope Global
                    break
                }

            }
        }
        return $this
    }

    [WidgetStateTransition] ReverseAll() {
        while ($this._states.Count -gt 0) {
            $null = $this.Reverse()
        }
        return $this
    }

    static [object] ReverseAllAsync() {
        $delegate = New-RunspacedDelegate ([Action[System.Threading.Tasks.Task[object], object]]{
            param([System.Threading.Tasks.Task[object]] $task, [object] $widgetStateTransition)
            $null = ($widgetStateTransition -as [WidgetStateTransition]).ReverseAll()
        })
        return $delegate
    }

    [WidgetStateTransition] Debounce([Action] $action) {
        if (-not $this._actions.Contains($action)) {
            $this._actions.Add($action)
        }
        return $this
    }

    static [delegate[]] SpliceEventHandlers([System.Windows.Forms.Control] $control, [string] $event, $handlers) {

        $type = $control.GetType()

        $propertyInfo = $type.GetProperty('Events',
            [System.Reflection.BindingFlags]::Instance -bor
            [System.Reflection.BindingFlags]::NonPublic -bor
            [System.Reflection.BindingFlags]::Static)

        $eventHandlerList = $propertyInfo.GetValue($control)

        $fieldInfo = [System.Windows.Forms.Control].GetField("Event$event",
            [System.Reflection.BindingFlags]::Static -bor
            [System.Reflection.BindingFlags]::NonPublic)

        $eventKey = $fieldInfo.GetValue($control)

        $eventHandler = $eventHandlerList[$eventKey]

        $old = [System.Collections.Generic.List[delegate]]::new()

        if ($eventHandler) {
            $invocationList = @($eventHandler.GetInvocationList())

            foreach ($handler in $invocationList) {
                $null = $old.Add($handler)
                $control."remove_$event"($handler)
            }
        }

        foreach ($handler in $handlers) {
            $control."add_$event"($handler)
        }

        return $old
    }

    [WidgetStateTransition] GlobalVariables([hashtable] $variables) {
        $state = @{}
        foreach ($variable in $variables.GetEnumerator()) {
            $value = Get-Variable -Name $variable.Key -Scope Global -ValueOnly -ErrorAction SilentlyContinue
            Set-Variable -Name $variable.Key -Value $variable.Value -Scope Global -Force
            $state[$variable.Key] = $value
        }
        $this._states.Push($state)
        return $this
    }

}


Function global:Show-DocView($packageName, $SetContent = $null, $Highlight = $null, [switch] $NoDefaultHighlighting) {
    if (-not $SetContent) {
        $content = (Get-PyDoc $packageName) -join "`n"
    } else {
        $content = $SetContent
    }

    $viewer = New-Object DocView -ArgumentList @($content, $packageName, $NoDefaultHighlighting)

    if ($Highlight) {
        $viewer.Highlight_Text($Highlight, ([System.Drawing.Color]::Navy), $false, $true)
    }

    return $viewer
}

Function Write-PipPackageCounter {
    $count = $global:dataModel.Rows.Count
    WriteLog "Now $count packages in the list."
}

Function Store-CheckedPipSearchResults() {
    $selected = New-Object System.Data.DataTable
    Init-PackageSearchColumns $selected

    $isInstallMode = $global:dataModel.Columns.Contains('Description')
    if ($isInstallMode) {
        foreach ($row in $global:dataModel) {
            if ($row.Select) {
                $selected.ImportRow($row)
            }
        }
    }

    return ,$selected
}

Function Get-PipSearchResults($request) {
    $pip_exe = Get-CurrentInterpreter 'PipExe' -Executable
    if (-not $pip_exe) {
        WriteLog 'pip is not found!'
        return 0
    }

    $args = New-Object System.Collections.ArrayList
    $null = $args.Add('search')
    $null = $args.Add("$request")
    $output = & $pip_exe $args

    $r = [regex] '^(.*?)\s*\((.*?)\)\s+-\s+(.*?)$'

    $count = 0
    foreach ($line in $output) {
        $m = $r.Match($line)
        if ([String]::IsNullOrEmpty($m.Groups[1].Value)) {
            continue
        }
        $row = $global:dataModel.NewRow()
        $row.Select = $false
        $row.Package = $m.Groups[1].Value
        $row.Installed = $m.Groups[2].Value
        $row.Description = $m.Groups[3].Value
        $row.Type = 'pip'
        $row.Status = ''
        $global:dataModel.Rows.Add($row)
        $count += 1
    }

    return $count
}

Function Get-CondaSearchResults($request) {
    $conda_exe = Get-CurrentInterpreter 'CondaExe' -Executable
    if (-not $conda_exe) {
        WriteLog 'conda is not found!'
        return 0
    }
    $arch = Get-CurrentInterpreter 'Arch'

    # channels [-c]:
    #   anaconda = main, free, pro, msys2[windows]
    # --info should give better details but not supported on every conda
    $totalCount = 0
    $channels = Get-PipsSetting 'CondaChannels'

    Function EnsureProperties {
        param($Object, $QueryNameLabel, $Separator = ', ')

        if (($Object -eq $null) -or ($Object -isnot [PSCustomObject])) {
            return [string]::Empty
        }

        $result = [System.Collections.Generic.List[string]]::new()

        foreach ($NameLabel in $QueryNameLabel.GetEnumerator()) {
            ($PropertyName, $Label) = ($NameLabel.Key, $NameLabel.Value)

            if (($Object.PSObject.Properties.Name -contains $PropertyName) -and
                (-not ([string]::IsNullOrWhiteSpace($Object."$PropertyName")))) {
                $null = $result.Add([string]::Concat($Label, ': ', $Object."$PropertyName"))
            }
        }

        return $result -join $Separator
    }

    foreach ($channel in $channels) {
        WriteLog "Searching on channel: $channel ... " -NoNewline

        $items = & $conda_exe search -c $channel --json $request | ConvertFrom-Json

        $count = 0
        $items.PSObject.Properties | ForEach-Object {
            $name = $_.Name
            $item = $_.Value

            foreach ($package in $item) {
                $row = $global:dataModel.NewRow()
                $row.Select = $false
                $row.Package = $name
                $row.Installed = $package.version

                $row.Description = EnsureProperties $package @{
                    channel='Channel';
                    arch='Architecture';
                    build='Build';
                    date='Date';
                    license_family='License';
                }

                $row.Type = 'conda'
                $row.Status = ''
                $global:dataModel.Rows.Add($row)
                $count += 1
            }
        }

        WriteLog "$count packages."
        $totalCount += $count
    }

    return $totalCount
}

Function Get-GithubSearchResults ($request) {
    $json = Download-String ($github_search_url -f [System.Web.HttpUtility]::UrlEncode($request))
    $info = $json | ConvertFrom-Json
    $items = $info.'items'
    $count = 0
    $items | ForEach-Object {
        $row = $global:dataModel.NewRow()
        $row.Select = $false
        $row.Package = $_.'full_name'
        $row.Installed = ($_.'pushed_at' -replace 'T',' ') -replace 'Z',''
        $row.Description = "$($_.'stargazers_count') $([char] 0x2729) $($_.'forks') $([char] 0x2442) $($_.'open_issues') $([char] 0x2757) $($_.'description')"
        $row.Type = 'git'
        $row.Status = ''
        $global:dataModel.Rows.Add($row)
        $count++
    }
    return $count
}

Function global:Get-PluginSearchResults($request) {
    $count = 0
    foreach ($plugin in $global:plugins) {
        $packages = $plugin.GetSearchResults(
            $request,
            (Get-CurrentInterpreter 'Version'),
            (Get-CurrentInterpreter 'Bits'))

        foreach ($package in $packages) {
            $row = $global:dataModel.NewRow()
            $row.Select = $false
            $row.Package = $package.Name
            $row.Installed = $package.Version
            $row.Description = $package.Description
            $row.Type = $package.Type
            $row.Status = ''
            $global:dataModel.Rows.Add($row)
            $count++
        }
    }
    return $count
}

Function global:Get-GithubRepoTags($gitLinkInfo, $ContinueWith = $null) {
    if (-not $gitLinkInfo) {
        return $null
    }

    if ($gitLinkInfo.PSObject.Properties.Name -contains 'Path') {
        $git = Get-Bin 'git'
        if (-not $git) {
            return $null
        }
        $tags = & $git -C $gitLinkInfo.Path tag
        if ($ContinueWith) {
            $null = & $ContinueWith $tags
            return
        } else {
            return $tags
        }
    }

    $github_tags_url = 'https://api.github.com/repos/{0}/{1}/tags'
    $url = $github_tags_url -f $gitLinkInfo.User,$gitLinkInfo.Repo

    if ($ContinueWith) {
        $null = Download-String $url -ContinueWith {
            param($json)
            if (-not [string]::IsNullOrEmpty($json)) {
                $tags = $json | ConvertFrom-Json | ForEach-Object { $_.Name }
                $null = & $ContinueWith $tags
                return
            }
        }.GetNewClosure()
        return
    }

    $json = Download-String $url
    if (-not [string]::IsNullOrEmpty($json)) {
        $tags = $json | ConvertFrom-Json | Select-Object -ExpandProperty Name # ForEach-Object { $_.Name }
        return $tags
    }

    return $null
}

Function Get-SearchResults($request) {
    $previousSelected = Store-CheckedPipSearchResults
    Clear-Rows
    Init-PackageSearchColumns $global:dataModel

    $dataGridView.BeginInit()
    $global:dataModel.BeginLoadData()

    foreach ($row in $previousSelected) {
        $global:dataModel.ImportRow($row)
    }

    $pipCount = Get-PipSearchResults $request
    $condaCount = Get-CondaSearchResults $request
    $pluginCount = Get-PluginSearchResults $request
    $githubCount = Get-GithubSearchResults $request

    $global:dataModel.EndLoadData()
    $dataGridView.EndInit()

    return @{PipCount=$pipCount; CondaCount=$condaCount; GithubCount=$githubCount; PluginCount=$pluginCount; Total=($pipCount + $condaCount + $githubCount + $pluginCount)}
}

 Function Get-PythonPackages($outdatedOnly = $true) {
    WriteLog
    WriteLog 'Updating package list... '

    $python_exe = Get-CurrentInterpreter 'PythonExe' -Executable
    $pip_exe = Get-CurrentInterpreter 'PipExe' -Executable
    $conda_exe = Get-CurrentInterpreter 'CondaExe' -Executable

    if ($python_exe) {
        WriteLog (& $python_exe --version 2>&1)
    } else {
        WriteLog 'Python is not found!'
    }

    if ($pip_exe) {
        WriteLog (& $pip_exe --version 2>&1)
    } else {
        WriteLog 'pip is not found!'
    }

    WriteLog

    Clear-Rows
    Init-PackageUpdateColumns $global:dataModel


    if ($pip_exe) {
        $args = New-Object System.Collections.ArrayList
        $null = $args.Add('list')
        $null = $args.Add('--format=columns')

        if ($outdatedOnly) {
            $null = $args.Add('--outdated')
        }

        if ($Script:isolatedCheckBox.Checked) {
            $null = $args.Add('--isolated')  # ignore user config
            $null = $args.Add('--local')     # ignore global packages
        }

        $pip_list = & $pip_exe $args
        $pipPackages = $pip_list `
            | Select-Object -Skip 2 `
            | % { $_ -replace '\s+', ' ' } `
            | ConvertFrom-Csv -Header $csv_header -Delimiter ' ' `
            | where { -not [String]::IsNullOrEmpty($_.Package) }
    }

    Function Add-PackagesToTable($packages, $defaultType = [String]::Empty) {
        for ($n = 0; $n -lt $packages.Count; $n++) {
            $row = $global:dataModel.NewRow()
            $row.Select = $false
            $package = $packages[$n]
            $row.Package = $package.Package

            $availableKeys = $package.PSObject.Properties.Name

            # write-host $availableKeys
            # Write-Host $package.Package $availableKeys

            if ($availableKeys -contains 'Installed') {
                $row.Installed = $package.Installed
            }

            if ($availableKeys -contains 'Latest') {
                $row.Latest = $package.Latest
            }

            if (($availableKeys -contains 'Type') -and (-not [string]::IsNullOrWhiteSpace($package.Type))) {
                $row.Type = $package.Type
            } else {
                $row.Type = $defaultType
            }

            $global:dataModel.Rows.Add($row)
        }
    }

    $global:dataModel.BeginLoadData()

    $pipCount = 0
    $condaCount = 0
    $builtinCount = 0
    $otherCount = 0

    if ($pip_exe) {
        Add-PackagesToTable $pipPackages 'pip'
        $pipCount = $pipPackages.Count
    }
    if ($conda_exe) {
        $condaPackages = Get-CondaPackages $outdatedOnly
        Add-PackagesToTable $condaPackages 'conda'
        $condaCount = $condaPackages.Count
    }
    if (-not $outdatedOnly) {
        $builtinPackages = Get-PythonBuiltinPackages
        Add-PackagesToTable $builtinPackages 'builtin'

        $otherPackages = Get-PythonOtherPackages
        Add-PackagesToTable $otherPackages 'other'

        $builtinCount = $builtinPackages.Count
        $otherCount = $otherPackages.Count
    }
    $global:dataModel.EndLoadData()

    $Script:outdatedOnly = $outdatedOnly
    Highlight-PythonPackages

    WriteLog 'Package list updated.'
    WriteLog 'Double click or [Ctrl+Enter] a table row to open PyPi, Anaconda.com or github.com in browser'

    $count = $global:dataModel.Rows.Count
    WriteLog "Total $count packages: $builtinCount builtin, $pipCount pip, $condaCount conda, $otherCount other"
    WriteLog
}

Function Select-VisiblePipPackages($value) {
    $global:dataModel.BeginLoadData()
    for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
        if ($dataGridView.Rows[$i].DataBoundItem.Row.Type -in @('builtin', 'other') ) {
            continue
        }
        $dataGridView.Rows[$i].DataBoundItem.Row.Select = $value
    }
    $global:dataModel.EndLoadData()
}

Function Select-PipPackages($value) {
    $global:dataModel.BeginLoadData()
    for ($i = 0; $i -lt $global:dataModel.Rows.Count; $i++) {
       $global:dataModel.Rows[$i].Select = $value
    }
    $global:dataModel.EndLoadData()
}

Function Set-Unchecked($index) {
    $global:dataModel.Rows[$index].Select = $false
}

Function Test-PackageInList($name) {
    $n = 0
    foreach ($item in $global:dataModel.Rows) {
        if ($item.Package -eq $name) {
            return $n
        }
        $n++
    }
    return -1
}

Function Clear-Rows() {
    $Script:outdatedOnly = $true
    $global:inputFilter.Clear()
    $global:dataModel.DefaultView.RowFilter = $null
    $dataGridView.ClearSelection()

    #if ($dataGridView.SortedColumn) {
    #    $dataGridView.SortedColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    #}

    $global:dataModel.DefaultView.Sort = [String]::Empty

    $dataGridView.BeginInit()
    $global:dataModel.BeginLoadData()

    $global:dataModel.Rows.Clear()

    $global:dataModel.EndLoadData()
    $dataGridView.EndInit()
}

Function global:Set-SelectedNRow($n) {
    if ($n -ge $dataGridView.RowCount -or $n -lt 0) {
        return
    }
    $dataGridView.ClearSelection()
    $dataGridView.FirstDisplayedScrollingRowIndex = $n
    $dataGridView.Rows[$n].Selected = $true
    $dataGridView.CurrentCell = $dataGridView[0, $n]
}

Function global:Set-SelectedRow($selectedRow) {
    $Script:dataGridView.ClearSelection()
    foreach ($vRow in $Script:dataGridView.Rows) {
        if ($vRow.DataBoundItem.Row -eq $selectedRow) {
            $vRow.Selected = $true
            $dataGridView.FirstDisplayedScrollingRowIndex = $vRow.Index
            $dataGridView.CurrentCell = $dataGridView[0, $vRow.Index]
            break
        }
    }
}

Function Check-PipDependencies {
    WriteLog 'Checking dependencies...'

    $pip_exe = Get-CurrentInterpreter 'PipExe' -Executable
    if (-not $pip_exe) {
        WriteLog 'pip is not found!'
        return
    }

    $result = & $pip_exe check 2>&1
    $result = $result -join "`n"

    if ($result -match 'No broken requirements found') {
        WriteLog "OK" -Background ([Drawing.Color]::LightGreen)
        WriteLog $result
    } else {
        WriteLog "NOT OK" -Background ([Drawing.Color]::LightSalmon)
        WriteLog $result
    }
}

Function global:Select-PipAction($actionName) {
    $n = 0
    foreach ($item in $global:actionsModel) {
        if ($item.Name -eq $actionName) {
            $actionListComboBox.SelectedIndex = $n
            return
        }
        $n++
    }
}

Function global:WidgetStateTransitionForCommandButton($button) {
    $widgetStateTransition = [WidgetStateTransition]::new()
    $disableThemWhileWorking = @(
        $WIDGET_GROUP_COMMAND_BUTTONS ;
        $WIDGET_GROUP_ENV_BUTTONS ;
        $WIDGET_GROUP_INSTALL_BUTTONS ;
        $interpretersComboBox ;
        $actionListComboBox
    )
    $null = $widgetStateTransition.AddRange($disableThemWhileWorking).Transform(@{Enabled=$false})
    $null = $widgetStateTransition.Add($button).Transform(@{Text='Cancel';Enabled=$true;Click={
            $response = [System.Windows.Forms.MessageBox]::Show('Sure?', 'Cancel', [System.Windows.Forms.MessageBoxButtons]::YesNo)
            if ($response -eq 'Yes') {

            }
        }})
    $null = $widgetStateTransition.GlobalVariables(@{
            APP_MODE=([AppMode]::Working);
        })
    return $widgetStateTransition
}

Function global:Execute-PipAction {
    $action = $global:actionsModel[$actionListComboBox.SelectedIndex]

    $tasksOkay = 0
    $tasksFailed = 0

    $queue = [System.Collections.Generic.Queue[object]]::new()

    for ($i = 0; $i -lt $global:dataModel.Rows.Count; $i++) {
       if ($global:dataModel.Rows[$i].Select -eq $true) {
            $null = $queue.Enqueue($global:dataModel.Rows[$i])
       }
    }

    $execActionTaskCompletionSource = [System.Threading.Tasks.TaskCompletionSource[object]]::new()

    $reportBatchResults = New-RunspacedDelegate([Action[System.Threading.Tasks.Task]] {
        $execActionTaskCompletionSource.SetResult($null)

        if (($tasksOkay -eq 0) -and ($tasksFailed -eq 0)) {
            WriteLog 'Nothing is selected.' -Background LightSalmon
        } else {
            WriteLog ''
            WriteLog '----'
            WriteLog "All tasks finished, $tasksOkay ok, $tasksFailed failed." -Background LightGreen
            WriteLog 'Select a row to highlight the relevant log piece'
            WriteLog 'Double click or [Ctrl+Enter] a table row to open PyPi, Anaconda.com or github.com in browser'
            WriteLog '----'
            WriteLog ''
        }
    }.GetNewClosure())

    if ($action.TakesList) {
        $null = $action.Execute($queue.ToArray())
        return
    } else {
        # ApplyAsync :: [DataRow] -> (ElementContext{ DataRow } -> int) -> DataRow )
        $null = ApplyAsync $queue ([Func[object, object]] {
            param($ElementContext)
            $dataRow = $ElementContext.Element
            $ApplyAsyncContext = $ElementContext.ApplyAsyncContext
            $ApplyAsyncContextType = ([System.Type] "System.Threading.Tasks.TaskCompletionSource[$($ElementContext.ApplyAsyncContextType)]")

            Set-SelectedRow $dataRow
            $name, $installed, $type = $dataRow.Package, $dataRow.Installed, $dataRow.Type

            WriteLog "$($action.Name) $name" -Background LightPink

            $taskCompletionSource = $ApplyAsyncContextType::new()

            # $result = $action.Execute($name, $type, $installed) -join "`n"
            if ($action) {
                # (Get-CurrentInterpreter 'PythonExe')
            }

            $process = [ProcessWithPipedIO]::new('cat', @('D:\work\pyfmt-big.txt'))
            $task = $process.StartWithLogging($true, $true)
            $continuation = New-RunspacedDelegate([Action[System.Threading.Tasks.Task[int], object]] {
                param([System.Threading.Tasks.Task[int]] $task, [object] $locals)

                if ($task.IsCompleted -and (-not $task.IsFaulted)) {
                    $exitCode = $task.Result
                    WriteLog "Exited with code $exitCode" -Background DarkGreen -Foreground White
                } else {
                    $message = $task.Exception.InnerException
                    WriteLog "Failed: $message" -Background DarkRed -Foreground White
                }

                $global:dataModel.Columns['Status'].ReadOnly = $false
                $global:dataModel.Columns['Status'].ReadOnly = $true
                $logTo = (GetLogLength) - $locals.logFrom
                $locals.dataRow | Add-Member -Force -MemberType NoteProperty -Name LogFrom -Value $locals.logFrom
                $locals.dataRow | Add-Member -Force -MemberType NoteProperty -Name LogTo -Value $logTo

                $null = $locals.taskCompletionSource.SetResult($locals.ApplyAsyncContext)
            })

            $null = $task.ContinueWith(
                $continuation,
                @{
                    taskCompletionSource=$taskCompletionSource;
                    ApplyAsyncContext=$ApplyAsyncContext;
                    dataRow=$dataRow;
                    process=$process;
                    logFrom=(GetLogLength);
                },
                [System.Threading.Tasks.TaskScheduler]::FromCurrentSynchronizationContext())

            return $taskCompletionSource.Task

        }.GetNewClosure()) $reportBatchResults
    }

    return $execActionTaskCompletionSource.Task
}

# by https://superuser.com/users/243093/megamorf
function Test-is64Bit {
    param($FilePath)

    [int32]$MACHINE_OFFSET = 4
    [int32]$PE_POINTER_OFFSET = 60

    [byte[]]$data = New-Object -TypeName System.Byte[] -ArgumentList 4096
    $stream = New-Object -TypeName System.IO.FileStream -ArgumentList ($FilePath, 'Open', 'Read')
    $null = $stream.Read($data, 0, 4096)

    [int32]$PE_HEADER_ADDR = [System.BitConverter]::ToInt32($data, $PE_POINTER_OFFSET)
    [int32]$machineUint = [System.BitConverter]::ToUInt16($data, $PE_HEADER_ADDR + $MACHINE_OFFSET)
    $stream.Close()

    $result = "" | select FilePath, FileType, Is64Bit
    $result.FilePath = $FilePath
    $result.Is64Bit = $false

    switch ($machineUint)
    {
        0      { $result.FileType = 'Native' }
        0x014c { $result.FileType = 'x86' }
        0x0200 { $result.FileType = 'Itanium' }
        0x8664 { $result.FileType = 'x64'; $result.is64Bit = $true; }
    }

    $result
}

# from https://gallery.technet.microsoft.com/scriptcenter/Check-for-Key-Presses-with-7349aadc/file/148286/2/Test-KeyPress.ps1
Function global:Test-KeyPress
{
    <#
        .SYNOPSIS
        Checks to see if a key or keys are currently pressed.

        .DESCRIPTION
        Checks to see if a key or keys are currently pressed. If all specified keys are pressed then will return true, but if
        any of the specified keys are not pressed, false will be returned.

        .PARAMETER Keys
        Specifies the key(s) to check for. These must be of type "System.Windows.Forms.Keys"

        .EXAMPLE
        Test-KeyPress -Keys ControlKey

        Check to see if the Ctrl key is pressed

        .EXAMPLE
        Test-KeyPress -Keys ControlKey,Shift

        Test if Ctrl and Shift are pressed simultaneously (a chord)

        .LINK
        Uses the Windows API method GetAsyncKeyState to test for keypresses
        http://www.pinvoke.net/default.aspx/user32.GetAsyncKeyState

        The above method accepts values of type "system.windows.forms.keys"
        https://msdn.microsoft.com/en-us/library/system.windows.forms.keys(v=vs.110).aspx

        .LINK
        http://powershell.com/cs/blogs/tips/archive/2015/12/08/detecting-key-presses-across-applications.aspx

        .INPUTS
        System.Windows.Forms.Keys

        .OUTPUTS
        System.Boolean
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Windows.Forms.Keys[]]
        $Keys
    )

    # use the User32 API to define a keypress datatype
    $Signature = @'
[DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)]
public static extern short GetAsyncKeyState(int virtualKeyCode);
'@
    $API = Add-Type -MemberDefinition $Signature -Name 'Keypress' -Namespace Keytest -PassThru

    # test if each key in the collection is pressed
    $Result = foreach ($Key in $Keys)
    {
        [bool]($API::GetAsyncKeyState($Key) -eq -32767)
    }

    # if all are pressed, return true, if any are not pressed, return false
    $Result -notcontains $false
}


Function Get-PipDistributionInfo {
    $python_code = @'
import pkg_resources
import json
pkgs = pkg_resources.working_set
info = {p.key: {'deps': {r.name: [str(s) for s in r.specifier] for r in p.requires()}, 'extras': p.extras} for p in pkgs}
print(json.dumps(info))
'@ -join ';'

    $output = & (Get-CurrentInterpreter 'PythonExe') -c "`"$python_code`""
    $pkgs = $output | ConvertFrom-Json
    return $pkgs
}

Function Get-AsciiTree($name,
                       $distributionInfo,
                       $indent = 0,
                       $hasSibling = $false,
                       $dangling = (New-Object 'System.Collections.Generic.Stack[int]'),
                       $isExtra = $false,
                       $loopTracking = (New-Object 'System.Collections.Generic.HashSet[string]')) {

    if (($distributionInfo -eq $null) -or ($distributionInfo.PSObject.Properties.Name -notcontains $name)) {
        $children = @()
    } else {
        $children = @($distributionInfo."$name"."deps" |
            ForEach-Object { $_.PSObject.Properties } |
            Select-Object -ExpandProperty Name)  # dep name list from JSON from pip

        $extras = @($distributionInfo."$name"."extras")
        if (($extras -ne $null) -and ($extras.Length -gt 0)) {
            $children = @($children; $extras)
        }

        if ($children.Length -gt 1) {
            $children = @($children) | Sort-Object
        }
    }

    # WriteLog "$name children: '$children' $($children.Length) sibl=$hasSibling"
    $hasChildren = $children.Length -gt 0
    $isLooped = $loopTracking.Contains($name)

    $prefix = if ($indent -gt 0) {
        if ($hasChildren -and -not $isLooped) {
            if ($hasSibling) {
                "$(' ' * $indent * 4)├$('─' * (4 - 1))┬ "
            } else {
                "$(' ' * $indent * 4)└$('─' * (4 - 1))┬ "
            }
        } else {
            if ($hasSibling) {
                "$(' ' * $indent * 4)├$('─' * 4) "
            } else {
                "$(' ' * $indent * 4)└$('─' * 4) "
            }
        }
    } else {
        ''
    }

    if (-not [string]::IsNullOrEmpty($prefix)) {
        foreach ($i in $dangling) {
            if ($i -ne -1) {
                $prefix = $prefix.Remove($i, 1).Insert($i, '│')
            }
        }
    }

    $suffixList = @()
    if ($isExtra) { $suffixList += @('*') }
    if ($isLooped) { $suffixList += @('∞') }
    if (-not $global:autoCompleteIndex.Contains($name.ToLower())) { $suffixList += @('x') }  # unavailable package
    $suffix = if ($suffixList.Count -eq 0) { '' } else { " ($($suffixList -join ' '))" }

    "${prefix}${name}${suffix}"  # Add a line to the Return Stack

    if ($hasSibling) {
           [void]$dangling.Push($indent * 4)
    } else {
        [void]$dangling.Push(-1)
    }

    if (-not $isLooped) {
        [void]$loopTracking.Add($name)
        $i = 1
        foreach ($child in $children) {
            $childHasSiblings = ($i -lt $children.Length) -and ($children.Length -gt 1)
            Get-AsciiTree $child `
                $distributionInfo `
                ($indent + 1) `
                $childHasSiblings `
                $dangling `
                ($child -in $extras) `
                $loopTracking
            $i++
        }
        [void]$loopTracking.Remove($name)
    }

    [void]$dangling.Pop()
}

Function global:Get-DependencyAsciiGraph($name) {
    Prepare-PackageAutoCompletion  # for checking presence of pkg in the index
    $distributionInfo = Get-PipDistributionInfo
    $asciiTree = Get-AsciiTree $name.ToLower() $distributionInfo
    return $asciiTree
}

Function global:Show-ConsoleWindow([bool] $show) {
    Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

    $consolePtr = [Console.Window]::GetConsoleWindow()
    $value = if ($show) { 1 } else { 0 }
    [Console.Window]::ShowWindow($consolePtr, $value)
}

Function Save-PipsSettings {
    $settingsPath = "$($env:LOCALAPPDATA)\pips"
    $null = New-Item -Force -ItemType Directory -Path $settingsPath
    $userInterpreterRecords = $interpreters | Where-Object { $_.User }
    $settings."envs" = @($userInterpreterRecords)
    try {
        $settings | ConvertTo-Json -Depth 25 | Out-File "$settingsPath\settings.json"
    } catch {
    }
}

Function Load-PipsSettings {
    $settingsFile = "$($env:LOCALAPPDATA)\pips\settings.json"
    $global:settings = @{}
    if (Exists-File $settingsFile) {
        try {
            $settings = (Get-Content $settingsFile | ConvertFrom-Json)
            $global:settings."envs" = $settings.envs
            $global:settings."condaChannels" = $settings.condaChannels
        } catch {
        }
    }
}

$global:plugins = [System.Collections.ArrayList]::new()

Function global:Serialize([object] $Object, [string] $FileName) {
    $fs = New-Object System.IO.FileStream "$FileName", ([System.IO.FileMode]::Create)
    $bf = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
    $bf.Serialize($fs, $Object)
    $fs.Close()
}

Function global:Deserialize([string] $FileName) {
    $fs = New-Object System.IO.FileStream "$FileName", ([System.IO.FileMode]::Open)
    $bf = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
    $result = $bf.Deserialize($fs)
    $fs.Close()
    return $result
}

Function global:Get-PackageInfoFromWheelName($name) {
    # Related documentation:
    # https://www.python.org/dev/peps/pep-0425/

    $parser = [regex] '(?<Distribution>[^-]+)-(?<VersionCanonical>[0-9]+(?:\.[0-9]+)*)(?<VersionExtra>[^-]+)?(:?-(?<BuildTag>\d[^-.]*))?(:?-(?<PythonTag>(:?py|cp|ip|pp|jy)[^-]+))?(:?-(?<AbiTag>none|(:?(:?abi|py|cp|ip|pp|jy)[^-.]+)))?(?:-(?<PlatformTag>(:?any|linux|mac|win)[^.]*))?(?<ArchiveExtension>(:?\.[a-z0-9]+)+)'

    $groups = $parser.Match($name).Groups

    if ($groups.Count -lt 2) {
        return $null
    }

    Function Get-GroupValueOrNull($group) {
        if (-not $group.Success -or [string]::IsNullOrWhiteSpace($group.Value)) {
            return $null
        } else {
            return $group.Value.Trim()
        }
    }

    $versionCanonical = (Get-GroupValueOrNull $groups['VersionCanonical'])
    $versionExtra = (Get-GroupValueOrNull $groups['VersionExtra'])

    return @{
        'Distribution'=(Get-GroupValueOrNull $groups['Distribution']);
        'VersionCanonical'=$versionCanonical;
        'VersionExtra'=$versionExtra;
        'Version'="${versionCanonical}${versionExtra}";
        'Build'=(Get-GroupValueOrNull $groups['BuildTag']);
        'Python'=(Get-GroupValueOrNull $groups['PythonTag']);
        'ABI'=(Get-GroupValueOrNull $groups['AbiTag']);
        'Platform'=(Get-GroupValueOrNull $groups['PlatformTag']);
        'ArchiveExtension'=(Get-GroupValueOrNull $groups['ArchiveExtension']);
    }
}

Function global:Test-CanInstallPackageTo([string] $pythonVersion, [string] $pythonArch) {
    $pythonVersion = $pythonVersion -replace '\.',''
    $pyMajor = $pythonVersion[0]
    $testVersion = [regex] "cp$pythonVersion|py$pyMajor"
    $testAbi = [regex] "none|$testVersion"

    return {
        param([hashtable] $wheelInfo)

        if (($wheelInfo.Python -and
            ($wheelInfo.Python -notmatch $testVersion)) `
            -or
            ($wheelInfo.ABI -and
            ($wheelInfo.ABI -notmatch $testAbi)) `
            -or
            ($wheelInfo.Platform -and
            ($wheelInfo.Platform -ne 'any') -and
            -not ($wheelInfo.Platform.Contains($pythonArch)))) {
                return $false
        } else {
            return $true
        }
    }.GetNewClosure()
}

Function Load-Plugins() {
    $PipsRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath('.\')
    $plugins = Get-ChildItem "$PipsRoot\external-repository-providers" -Directory -Depth 0
    foreach ($plugin in $plugins) {
        WriteLog "Loading plugin $($plugin.Name)"
        Import-Module "$($plugin.FullName)"
        $instance = & (Get-Module $plugin.Name).ExportedFunctions.NewPluginInstance
        $pluginConfigPath = "$($env:LOCALAPPDATA)\pips\plugins\$($instance.GetPluginName())"
        [void] $instance.Init(
            $pluginConfigPath,
            {
                param($url, $destination)
                [string[]] $output = & (Get-CurrentInterpreter 'PythonExe') -m pip download --no-deps --no-index --progress-bar off --dest $destination $url
                return @{
                    'output'=($output -join "`n");
                }
            }.GetNewClosure(),
            ${function:Download-Data},
            {
                WriteLog "$($instance.GetPluginName()) : $args" -Background ([Drawing.Color]::LightBlue)
            }.GetNewClosure(),
            ${function:Exists-File},
            ${function:Serialize},
            ${function:Deserialize},
            ${function:Get-PackageInfoFromWheelName},
            ${function:Test-CanInstallPackageTo},
            { param([string] $name) return $global:autoCompleteIndex.Contains($name) }.GetNewClosure(),
            ${function:Prepare-PackageAutoCompletion},
            ${function:Recode})
        [void] $global:plugins.Add($instance)
        [void] $global:packageTypes.AddRange($instance.GetSupportedPackageTypes())
        WriteLog $instance.GetDescription()
    }
}

Function global:Start-Main([switch] $HideConsole, [switch] $Debug) {
    $env:PYTHONIOENCODING="UTF-16"
    $env:LC_CTYPE="UTF-16"

    if (-not $Debug) {
       Set-StrictMode -Off
       Set-PSDebug -Off
    } else {
       Set-StrictMode -Version latest
       Set-PSDebug -Strict -Trace 0  # -Trace ∈ (0, 1=lines, 2=lines+vars+calls)

       $ignoredExceptions = @(
           [System.IO.FileNotFoundException],
           [System.Management.Automation.MethodInvocationException],
           [System.Management.Automation.ParameterBindingException],
           [System.Management.Automation.ItemNotFoundException],
           [System.Management.Automation.ValidationMetadataException],
           [System.Management.Automation.DriveNotFoundException],
           [System.Management.Automation.CommandNotFoundException],
           [System.NotSupportedException],
           [System.ArgumentException],
           [System.Management.Automation.PSNotSupportedException]
       )

       $exceptionsWithScriptBacktrace = @(
           [System.Management.Automation.RuntimeException],
           [System.Management.Automation.PSInvalidCastException],
           [System.Management.Automation.PipelineStoppedException]
       )

       $appExceptionHandler = {
            param($Exception, $ScriptStackTrace = $null)

            if ($Exception.GetType() -in $ignoredExceptions) {
                return
            }

            $color = Get-Random -Maximum 16
            $color = @{
                BackgroundColor=$color;
                ForegroundColor=(15 - $color);
            }

            Write-Host ('=' * 70) @color
            Write-Host $Exception.GetType() @color
            Write-Host ('-' * 70) @color
            Write-Host $Exception.Message @color
            Write-Host ('-' * 70) @color

            if (($Exception.GetType() -in $exceptionsWithScriptBacktrace) -or
                ($Exception.GetType().BaseType -in $exceptionsWithScriptBacktrace)) {
                $ScriptStackTrace = $Exception.ErrorRecord.ScriptStackTrace
                if (-not [string]::IsNullOrWhiteSpace($ScriptStackTrace)) {
                    Write-Host $ScriptStackTrace @color
                    Write-Host ('-' * 70) @color
                }
            }

            Write-Host $Exception.StackTrace @color
            Write-Host ('=' * 70) @color

            Show-ConsoleWindow $true
       }.GetNewClosure()

       $firstChanceExceptionHandler = New-RunspacedDelegate ([ System.EventHandler`1[System.Runtime.ExceptionServices.FirstChanceExceptionEventArgs]] {
            param($Sender, $EventArgs)
            $null = $appExceptionHandler.Invoke($EventArgs.Exception)
       }.GetNewClosure())

       $unhandledExceptionHandler = New-RunspacedDelegate ([UnhandledExceptionEventHandler] {
            param($Sender, $EventArgs)
            $null = $appExceptionHandler.Invoke($EventArgs.Exception)
       }.GetNewClosure())

       [System.AppDomain]::CurrentDomain.Add_FirstChanceException($firstChanceExceptionHandler)
       [System.AppDomain]::CurrentDomain.Add_UnhandledException($unhandledExceptionHandler)
    }

    if ($HideConsole) {
        Show-ConsoleWindow $false
    }

    Load-PipsSettings
    Load-Plugins

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = Generate-Form
    $form.Add_Closing({

        foreach ($plugin in $global:plugins) {
            $plugin.Release()
        }

        Save-PipsSettings

        [System.Windows.Forms.Application]::Exit()
        if (-not [Environment]::UserInteractive) {
            Stop-Process $pid
        }
    })

    $form.Show()
    $form.Activate()

    $appContext = New-Object System.Windows.Forms.ApplicationContext
    $null = [System.Windows.Forms.Application]::Run($appContext)
}
