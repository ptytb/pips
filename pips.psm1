$PSDefaultParameterValues['*:Encoding'] = 'UTF8'
$null = Set-StrictMode -Version latest

$global:FRAMEWORK_VERSION = [version]([Runtime.InteropServices.RuntimeInformation]::FrameworkDescription -replace '^.[^\d.]*','')

$global:PIPS_SPELLING_PIPE = 'pips_spelling_server'
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
$global:logView = $null
$global:inputFilter = $null
$global:actionsModel = $null
$global:dataModel = $null  # [DataRow] keeps actual rows for the table of packages
$global:header = ("Select", "Package", "Installed", "Latest", "Type", "Status")
$global:csv_header = ("Package", "Installed", "Latest", "Type", "Status")
$global:search_columns = ("Select", "Package", "Installed", "Description", "Type", "Status")
$global:outdatedOnly = $true
$global:interpreters = $null
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

enum MainFormModes {
    Default;
    AltModeA;
    AltModeB;
    AltModeC;
}

$global:PlatformKeysForModes = @{
    ([MainFormModes]::Default)=$null;
    ([MainFormModes]::AltModeA)='Shift';
    ([MainFormModes]::AltModeB)='Alt';
    ([MainFormModes]::AltModeC)='Control';
}

Function global:GetAlternativeMainFormMode($keys) {
    $mode = [MainFormModes]::Default
    if ($keys.Shift) {
        $mode = [MainFormModes]::AltModeA
    } elseif ($keys.Alt) {
        $mode = [MainFormModes]::AltModeB
    } elseif ($keys.Control) {
        $mode = [MainFormModes]::AltModeC
    }
    return $mode
}

[AppMode] $global:APP_MODE = [AppMode]::Idle

$global:packageTypes = [System.Collections.ArrayList]::new()
$global:packageTypes.AddRange(@('pip', 'conda', 'git', 'wheel', 'https'))
$Global:PyPiPackageJsonCache = New-Object 'System.Collections.Generic.Dictionary[string,PSCustomObject]'

$global:PIP_DEP_TREE_LEGEND = "
Tree legend:
* = Extra package
x = Package doesn't exist in index
∞ = Dependency loop found
"

$global:ActionCommands = @{
    common=@{
        documentation  = @{ Command={ (ShowDocView $package).Show() } };
        copy_reqs      = @{ Command={ Set-Clipboard (($package -split ' ') -join [Environment]::NewLine) ; WriteLog "Copied $count items to clipboard." } };
    };
    other=@{
        files          = @{ Command= { Get-ChildItem -Recurse "$(py 'SitePackagesDir')\$package" | ForEach-Object { WriteLog $_.FullName } } };
    };
    pip=@{
        info           = @{ Command='PythonExe'; Args={ ('-m', 'pip', 'show', $package)                 }; Validate={ $exitCode -eq 0 }; };
        files          = @{ Command='PythonExe'; Args={ ('-m', 'pip', 'show', '--files', $package)      }; Validate={ $exitCode -eq 0 }; };
        update         = @{ Command='PythonExe'; Args={ ('-m', 'pip', 'install', '-U', $package)        }; Validate={ $output -match "Successfully installed (?:[^\s]+\s+)*$package" } };
        install        = @{ Command='PythonExe'; Args={ ('-m', 'pip', 'install', $package)              }; Validate={ } };
        install_nodeps = @{ Command='PythonExe'; Args={ ('-m', 'pip', 'install', '--no-deps', $package) }; Validate={ } };
        download       = @{ Command='PythonExe'; Args={ ('-m', 'pip', 'download', $package)             }; Validate={ } };
        uninstall      = @{ Command='PythonExe'; Args={ ('-m', 'pip', 'uninstall', '--yes', $package)   }; Validate={ } };
        deps_reverse   = @{ Command={ GetReverseDependencies }; }
        deps_tree      = @{ Command={ WriteLog $PIP_DEP_TREE_LEGEND ; WriteLog (GetDependencyAsciiGraph $package) } };
    };
    conda=@{
        info          = @{ Command='CondaExe'; Args= { ('list', '--prefix', (py 'Path'), '-v', '--json', $package) } };
        files         = @{ Command={
            $path = "$(py 'Path')\conda-meta"
            $query = "$package*.json"
            $file = Get-ChildItem -Path $path -Name $query -Depth 0 -File
            $json = Get-Content -Raw "$path\$file" | ConvertFrom-Json
            WriteLog ($json.files -join ([Environment]::NewLine))
        } };
        update        = @{ Command='CondaExe'; Args={ ('update', '--prefix', (py 'Path'), '--yes', '-q', $package) } };
        install       = @{ Command='CondaExe'; Args={ ('install', (Get-PipsSetting 'CondaChannels' -AsArgs -First), '--prefix', (py 'Path'), '--yes', '-q', '--no-shortcuts', $package) } };
        install_dry   = @{ Command='CondaExe'; Args={ ('install', (Get-PipsSetting 'CondaChannels' -AsArgs -First), '--prefix', (py 'Path'), '--dry-run', $package) } };
        install_nodeps= @{ Command='CondaExe'; Args={ ('install', (Get-PipsSetting 'CondaChannels' -AsArgs -First), '--prefix', (py 'Path'), '--yes', '-q', '--no-shortcuts', '--no-deps', '--no-update-dependencies', $package) } };
        uninstall     = @{ Command='CondaExe'; Args={ ('uninstall', '--prefix', (py 'Path'), '--yes', $package) } };
        deps_reverse  = @{ Command='CondaExe'; Args={ ('search', <#'--json',#> '--reverse-dependency', $package) }; <# PostprocessOutput={
            WriteLog ($output `
                | ConvertFrom-Json `
                | Get-Member -Type NoteProperty `
                | Select-Object -ExpandProperty Name)
        } #> };
    }
}

$global:ActionCommandInheritance = @{
    common=('pip', 'conda');
    pip=('wheel', 'sdist', 'builtin', 'other', 'git', 'https');
}

Function global:GetActionCommand($type, $command) {
    if ($global:ActionCommands.ContainsKey($type) -and $global:ActionCommands.Item($type).ContainsKey($command)) {
        return ,$global:ActionCommands.Item($type).Item($command)
    } else {
        if ($type -eq 'common') {
            return $null
        }
        foreach ($pair in $global:ActionCommandInheritance.GetEnumerator()) {
            $nextType, $fallbackFrom = $pair.Key, $pair.Value
            if ($type -in $fallbackFrom) {
                return ,(GetActionCommand $nextType $command)
            }
        }
    }
}


Function global:MakeEvent([hashtable] $properties) {
    [EventArgs] $EventArgs = [EventArgs]::new()  # Always use new(), not [EventArgs]::Empty !
    foreach ($p in $properties.GetEnumerator()) {
        $null = Add-Member -InputObject $EventArgs -Force -MemberType NoteProperty -Name $p.Key -Value $p.Value
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

Function global:ApplyAsync([object] $FunctionContext, [object] $Queue, [Func[object, object, System.Threading.Tasks.Task]] $Function, [delegate] $Finally) {

    $iterator = New-RunspacedDelegate([Action[System.Threading.Tasks.Task, object]] {
        param([System.Threading.Tasks.Task] $task, [object] $locals)
        $queue = $locals.queue
        if ($queue.Count -gt 0) {
            $element = $queue.Dequeue()
            $taskFromFunction = $locals.function.Invoke($element, $locals.FunctionContext)
            $null = $taskFromFunction.ContinueWith($locals.iterator, $locals,
                [System.Threading.CancellationToken]::None,
                ([System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously -bor
                    [System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent -bor
                    [System.Threading.Tasks.TaskContinuationOptions]::PreferFairness),
                $global:UI_SYNCHRONIZATION_CONTEXT)
        } else {
            $null = $task.ContinueWith($locals.Finally, $locals.FunctionContext,
                [System.Threading.CancellationToken]::None,
                ([System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously -bor
                    [System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent),
                $global:UI_SYNCHRONIZATION_CONTEXT)
        }
    })

    $locals = @{
        functionContext=$FunctionContext;
        queue=$Queue;
        function=$Function;
        finally=$Finally;
        iterator=$iterator;
    }

    $t = [System.Threading.Tasks.Task]::FromResult(@{})
    $null = $t.ContinueWith($iterator, $locals,
        [System.Threading.CancellationToken]::None,
        ([System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously -bor
            [System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent),
            $global:UI_SYNCHRONIZATION_CONTEXT)
}

Function global:InvokeWithContext([scriptblock] $function, [hashtable] $functions, [hashtable] $variables, [object[]] $arguments) {
    $vs = [System.Collections.Generic.List[psvariable]]::new()
    $null = $variables.GetEnumerator().ForEach({ $vs.Add([psvariable]::new($_.Key, $_.Value)) })
    $fs = [System.Collections.Generic.Dictionary[string,ScriptBlock]]::new()
    $null = $functions.GetEnumerator().ForEach({ $fs[$_.Key] = $_.Value })
    return $function.InvokeWithContext($fs, $vs, $arguments)
}

Function global:BindWithContext([scriptblock] $function, [hashtable] $functions, [hashtable] $variables, [object[]] $arguments) {
    $vs = [System.Collections.Generic.List[psvariable]]::new()
    $null = $variables.GetEnumerator().ForEach({ $vs.Add([psvariable]::new($_.Key, $_.Value)) })
    $fs = [System.Collections.Generic.Dictionary[string,ScriptBlock]]::new()
    $null = $functions.GetEnumerator().ForEach({ $fs[$_.Key] = $_.Value })
    return { return $function.InvokeWithContext($fs, $vs, ($arguments + $args)) }.GetNewClosure()
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
[DllImport("user32.dll")]public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, int lParam);
[DllImport("user32.dll")]public static extern int PostMessage(IntPtr hWnd, int uMsg, int wParam, int lParam);
[DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)] public static extern short GetAsyncKeyState(int virtualKeyCode);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
'
$global:WinAPI = Add-Type -MemberDefinition $MemberDefinition -Name 'WinAPI' -PassThru
${function:global:SendMessage} = { return $global:WinAPI::SendMessage.Invoke($args) }
${function:global:PostMessage} = { return $global:WinAPI::PostMessage.Invoke($args) }


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


$null = Add-Type -Name TerminateGracefully -Namespace Console -MemberDefinition @'
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


Function CheckPipsAlreadyRunning {
    [CmdletBinding()]
    param()

    $pips_pipe_instance = [System.IO.Directory]::GetFiles("\\.\\pipe\\") | Where-Object { $_ -match "$PIPS_SPELLING_PIPE"}
    if ($pips_pipe_instance) {
        [System.Windows.Forms.MessageBox]::Show(
            "There's another pips instance running, exiting.",
            "pips",
            [System.Windows.Forms.MessageBoxButtons]::OK)
    }
    return [bool] $pips_pipe_instance
}

Function StartPipsSpellingServer {
    $startServer = New-RunspacedDelegate ( [Func[Object]] {
        Write-Information "Start spellchecker on pipe \\.\pipe\$PIPS_SPELLING_PIPE"
        Start-Process -WindowStyle Hidden -FilePath powershell -ArgumentList "-ExecutionPolicy Bypass $PSScriptRoot\pips-spelling-server.ps1"
        Write-Information 'Server started.'
    });
    $task = [System.Threading.Tasks.Task[Object]]::new($startServer);
    $continuation = New-RunspacedDelegate ( [Action[System.Threading.Tasks.Task[Object]]] {
        Write-Information 'Connecting.'

        $pipe = $null
        while (-not $pipe -or -not $pipe.IsConnected) {
            $pipe = [System.IO.Pipes.NamedPipeClientStream]::new("\\.\pipe\$PIPS_SPELLING_PIPE");

            if ($pipe) {
                $milliseconds = 250
                try {
                    $pipe.Connect($milliseconds);
                } catch {
                }
            } else {
                Write-Information 'Waiting for spellchecker pipe...'
            }
        }
        Write-Information 'Connected!'
        $Global:sw = new-object System.IO.StreamWriter($pipe);
        $Global:sr = new-object System.IO.StreamReader($pipe);
        $Global:sw.AutoFlush = $false
        [bool] $Global:SuggestionsWorking = $false
    });

    $null = $task.ContinueWith($continuation, [System.Threading.CancellationToken]::None, (
        [System.Threading.Tasks.TaskContinuationOptions]::DenyChildAttach -bor
        [System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously),
        [System.Threading.Tasks.TaskScheduler]::Default)
    $task.Start([System.Threading.Tasks.TaskScheduler]::Default)
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

Function global:GuessEnvPath ($path, $fileName, [switch] $directory, [switch] $Executable) {
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

Function global:GetCurrentInterpreter($item, [switch] $Executable) {
    if (-not [string]::IsNullOrEmpty($item)) {
        $item = $Global:interpretersComboBox.SelectedItem."$item"
        if ($Executable) {
            $item = GetExistingFilePathOrNull $item
        }
        return $item
    } else {
        return $Global:interpretersComboBox.SelectedItem
    }
}

Function global:DeleteCurrentInterpreter() {
    if (-not (GetCurrentInterpreter 'User')) {
        WriteLog 'Can only delete venv which was added manually with env:Open or env:Create.'
        return
    }

    $path = GetCurrentInterpreter 'Path'
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
        WriteLog "Switching to '$(GetCurrentInterpreter 'Path')'"
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

Function global:GetExistingFilePathOrNull($path) {
    if (Exists-File $path) {
        return $path
    } else {
        return $null
    }
}

Function global:GetExistingPathOrNull($path) {
    if (Exists-Directory $path) {
        return $path
    } else {
        return $null
    }
}

Function SetWebClientWorkaround {
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

Function global:DownloadString($url, $ContinueWith = $null) {
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
            # $result = $wc.DownloadStringTaskAsync($url)
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


$global:_WritePipLogBacklog = [System.Collections.Generic.List[hashtable]]::new()
[int] $global:_LogViewEventMask = 0
[bool] $global:_LogViewHasBeenScrolledToEnd = $false
[bool] $global:_LogViewAutoScroll = $true
[System.Threading.SpinLock] $global:_LogViewCriticalSection = [System.Threading.SpinLock]::new($false)

Function global:WriteLogHelper {
    param(
        [object[]] $Lines,
        [bool] $UpdateLastLine,
        [bool] $NoNewline,
        [object] $Background,
        [object] $Foreground
    )

    [bool] $taken = $false
    try {
        $global:_LogViewCriticalSection.TryEnter([ref] $taken)

        $hidden = $global:form.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized

        if (-not $hidden) {
            $null = SendMessage $logView.Handle $WM_SETREDRAW 0 0
            $eventMask = SendMessage $logView.Handle $EM_SETEVENTMASK 0 0
        }

        $text = $Lines -join ' '

        # TextSelection rtb.Selection;
        $caretPosition = $logView.SelectionStart

        if ($UpdateLastLine) {
            $text = $text -replace "`r|`n",''

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
            $null = $logView.AppendText($text)
            $logTo = $logView.TextLength

            if (-not $NoNewline) {
                $logView.AppendText([Environment]::NewLine)
            }
        }

        if (($Background -ne $null) -or ($Foreground -ne $null)) {
            $logView.Select($logFrom, $logTo - $logFrom)
            if ($Background -ne $null) {
                $logView.SelectionBackColor = $Background
            }
            if ($Foreground -ne $null) {
                $logView.SelectionColor = $Foreground
            }
        }

        $logView.DeselectAll()

        if (-not $hidden) {
            $null = SendMessage $logView.Handle $WM_SETREDRAW 1 0
            $null = SendMessage $logView.Handle $EM_SETEVENTMASK 0 $eventMask
        }

        if ($global:_LogViewAutoScroll) {
            if (-not $hidden) {
                $textLength = $logView.TextLength
                $logView.Select($textLength, 0)
                $logView.Invalidate()
                $null = PostMessage $logView.Handle $WM_VSCROLL $SB_PAGEBOTTOM 0
            } else {
                $global:_LogViewHasBeenScrolledToEnd = $true
            }
        } else {
            $logView.Select($caretPosition, 0)
        }

    } finally {
        if ($taken) {
            $global:_LogViewCriticalSection.Exit($false)
        }
    }
}

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
        Background=$Background;
        Foreground=$Foreground;
    }

    if ($global:logView -eq $null) {
        [void] $global:_WritePipLogBacklog.Add($arguments)
        return
    }

    if ($logView.InvokeRequired) {
        $EventArgs = MakeEvent @{
            arguments=$arguments
        }
        $null = $logView.BeginInvoke($global:WritePipLogDelegate, ($logView, $EventArgs))
    } else {
        $null = global:WriteLogHelper @arguments
    }
}

Function global:ClearLog {
    $logView.Clear()
}

Function global:GetLogLength {
    return $logView.TextLength
}


Function Add-TopWidget($widget, $span=1) {
    $widget.Location = New-Object Drawing.Point $lastWidgetLeft,$lastWidgetTop
    $widget.size = New-Object Drawing.Point ($span*100-5),$widgetLineHeight
    $global:form.Controls.Add($widget)
    $Script:lastWidgetLeft = $lastWidgetLeft + ($span*100)
}

Function Add-HorizontalSpacer() {
    $Script:lastWidgetLeft = $lastWidgetLeft + 100
}

Function NewLine-TopLayout() {
    $Script:lastWidgetTop  = $Script:lastWidgetTop + $widgetLineHeight + 5
    $Script:lastWidgetLeft = 5
}

Function AddButton {
    [CmdletBinding()]
    param($name, $handler, [switch] $AsyncHandlers, [Parameter(Mandatory=$false)] [hashtable] $Modes)

    $button = [System.Windows.Forms.Button]::new()
    $button.Text = $name

    $button.Tag = if ($Modes) { $Modes } else { @{} }
    $button.Tag.Add([MainFormModes]::Default, @{ Click=$handler })

    if ($AsyncHandlers) {
        foreach ($transformationForMode in $button.Tag.GetEnumerator()) {
            $mode, $transformation = $transformationForMode.Key, $transformationForMode.Value
            $local:handler = $transformation.Item('Click')

            $wrappedHandler = New-RunspacedDelegate ([EventHandler] {
                param($Sender, $EventArgs)
                $widgetStateTransition = WidgetStateTransitionForCommandButton $button
                $doReverseWidgetState = [WidgetStateTransition]::ReverseAllAsync()
                $task = $handler.InvokeReturnAsIs($Sender, $EventArgs)
                $null = $task.ContinueWith($doReverseWidgetState, $widgetStateTransition,
                    [System.Threading.CancellationToken]::None,
                    ([System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously -bor
                        [System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent),
                        $global:UI_SYNCHRONIZATION_CONTEXT)
            }.GetNewClosure())

            $transformation.Remove('Click')
            $transformation.Add('Click', $wrappedHandler)
        }
    }

    $button.Add_Click($button.Tag.Item([MainFormModes]::Default).Item('Click'))

    if ($button.Tag.Count -gt 1) {
        $button.add_MouseEnter({
            param([System.Windows.Forms.Button] $Sender, [EventArgs] $EventArgs)
            $keys = ($Sender.Tag.GetEnumerator() `
                | Where-Object { $_.Key -ne [MainFormModes]::Default } `
                | ForEach-Object { [string]::Concat($PlatformKeysForModes[$_.Key], ' - ', $_.Value.Text) }) -join ', '
            $global:statusLabel.Text = "Alternative commands: hold $keys"
        })

        $button.add_MouseLeave({
            param($Sender, [EventArgs] $EventArgs)
            $global:statusLabel.Text = [string]::Empty
        })
    }

    Add-TopWidget $button
    return $button
}

Function AddButtonMenu ($text, $tools, $onclick) {
    $form = $global:form  # to be captured by $handler's closure

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
        $menuStrip.Show($global:form.PointToScreen($point))
    }.GetNewClosure()

    $button = AddButton $text $handler
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


Function AddButtons {
    $global:WIDGET_GROUP_COMMAND_BUTTONS = @(
        AddButton "Check Updates" ${function:GetPythonPackages} -AsyncHandlers -Modes @{
            ([MainFormModes]::AltModeA)=@{Text='Check conda'; Click={ ; [System.Threading.Tasks.Task]::FromResult(@{}) } }
            ([MainFormModes]::AltModeB)=@{Text='Check pip'; Click={ ; [System.Threading.Tasks.Task]::FromResult(@{}) } }
            };
        AddButton "List Installed" { GetPythonPackages($false) } -AsyncHandlers -Modes @{
            ([MainFormModes]::AltModeA)=@{Text='List only conda'; Click={ ; [System.Threading.Tasks.Task]::FromResult(@{}) } }
            ([MainFormModes]::AltModeB)=@{Text='List only pip'; Click={ ; [System.Threading.Tasks.Task]::FromResult(@{}) } }
            ([MainFormModes]::AltModeC)=@{Text='List w/o builtin'; Click={ ; [System.Threading.Tasks.Task]::FromResult(@{}) } }
            };
        AddButton "Sel All Visible" { SetVisiblePackageCheckboxes $true } -Modes @{
                ([MainFormModes]::AltModeA)=@{Text='Sel conda'; Click={ SetVisiblePackageCheckboxes $true @('conda') ; [System.Threading.Tasks.Task]::FromResult(@{}) }};
                ([MainFormModes]::AltModeB)=@{Text='Sel pip'; Click={ SetVisiblePackageCheckboxes $true @('pip') ; [System.Threading.Tasks.Task]::FromResult(@{}) }};
                ([MainFormModes]::AltModeC)=@{Text='Inverse'; Click={ SetVisiblePackageCheckboxes -Inverse ; [System.Threading.Tasks.Task]::FromResult(@{}) }};
            };
        AddButton "Select None" { SetAllPackageCheckboxes($false) } ;
        AddButton "Check Deps" ${function:CheckDependencies} -AsyncHandlers ;
        AddButton "Execute" ${function:ExecuteAction} -AsyncHandlers -Modes @{
                ([MainFormModes]::AltModeA)=@{Text='Show command'; Click={ ExecuteAction -ShowCommand }};
                ([MainFormModes]::AltModeB)=@{Text='Execute...'; Click={
                    $history = Get-Variable -Name CUSTOM_COMMAND_ARGUMENTS_HISTORY -Scope Global -ErrorAction SilentlyContinue -ValueOnly
                    if (-not $history) {
                        $history = [System.Collections.Generic.List[string]]::new()
                        $global:CUSTOM_COMMAND_ARGUMENTS_HISTORY = $history
                    }
                    $message = @'
Enter the arguments for python, for example: -m pip show -v

The list of packages will be appended to the end.
'@
                    $default = if ($history.Count -gt 0) { $history[-1] } else { '' }
                    $request = RequestUserString $message 'Custom arguments' $default $history
                    if ($request -eq $null) {
                        return [System.Threading.Tasks.Task]::FromException([Exception]::new('Cancelled by user'))
                    }
                    $history.Add($request)
                    ExecuteAction -CustomArguments ($request.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries))
                    }};
                ([MainFormModes]::AltModeC)=@{Text='Execute serial'; Click={ ExecuteAction -Serial }};
            };
    )
}

Function global:GetPyDoc($request) {
    $requestNormalized = $request -replace '-','_'
    $output = & (GetCurrentInterpreter 'PythonExe') -m pydoc $requestNormalized

    if ("$output".StartsWith('No Python documentation found')) {
        $output = & (GetCurrentInterpreter 'PythonExe') -m pydoc ($requestNormalized).ToLower()
    }

    # TODO: pass 'em to ProcessWithPipedIO
    # $env:PYTHONIOENCODING="UTF-8"
    # $env:LC_CTYPE="UTF-8"

    # $output = Recode ([Text.Encoding]::UTF8) ([Text.Encoding]::Unicode) $output

    return $output
}

Function global:GetPythonBuiltinPackagesAsync() {

    $delegate = New-RunspacedDelegate ([Func[object, object]] {
        param([object] $locals)

        $builtinLibs = [System.Collections.Generic.List[PSObject]]::new()
        $path = GetCurrentInterpreter 'Path'
        $libs = "${path}\Lib"
        $ignore = [regex] '^__'
        $filter = [regex] '\.py.?$'

        $trackDuplicates = [System.Collections.Generic.HashSet[String]]::new()

        foreach ($item in Get-ChildItem -Directory $libs) {
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
        $sys_builtin_module_names = & (GetCurrentInterpreter 'PythonExe') -c $getBuiltinsScript
        $modules = $sys_builtin_module_names.Split(',')
        foreach ($builtinModule in $modules) {
            if ($trackDuplicates.Contains("$builtinModule")) {
                continue
            }
            $null = $builtinLibs.Add([PSCustomObject] @{Package=$builtinModule; Type='builtin'})
        }

        return ,$builtinLibs
    })

    $token = [System.Threading.CancellationToken]::None
    $options = ([System.Threading.Tasks.TaskCreationOptions]::AttachedToParent)
    $taskGetBuiltinPackages = [System.Threading.Tasks.Task]::Factory.StartNew($delegate, @{}, $token, $options,
        $global:UI_SYNCHRONIZATION_CONTEXT)

    return $taskGetBuiltinPackages
}

Function global:GetPythonOtherPackagesAsync {
    $delegate = New-RunspacedDelegate([Func[object, object]] {
        param([object] $locals)

        $otherLibs = [System.Collections.Generic.List[PSCustomObject]]::new()
        $path = GetCurrentInterpreter 'Path'
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
    })

    $token = [System.Threading.CancellationToken]::None
    $options = ([System.Threading.Tasks.TaskCreationOptions]::AttachedToParent)
    $taskGetOtherPackages = [System.Threading.Tasks.Task]::Factory.StartNew($delegate, @{}, $token, $options,
        $global:UI_SYNCHRONIZATION_CONTEXT)

    return $taskGetOtherPackages
}

Function global:GetCondaJsonAsync([bool] $outdatedOnly) {
    $conda_exe = GetCurrentInterpreter 'CondaExe' -Executable
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
    $null = $arguments.Add((GetCurrentInterpreter 'Path'))

    $process = [ProcessWithPipedIO]::new($conda_exe, $arguments)
    $taskProcessDone = $process.StartWithLogging($false, $true)
    $task = $process.ReadOutputToEndAsync()
    return $task
}

Function global:GetCondaPackagesAsync([bool] $outdatedOnly) {

    $continuationParseJson = New-RunspacedDelegate([Func[System.Threading.Tasks.Task, object]] {
        param([System.Threading.Tasks.Task] $task)

        $condaPackages = [System.Collections.Generic.List[PSCustomObject]]::new()

        trap {
            return ,$condaPackages
        }

        [string[]] $JsonList = $task.Result
        $outdatedOnly = $JsonList.Length -eq 2

        $installed = @{}
        $items = $JsonList[0] | ConvertFrom-Json

        foreach ($item in $items) {
            if (-not $outdatedOnly) {
                $null = $condaPackages.Add([PSCustomObject] @{Type='conda'; Package=$item.name; 'Installed'=$item.version})
            } else {
                $null = $installed.Add($item.name, $item.version)
            }
        }

        if ($outdatedOnly) {
            $items = $JsonList[1] | ConvertFrom-Json
            foreach ($_ in $items.PSObject.Properties.GetEnumerator()) { # ForEach-Object doesn't work here
                $name = $_.Name
                $archives = $_.Value

                if (($archives.Count -eq 0) -or (-not $installed.Contains($name))) {
                    continue
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
    })

    $tasks = [System.Collections.Generic.List[System.Threading.Tasks.Task[string]]]::new()
    $taskGetInstalledJson = GetCondaJsonAsync $false
    $null = $tasks.Add($taskGetInstalledJson)

    if ($outdatedOnly) {  # get updated versions and keep only those newer than installed
        $taskGetUpdatedJson = GetCondaJsonAsync $true
        $null = $tasks.Add($taskGetUpdatedJson)
    }

    $taskGetJsonList = [System.Threading.Tasks.Task[string]]::WhenAll($tasks)
    $taskProcessJson = $taskGetJsonList.ContinueWith($continuationParseJson,
            [System.Threading.CancellationToken]::None,
            ([System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent -bor
                [System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously),
            $global:UI_SYNCHRONIZATION_CONTEXT)

    return $taskProcessJson
}

Function Make-PipActionItem($name, $code, $validator, $takesList = $false, $id = $null) {
    $action = New-Object psobject -Property @{Name=$name; TakesList=$takesList; Index=(++$Script:actionItemCount);}
    $action | Add-Member -MemberType ScriptMethod -Name Execute  -Value $code -Force
    $action | Add-Member -MemberType ScriptMethod -Name Validate -Value $validator -Force
    $action | Add-Member -MemberType NoteProperty -Name Id       -Value $id -Force
    return $action
}

Function Add-ComboBoxActions {
    $actionsModel = New-Object System.Collections.ArrayList
    Set-Variable -Name index -Value 1 -Option AllScope
    Function AddAction {
        <#
        .PARAMETER Bulk
        Bulked command processes all the packages of the same type at once, for each of all distinct package types (pip, conda, ...)
        .PARAMETER AllTypes
        Command accepts all package types
        #>
        param([hashtable] $actionProperties, [switch] $Bulk, [switch] $AllTypes, [switch] $Singleton)
        $action = New-Object PSObject -Property $actionProperties
        $action | Add-Member -MemberType NoteProperty -Name Bulk -Value $Bulk -Force
        $action | Add-Member -MemberType NoteProperty -Name AllTypes -Value $AllTypes -Force
        $action | Add-Member -MemberType NoteProperty -Name Index -Value $index -Force
        $action | Add-Member -MemberType NoteProperty -Name Singleton -Value $Singleton -Force
        $action | Add-Member -MemberType ScriptMethod -Name ToString -Value { "$($this.Name) [F$($this.Index)]" } -Force
        $null = $actionsModel.Add($action)
        ++$index
    }

    AddAction @{ Name='Show information'; Id='info'; }
    AddAction @{ Name='Show documentation'; Id='documentation'; } -Singleton
    AddAction @{ Name='Show dependency tree'; Id='deps_tree'; }
    AddAction @{ Name='Show dependent packages'; Id='deps_reverse'; }
    AddAction @{ Name='Update'; Id='update'; } -Bulk
    AddAction @{ Name='Install'; Id='install'; } -Bulk
    AddAction @{ Name='Install without dependencies'; Id='install_nodeps'; } -Bulk
    AddAction @{ Name='Uninstall'; Id='uninstall'; } -Bulk
    AddAction @{ Name='Install (dry run)'; Id='install_dry'; } -Bulk
    AddAction @{ Name='Download'; Id='download'; } -Bulk
    AddAction @{ Name='Copy as requirements'; Id='copy_reqs'; } -Bulk -AllTypes
    AddAction @{ Name='List files'; Id='files'; }

    $global:actionsModel = $actionsModel

    $actionListComboBox = New-Object System.Windows.Forms.ComboBox
    $actionListComboBox.DataSource = $actionsModel
    $actionListComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    Add-TopWidget $actionListComboBox 2.0

    return $actionListComboBox

    & $Add (Make-PipActionItem 'Install' {
            param($pkg,$type,$version)
            $git_url = Validate-GitLink $pkg
            if ($git_url) {
                $pkg = $git_url
            }
            if ((-not [string]::IsNullOrEmpty($version)) -and ($type -notin @('git', 'wheel'))) {  # as git version is a timestamp; wheel version is in the file name
                $pkg = "$pkg==$version"
            }
            $ActionCommands[$type].install.Invoke($pkg) } `
        { param($pkg,$out); ($out -match "Successfully installed (?:[^\s]+\s+)*$pkg") } )
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

    $python = GuessEnvPath $path 'python' -Executable
    if (-not $python) {
        return
    }

    Write-Information "Found python at $python"

    # Maybe in future; let it be KISS for now...
    # $process = [ProcessWithPipedIO]::new()
    # $taskProcessDone = $process.Start()
    # $taskReadOutput = $process.ReadOutputToEndAsync()
    # $taskReadError = $process.ReadErrorToEndAsync()
    # $allTasks = [System.Collections.Generic.List[System.Threading.Tasks.Task[string]]]::new()
    # $gathered = [System.Threading.Tasks.Task[string]]::WhenAll($allTasks)
    # $taskAllDone = $gathered.ContinueWith()

    $versionString = (& $python --version 2>&1).ToString()  # redirection from stderr produces Error object instead of string

    $version = [regex]::Match($versionString, '\s+(\d+\.\d+)').Groups[1]
    $arch = (Test-is64Bit $python).FileType

    $action = New-Object PSObject -Property @{
        Path                 = $path;
        Version              = "$version";
        Arch                 = $arch;
        Bits                 = @{"x64"="64"; "x86"="32";}[$arch];
        PythonExe            = $python;
        PipExe               = GuessEnvPath $path 'pip' -Executable;
        CondaExe             = GuessEnvPath $path 'conda' -Executable;
        VirtualenvExe        = GuessEnvPath $path 'virtualenv' -Executable;
        VenvActivate         = GuessEnvPath $path 'activate' -Executable;
        PipenvExe            = GuessEnvPath $path 'pipenv' -Executable;
        RequirementsTxt      = GuessEnvPath $path 'requirements.txt';
        Pipfile              = GuessEnvPath $path 'Pipfile';
        PipfileLock          = GuessEnvPath $path 'Pipfile.lock';
        SitePackagesDir      = GuessEnvPath $path 'Lib\site-packages' -directory;
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

Function global:PreparePackageAutoCompletion {

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

Function global:CreateInstallForm {
    PreparePackageAutoCompletion

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

        $null = PostMessage $cb.Handle $WM_CHAR 0x20 0  # write space...
        $null = PostMessage $cb.Handle $WM_CHAR $VK_BACKSPACE 0  # and erase it to trigger completion pop-up
    })

    $FuncSetAutoCompleteMode = [EventHandler] {
        param($Sender, $EventArgs)

        [InstallAutoCompleteMode] $mode = $EventArgs.Mode

        # Write-Information "ACTIVATING MODE !!! $mode"

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
                # Write-Information '<###### 1>=' $cb.Handle
                $cb.AutoCompleteCustomSource = $global:autoCompleteIndex
                $cb.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::Suggest
                $cb.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource
                # Write-Information '<###### 2>=' $cb.Handle
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
                $null = DownloadString $jsonUrl -ContinueWith {
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
                # Write-Information 'PP='$possibleDirectoryPath
                $autoCompleteWheels = [System.Windows.Forms.AutoCompleteStringCollection]::new()
                $wheelFiles = Get-ChildItem -Path $possibleDirectoryPath -Filter '*.whl' -File -Depth 0
                foreach ($wheel in $wheelFiles) {
                    # Write-Information 'WHEEL=' $($wheel.Name)
                    [void] $autoCompleteWheels.Add("$possibleDirectoryPath$($wheel.Name)")
                }

                $cb.AutoCompleteCustomSource = $autoCompleteWheels
                $cb.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
                $cb.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource

                $null = PostMessage $cb.Handle $WM_CHAR 0x20 0  # write space...
                $null = PostMessage $cb.Handle $WM_CHAR $VK_BACKSPACE 0  # and erase it to trigger completion pop-up
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

            #Write-Information "old='$oldVersion', new='$version'"

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
        # Write-Information 'MODE GUESS=' $guessedCompletionMode
        if ($guessedCompletionMode -ne $form.currentInstallAutoCompleteMode) {
            # Write-Information 'MODE WAS=' $form.currentInstallAutoCompleteMode 'CHANGE TO=' $guessedCompletionMode

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
        SelectPipAction 'Install'
        ExecuteAction
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
    $form.Icon = $global:form.Icon

    $null = $form.ShowDialog()
}

Function global:RequestUserString($message, $title, $default, $completionItems = $null, [ref] $ControlKeysState) {
    $Form                            = New-Object system.Windows.Forms.Form
    $Form.ClientSize                 = '421,247'
    $Form.text                       = $title
    $Form.TopMost                    = $false
    $Form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Hide
    $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $Form.MinimizeBox = $false
    $Form.MaximizeBox = $false
    $Form.Icon = $global:form.Icon

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

Function global:CreateSearchForm {
    $pluginNames = ($global:plugins | ForEach-Object { $_.GetPluginName() }) -join ', '
    $message = "Enter keywords to search with PyPi, Conda, Github and plugins: $pluginNames`n`nChecked items will be kept in the search list"
    $title = "pip, conda, github search"
    $default = ""
    $input = RequestUserString $message $title $default
    if (-not $input) {
        return
    }

    WriteLog "Searching for $input"
    WriteLog 'Double click or [Ctrl+Enter] a table row to open a package home page in browser'
    $stats = Get-SearchResults $input
    WriteLog "Found $($stats.Total) packages: $($stats.PipCount) pip, $($stats.CondaCount) conda, $($stats.GithubCount) github, $($stats.PluginCount) from plugins. Total $($global:dataModel.Rows.Count) packages in list."
    WriteLog
}

Function global:InitPackageGridViewProperties() {
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

Function global:InitPackageUpdateColumns($dataTable) {
    $dataTable.Columns.Clear()
    foreach ($c in $global:header) {
        if ($c -eq "Select") {
            $column = New-Object System.Data.DataColumn $c,([bool])
        } else {
            $column = New-Object System.Data.DataColumn $c,([string])
            $column.ReadOnly = $true
        }
        $dataTable.Columns.Add($column)
    }
}

Function global:InitPackageSearchColumns($dataTable) {
    $dataTable.Columns.Clear()
    foreach ($c in $global:search_columns) {
        if ($c -eq "Select") {
            $column = New-Object System.Data.DataColumn $c,([bool])
        } else {
            $column = New-Object System.Data.DataColumn $c,([string])
            $column.ReadOnly = $true
        }
        $dataTable.Columns.Add($column)
    }
}

Function global:HighlightPackages {
    # $global:outdatedOnly is needed because when row filter changes, we need to colorize rows again
    if ($outdatedOnly) {
        return
    }

    $dataGridView.BeginInit()
    foreach ($row in $dataGridView.Rows) {
        $color = switch ($row.DataBoundItem.Row.Type) {
            'builtin' { [Drawing.Color]::LightGreen }
            'other' { [Drawing.Color]::LightPink }
            'conda' { [Drawing.Color]::LightGoldenrodYellow }
            Default { $null }
        }
        if ($color) {
            $row.DefaultCellStyle.BackColor = $color
        }
    }
    $dataGridView.EndInit()
}

Function global:OpenLinkInBrowser($url) {
    if ((-not [string]::IsNullOrWhiteSpace($url)) -and ($url -match '^https?://')) {
        $url = [System.Uri]::EscapeUriString($url)
        Start-Process -FilePath $url
    }
}


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
    $json = DownloadString($pypi_json_url -f $packageName)
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
    $path = GetExistingPathOrNull $path
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
                $output = & (GetCurrentInterpreter 'VirtualenvExe') --python="$(GetCurrentInterpreter 'PythonExe')" $path 2>&1
                return $output
            };
            IsAccessible = { [bool] (GetCurrentInterpreter 'VirtualenvExe' -Executable) };
        };
        @{
            MenuText = 'with pipenv';
            Code = {
                param($path)
                $env:PIPENV_VENV_IN_PROJECT = 1
                Set-Location -Path $path
                $output = & (GetCurrentInterpreter 'PipenvExe') --python "$(GetCurrentInterpreter 'Version')" install 2>&1
                return $output
            };
            IsAccessible = { [bool] (GetCurrentInterpreter 'PipenvExe' -Executable) };
        };
        @{
            MenuText = 'with conda';
            Code = {
                param($path)
                $version = GetCurrentInterpreter 'Version'
                $output = & (GetCurrentInterpreter 'CondaExe') create -y -q --prefix $path python=$version 2>&1
                return $output
            };
            IsAccessible = { [bool] (GetCurrentInterpreter 'CondaExe' -Executable) };
        };
        @{
            NoTargetPath = $true;
            MenuText = '(tool required) Install virtualenv';
            Code = {
                param($path)
                $output = & (GetCurrentInterpreter 'PythonExe') -m pip install virtualenv 2>&1
                return $output
            };
            IsAccessible = { -not [bool] (GetCurrentInterpreter 'VirtualenvExe') };
        };
        @{
            NoTargetPath = $true;
            MenuText = '(tool required) Install pipenv';
            Code = {
                param($path)
                $output = & (GetCurrentInterpreter 'PythonExe') -m pip install pipenv 2>&1
                return $output
            };
            IsAccessible = { -not [bool] (GetCurrentInterpreter 'PipenvExe') };
        };
        @{
            NoTargetPath = $true;
            MenuText = '(tool required) Install conda';
            Code = {
                param($path)
                # menuinst, cytoolz are required by conda to run
                $menuinst = Validate-GitLink "https://github.com/ContinuumIO/menuinst@1.4.8"
                $output_0 = & (GetCurrentInterpreter 'PythonExe') -m pip install $menuinst 2>&1
                $output_1 = & (GetCurrentInterpreter 'PythonExe') -m pip install cytoolz conda 2>&1
                $CondaExe = GuessEnvPath (GetCurrentInterpreter 'Path') 'conda' -Executable

                # conda needs a little caress to run together with pip
                $path = (GetCurrentInterpreter 'Path')
                $stub = "$path\Lib\site-packages\conda\cli\pip_warning.py"
                $main = "$path\Lib\site-packages\conda\cli\main.py"
                Move-Item $stub "${stub}_"
                Copy-Item $main $stub
                # New-Item -Path $stub -ItemType SymbolicLink -Value $main

                $record = GetCurrentInterpreter
                $record.CondaExe = $CondaExe
                return @($output_0, $output_1) | ForEach-Object { $_ }
            };
            IsAccessible = { -not [bool] (GetCurrentInterpreter 'CondaExe' -Executable) };
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
        $pythonExe = (GetCurrentInterpreter 'PythonExe')
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
        return (GetCurrentInterpreter 'Version')
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

    $buttonEnvCreate = AddButtonMenu 'env: Create' $tools $menuclick
    return $buttonEnvCreate
}

Function Add-EnvToolButtonMenu {
    $menu = @(
        @{
            Persistent = $true;
            MenuText = 'Python REPL';
            Code = { Start-Process -FilePath (GetCurrentInterpreter 'PythonExe') -WorkingDirectory (GetCurrentInterpreter 'Path') };
        };
        @{
            MenuText = 'Shell with Virtualenv Activated';
            Code = { Start-Process -FilePath cmd.exe -WorkingDirectory (GetCurrentInterpreter 'Path') -ArgumentList "/K $(GetCurrentInterpreter 'VenvActivate')" };
            IsAccessible = { (GetCurrentInterpreter 'VenvActivate') };
        };
        @{
            Persistent = $true;
            MenuText = 'Open IDLE'
            Code = { Start-Process -FilePath (GetCurrentInterpreter 'PythonExe') -WorkingDirectory (GetCurrentInterpreter 'Path') -ArgumentList '-m idlelib.idle' -WindowStyle Hidden };
        };
        @{
            MenuText = 'pipenv shell'
            Code = { Start-Process -FilePath (Get-Bin 'pipenv.exe') -WorkingDirectory (GetCurrentInterpreter 'Path') -ArgumentList 'shell' };
            IsAccessible = { [bool] (Get-Bin 'pipenv.exe') -and [bool] (GetCurrentInterpreter 'Pipfile') };
        };
        @{
            Persistent = $true;
            MenuText = 'Environment variables...';
            Code = {
                $interpreter = GetCurrentInterpreter
                if ($interpreter.User) {
                    CreateFormEnvironmentVariables $interpreter
                } else {
                    $path = $interpreter.Path
                    WriteLog "$path is not a venv. Editing system-wide environment variables." -Background 'LightSalmon'
                    & rundll32 sysdm.cpl,EditEnvironmentVariables
                }
            };
        };
        @{
            Persistent = $true;
            MenuText = 'Open containing directory'
            Code = { Start-Process -FilePath 'explorer.exe' -ArgumentList "$(GetCurrentInterpreter 'Path')" };
        };
        @{
            Persistent = $false;
            MenuText = 'Remove environment...'
            Code = {
                DeleteCurrentInterpreter
            };
            IsAccessible = { [bool] (GetCurrentInterpreter 'User') };
        };
    )

    $menuclick = {
        param($item)
        $output = $item.Code.Invoke()
    }

    $buttonEnvTools = AddButtonMenu 'env: Tools' $menu $menuclick
    return $buttonEnvTools
}

Function global:GetPyDocTopics() {
    $pythonExe = GetCurrentInterpreter 'PythonExe'
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

Function global:GetPyDocApropos($request) {
    $pythonExe = GetCurrentInterpreter 'PythonExe'
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
                $topics = GetPyDocTopics
                $input = RequestUserString $message $title $default $topics
                if (-not $input) {
                    return
                }
                (ShowDocView $input).Show()
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
                $input = RequestUserString $message $title $default
                if (-not $input) {
                    return
                }
                WriteLog "Searching apropos for $input"
                $apropos = GetPyDocApropos $input
                if ($apropos -and $apropos.Count -gt 0) {
                    WriteLog "Found $($apropos.Count) topics"
                    $docView = ShowDocView -SetContent ($apropos -join "`n") -Highlight $Script:pyRegexNameChain -NoDefaultHighlighting
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
                $version = GetCurrentInterpreter 'Version'
                $arch = GetCurrentInterpreter 'Arch'
                $bits = GetCurrentInterpreter 'Bits'
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
                $list = RequestUserString "Enter conda channels, separated with space.`n`nLeave empty to use defaults.`n`nPopular channels: conda-forge anaconda defaults bioconda`n`nThe first channel is for installing, the others are for searching only." 'Edit conda channels' $channels
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
                    $cacheFolder = & (GetCurrentInterpreter 'PythonExe') -c $getCacheFolderScript_b10
                    if ([string]::IsNullOrWhiteSpace($cacheFolder)) {
                        $cacheFolder = & (GetCurrentInterpreter 'PythonExe') -c $getCacheFolderScript_a10
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
                ClearRows
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
                OpenLinkInBrowser "https://github.com/ptytb/pips/issues/new?title=$title&body=$hostInfo"
            };
        };
        @{
            Persistent=$true;
            MenuText = "Binge spawn processes";
            Code = {
                $delegate = New-RunspacedDelegate([Action[System.Threading.Tasks.Task, object]] {
                    $process = [ProcessWithPipedIO]::new('cat', @('D:\work\pyfmt-big.txt'))
                    # $process = [ProcessWithPipedIO]::new('py', @('--help'))
                    $taskProcessDone = $process.StartWithLogging($true, $true)
                    # $taskReadOutput = $process.ReadOutputToEndAsync()
                })
                $task = [System.Threading.Tasks.Task]::FromResult(0)
                while ($true) {
                    $options = ([System.Threading.Tasks.TaskContinuationOptions]::DenyChildAttach -bor `
                        [System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously)
                    $task = $task.ContinueWith($delegate, @{}, [System.Threading.CancellationToken]::None, $options,
                        [System.Threading.Tasks.TaskScheduler]::Default)
                    [System.Windows.Forms.Application]::DoEvents()
                }
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

    $toolsButton = AddButtonMenu 'Tools' $menu $menuclick
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
            OpenLinkInBrowser "$packageHomepageFromPlugin"
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
        OpenLinkInBrowser "${url}${urlName}"
    }
}

Function CreateMainForm {
    $form = New-Object Windows.Forms.Form
    $form.Text = "pips - python package browser"
    $form.Size = New-Object Drawing.Point 1125, 840
    $form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Show
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $form.StartPosition = 'CenterScreen'
    $form.Topmost = $false
    $form.KeyPreview = $true
    $form.Icon = Convert-Base64ToICO $iconBase64_Snakes
    $global:form = $form

    $null = AddButtons

    $actionListComboBox = Add-ComboBoxActions
    $global:actionListComboBox = $actionListComboBox

    $group = New-Object System.Windows.Forms.Panel
    $group.Location = New-Object System.Drawing.Point 502,2
    $group.Size = New-Object System.Drawing.Size 300,28
    $group.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $form.Controls.Add($group)

    $global:WIDGET_GROUP_INSTALL_BUTTONS = @(
        AddButton "Search..." ${function:CreateSearchForm} ;
        AddButton "Install..." ${function:CreateInstallForm} ;
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

        #Write-Information $query

        if ($searchText.Length -gt 0) {
            $global:dataModel.DefaultView.RowFilter = $query
        } else {
            $global:dataModel.DefaultView.RowFilter = $null
        }

        if ($selectedRow) {
            SetSelectedRow $selectedRow
        }

        HighlightPackages
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
    $buttonEnvOpen = AddButton "env: Open..." {
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

    $alternateFunctionality_WidgetStateTransition = [WidgetStateTransition]::new()

    $form.add_KeyDown({
        if ($Global:interpretersComboBox.Focused) {
            if (($_.KeyCode -eq 'C') -and $_.Control) {
                $python_exe = GetCurrentInterpreter 'PythonExe'
                Set-Clipboard $python_exe
                WriteLog "Copied to clipboard: $python_exe"
            }
            if ($_.KeyCode -eq 'Delete') {
                DeleteCurrentInterpreter
            }
        }

        if ($global:APP_MODE -eq [AppMode]::Idle) {
            $wst = $alternateFunctionality_WidgetStateTransition
            $mode = GetAlternativeMainFormMode $_
            [bool] $hasEntered = $false
            $null = $wst.EnterMode($mode, [ref] $hasEntered)
            if ($hasEntered) {
                foreach ($button in $WIDGET_GROUP_COMMAND_BUTTONS) {
                    [hashtable] $buttonModes = $button.Tag
                    if ($buttonModes.ContainsKey($mode)) {
                        [hashtable] $modeTransformation = $buttonModes[$mode]
                        $null = $wst.Add($button).Transform($modeTransformation)
                    }
                }
                $_.Handled = $true
            }
        }

    }.GetNewClosure())

    $form.add_KeyUp({
        if ($global:APP_MODE -eq [AppMode]::Idle) {
            $wst = $alternateFunctionality_WidgetStateTransition
            [bool] $activeMode = $false
            $mode = GetAlternativeMainFormMode $_
            $null = $wst.IsModeActive($mode, [ref] $activeMode)
            if (-not $activeMode) {
                $null = $wst.ReverseAll().ExitMode()
            }
        }
    }.GetNewClosure())

    $form.add_Deactivate({
        if ($global:APP_MODE -eq [AppMode]::Idle) {
            $wst = $alternateFunctionality_WidgetStateTransition
            $null = $wst.ReverseAll().ExitMode()
        }
    }.GetNewClosure())

    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $Script:dataGridView = $dataGridView
    $Global:dataGridView = $dataGridView

    $dataGridView.Location = New-Object Drawing.Point 7,($Script:lastWidgetTop + $Script:widgetLineHeight)
    $dataGridView.Size = New-Object Drawing.Point 800,450
    $dataGridView.ShowCellToolTips = $false
    $dataGridView.Add_Sorted({ HighlightPackages })

    $dataGridToolTip = New-Object System.Windows.Forms.ToolTip

    # $dataGridView.Add_CellMouseEnter({
    #     if (($_.RowIndex -gt -1)) {
    #         Update-PythonPackageDetails $_
    #         $text = $dataGridView.Rows[$_.RowIndex].Cells['Package'].ToolTipText
    #         $dataGridToolTip.RemoveAll()
    #         if (-not [string]::IsNullOrEmpty($text)) {
    #             $dataGridToolTip.InitialDelay = 50
    #             $dataGridToolTip.ReshowDelay = 10
    #             $dataGridToolTip.AutoPopDelay = [Int16]::MaxValue
    #             $dataGridToolTip.ShowAlways = $true
    #             $dataGridToolTip.SetToolTip($dataGridView, $text)
    #         }
    #     }
    # }.GetNewClosure())
    #
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

    InitPackageGridViewProperties

    $dataModel = New-Object System.Data.DataTable
    $dataGridView.DataSource = $dataModel
    $global:dataModel = $dataModel
    InitPackageUpdateColumns $dataModel

    $form.Controls.Add($dataGridView)

    $logView = $global:RichTextBox_t::new()
    $logView.Location = New-Object Drawing.Point 7,520
    $logView.Size = New-Object Drawing.Point 800,270
    $logView.ReadOnly = $true
    $logView.Multiline = $true
    $logView.Font = [System.Drawing.Font]::new('Consolas', 13)
    $form.Controls.Add($logView)

    $logView.add_HandleCreated({
        param($Sender)
        $null = [SearchDialogHook]::new($Sender)
        $null = [TextBoxNavigationHook]::new($Sender)

        $logView = $Sender
        $global:logView = $logView

        $global:WritePipLogDelegate = New-RunspacedDelegate ([EventHandler] {
            param($Sender, $EventArgs)
            $arguments = $EventArgs.arguments
            $null = WriteLogHelper @arguments
        })

        foreach ($arguments in $global:_WritePipLogBacklog) {
            WriteLogHelper @arguments
        }
        $_WritePipLogBacklog = $null
        Remove-Variable -Scope Global _WritePipLogBacklog
    })

    $logView.Add_LinkClicked({
        param($Sender, $EventArgs)
        OpenLinkInBrowser $EventArgs.LinkText
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

        if (Get-Member -inputobject $row -name "LogTo" -Membertype Properties) {
            $logView.SelectAll()
            $logView.SelectionBackColor = $logView.BackColor

            $logView.Select($row.LogFrom, $row.LogTo) # TODO: 2nd must be len
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

    $lastWidgetTop = $Script:lastWidgetTop

    $FuncResizeForm = {

        Write-Host "Main form has been resized."

        $hidden = $form.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized
        if ($hidden) {
            $null = SendMessage $logView.Handle $WM_SETREDRAW 0 0
            $global:_LogViewEventMask = SendMessage $logView.Handle $EM_SETEVENTMASK 0 0
        } else {
            $null = SendMessage $logView.Handle $WM_SETREDRAW 1 0
            $null = SendMessage $logView.Handle $EM_SETEVENTMASK 0 $global:_LogViewEventMask

            if ($global:_LogViewHasBeenScrolledToEnd) {
                $global:_LogViewHasBeenScrolledToEnd = $false
                $null = PostMessage $logView.Handle $WM_VSCROLL $SB_PAGEBOTTOM 0
            }
        }

        $dataGridView.Width = $form.ClientSize.Width - 15
        $dataGridView.Height = $form.ClientSize.Height / 2
        $logView.Top = $dataGridView.Bottom + 15
        $logView.Width = $form.ClientSize.Width - 15
        $logView.Height = $form.ClientSize.Height - $dataGridView.Bottom - $lastWidgetTop - 10
    }.GetNewClosure()

    $null = & $FuncResizeForm
    $form.Add_Resize($FuncResizeForm)
    $form.Add_Shown({
        WriteLog "`nHold Shift and hover the rows to fetch the detailed package info form PyPi"
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
                ExecuteAction
                return
            }

            if ($_.KeyCode -in @('Space', 'Return')) {
                $oldSelect = $Script:dataGridView.CurrentRow.DataBoundItem.Row.Select
                $Script:dataGridView.CurrentRow.DataBoundItem.Row.Select = -not $oldSelect
                $_.Handled = $true
            }
        }
    }.GetNewClosure())

    $form.add_HandleCreated({
        $global:UI_SYNCHRONIZATION_CONTEXT = [System.Threading.Tasks.TaskScheduler]::FromCurrentSynchronizationContext()
        $global:UI_THREAD_ID = [System.Threading.Thread]::CurrentThread.ManagedThreadId

        $global:FuncRPCSpellCheck = {
            param([string] $text, [int] $distance)

            if ([String]::IsNullOrEmpty($text) -or ($text.IndexOfAny('=@\/:') -ne -1)) {
                return
            }

            if ($global:SuggestionsWorking) {
                return
            } else {
                $global:SuggestionsWorking = $true
            }

            $text = $text.ToLower()
            $request = @{ 'Request'=$text; 'Distance'=$distance; } | ConvertTo-Json -Depth 5 -Compress

        	$tw = $global:sw.WriteLineAsync($request);
        	$continuation1 = New-RunspacedDelegate ( [Action[System.Threading.Tasks.Task]] {
            	$tf = $global:sw.FlushAsync();
            	$continuation2 = New-RunspacedDelegate ( [Action[System.Threading.Tasks.Task]] {
                	$tr = $global:sr.ReadLineAsync()
                    $null = $tr.ContinueWith($global:FuncRPCSpellCheck_Callback,
                        [System.Threading.CancellationToken]::None,
                        ([System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent -bor
                            [System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously),
                        $global:UI_SYNCHRONIZATION_CONTEXT);
            	});
                [void] $tf.ContinueWith($continuation2,
                    [System.Threading.CancellationToken]::None,
                    ([System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent -bor
                        [System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously),
                    $global:UI_SYNCHRONIZATION_CONTEXT);
        	});
            [void] $tw.ContinueWith($continuation1,
                [System.Threading.CancellationToken]::None,
                ([System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent -bor
                    [System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously),
                $global:UI_SYNCHRONIZATION_CONTEXT);
        }

        $global:ProcessExitedDelegate = New-RunspacedDelegate([EventHandler] {
            param($Sender, $EventArgs)
            $self = $EventArgs.self
            WriteLog "Got exit code $($self._exitCode)" -Foreground 'Red'
        })

        $global:ProcessErrorDelegate = New-RunspacedDelegate([EventHandler] {
            param($Sender, $EventArgs)
        })

        $global:ProcessOutputDelegate = New-RunspacedDelegate([EventHandler] {
            param($Sender, $EventArgs)
        })

        $global:ProcessFlushBuffersDelegate = New-RunspacedDelegate([EventHandler] {
            param($Sender, $EventArgs)
            WriteLog $EventArgs.Text -Background $EventArgs.Color
        })

    })

    # Status strip

    $statusStrip = [System.Windows.Forms.StatusStrip]::new()
    $statusStrip.ShowItemToolTips = $true

    $global:statusLabel = [System.Windows.Forms.ToolStripStatusLabel]::new()
    $statusLabel.Alignment = [System.Windows.Forms.ToolStripItemAlignment]::Left

    $spacer = [System.Windows.Forms.ToolStripStatusLabel]::new()
    $spacer.Spring = $true

    $buttonAutoScroll = [System.Windows.Forms.ToolStripButton]::new()
    $buttonAutoScroll.CheckState = if ($global:_LogViewAutoScroll) {
        [System.Windows.Forms.CheckState]::Checked } else { [System.Windows.Forms.CheckState]::Unchecked }
    $buttonAutoScroll.CheckOnClick = $true
    $buttonAutoScroll.Text = 'Autoscroll'
    $buttonAutoScroll.add_CheckStateChanged({
            param($Sender, $EventArgs)
            $global:_LogViewAutoScroll = $Sender.CheckState -eq [System.Windows.Forms.CheckState]::Checked
        })

    $dropDown = [System.Windows.Forms.ToolStripDropDown]::new()
    $dropDownItems = ('Very verbose', 'Verbose', 'Normal', 'Quiet') `
        | ForEach-Object { [System.Windows.Forms.ToolStripButton]::new($_) }
    $null = $dropDown.Items.AddRange($dropDownItems)

    $comboLogLevel = [System.Windows.Forms.ToolStripDropDownButton]::new()
    $comboLogLevel.Text = 'Log verbosity:'
    $comboLogLevel.Alignment = [System.Windows.Forms.ToolStripItemAlignment]::Right
    $comboLogLevel.DropDown = $dropDown
    $comboLogLevel.ShowDropDownArrow = $true

    $statusProgress = [System.Windows.Forms.ToolStripProgressBar]::new()
    $statusProgress.Value = 50
    $statusProgress.Alignment = [System.Windows.Forms.ToolStripItemAlignment]::Right
    $null = $statusStrip.Items.AddRange(($statusLabel, $spacer, $buttonAutoScroll, $comboLogLevel, $statusProgress))
    $null = $form.Controls.Add($statusStrip)

    return ,$form
}

Function CreateFormEnvironmentVariables($interpreterRecord) {
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
    $Form.Icon = $global:form.Icon

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

        $this.formDoc.Icon = $global:form.Icon

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

                    $childViewer = ShowDocView "$($self.PackageName).${clickedWord}"
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
                OpenLinkInBrowser $_.LinkText
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
                        $query = RequestUserString @"
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
                $null = SendMessage $Sender.Handle $WM_SCROLL $SB_LINELEFT 0
                $_.Handled = $true
            }
            if ($_.KeyCode -eq 'J') {
                $null = SendMessage $Sender.Handle $WM_VSCROLL $SB_LINEDOWN 0
                $_.Handled = $true
            }
            if ($_.KeyCode -eq 'K') {
                $null = SendMessage $Sender.Handle $WM_VSCROLL $SB_LINEUP 0
                $_.Handled = $true
            }
            if ($_.KeyCode -eq 'L') {
                $null = SendMessage $Sender.Handle $WM_SCROLL $SB_LINERIGHT 0
                $_.Handled = $true
            }
            if (($_.KeyCode -eq 'G') -and (-not $_.Shift)) {
                $null = SendMessage $Sender.Handle $WM_VSCROLL $SB_PAGETOP 0
                $richTextBox.Select(0, 0)
                $_.Handled = $true
            }
            if (($_.KeyCode -eq 'G') -and $_.Shift) {
                $null = SendMessage $Sender.Handle $WM_VSCROLL $SB_PAGEBOTTOM 0
                $textLength = $richTextBox.TextLength
                $richTextBox.Select($textLength, 0)
                $_.Handled = $true
            }
            if ($_.KeyCode -eq 'Space') {
                $null = SendMessage $Sender.Handle $WM_VSCROLL $SB_PAGEDOWN 0
                $_.Handled = $true
            }
            if (($_.KeyCode -eq 'F') -and $_.Control) {
                $null = SendMessage $Sender.Handle $WM_VSCROLL $SB_PAGEDOWN 0
                $_.Handled = $true
            }
            if (($_.KeyCode -eq 'B') -and $_.Control) {
                $null = SendMessage $Sender.Handle $WM_VSCROLL $SB_PAGEUP 0
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
    hidden [string] $_command
    hidden [string] $_arguments
    hidden [hashtable] $_environment = $null
    hidden [int] $_pid  # _process properties can change after terminating, better keep tabs on PID
    hidden [System.Diagnostics.Process] $_process = $null
    hidden [System.Threading.Tasks.TaskCompletionSource[int]] $_taskCompletionSource  # Keeps the exit code of a process
    hidden [System.Collections.Concurrent.ConcurrentQueue[string]] $_processOutput = $null
    hidden [System.Collections.Concurrent.ConcurrentQueue[string]] $_processError = $null
    hidden [bool] $_processOutputEnded = $false
    hidden [bool] $_processErrorEnded = $false
    hidden [bool] $_hasStarted = $false
    hidden [bool] $_hasFinished = $false
    hidden [bool] $_missedExitEvent = $false
    hidden [int] $_exitCode = -1
    hidden [System.Windows.Forms.Timer] $_timer = $null
    hidden [int] $_timerIdleThreshold = 20
    hidden [delegate] $_exitedCallback
    hidden [delegate] $_outputCallback
    hidden [delegate] $_errorCallback
    hidden [bool] $_LogOutput
    hidden [bool] $_LogErrors

    ProcessWithPipedIO($Command, $Arguments) {
        $this._command = $Command
        $this._arguments = $Arguments
    }

    ProcessWithPipedIO($Command, $Arguments, $Environment) {
        $this._command = $Command
        $this._arguments = $Arguments
        $this._environment = $Environment
    }

    hidden _Initialize() {
        $startInfo = [System.Diagnostics.ProcessStartInfo]::new()

        $startInfo.FileName = $this._command
        $startInfo.Arguments = $this._arguments

        if ($this._environment) {
            foreach ($var in $this._environment.GetEnumerator()) {
                $null = $startInfo.Environment.Add($var.Key, $var.Value)
            }
        }

        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true

        $startInfo.RedirectStandardInput = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true

        $startInfo.StandardOutputEncoding = [Text.Encoding]::UTF8
        $startInfo.StandardErrorEncoding = [Text.Encoding]::UTF8

        $process = [System.Diagnostics.Process]::new()
        $process.StartInfo = $startInfo
        $process.EnableRaisingEvents = $true
        $this._process = $process

        $self = $this

        $this._exitedCallback = New-RunspacedDelegate ([EventHandler] {
            param($Sender, $EventArgs)
            # $self = $EventArgs.self
            if ($self._hasFinished) {
                return
            }
            $self._hasFinished = $true
            try {
                $self._exitCode = $self._process.ExitCode
            } catch { }
            if (($self._timer -eq $null) -and $self._processOutputEnded -and $self._processErrorEnded) {
                $self._ConfirmExit($null)
            }
            # $null = Add-Member -InputObject $EventArgs -Type NoteProperty -Name self -Value $self -Force
            # $null = $form.BeginInvoke($ProcessExitedDelegate, ($Sender, $EventArgs))
        }.GetNewClosure())
        $self._process.add_Exited($this._exitedCallback)

        $this._taskCompletionSource = [System.Threading.Tasks.TaskCompletionSource[int]]::new()
    }

    hidden [System.Threading.Tasks.Task[int]] _Start() {
        $exceptionMessage = $null
        try {
            $this._hasStarted = $this._process.Start()
        } catch {
            $exceptionMessage = $_.Exception.Message
        }

        if ($this._hasStarted) {
            try {
                $this._process.StandardInput.Close()
            } catch { }
            $this._pid = $this._process.Id
        } else {
            $this._ConfirmExit([Exception]::new("Failed to start process $($this._command): $exceptionMessage"))
        }

        return $this._taskCompletionSource.Task
    }

    [System.Threading.Tasks.Task[int]] StartWithLogging([bool] $LogOutput, [bool] $LogErrors) {
        $this._Initialize()

        $self = $this
        $self._LogErrors = $LogErrors
        $self._LogOutput = $LogOutput

        WriteLog "StartWithLogging <1>"

        if ($LogOutput) {  # ReadOutputToEndAsync() is supposed to be called otherwise!
            $this._processOutput = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

            $this._outputCallback = New-RunspacedDelegate ([System.Diagnostics.DataReceivedEventHandler] {
                param($Sender, $EventArgs)
                # $self = $EventArgs.self
                $line = $EventArgs.Data
                if ($line -eq $null) {  # IMPORTANT
                    $self._processOutputEnded = $true
                } else {
                    $null = $self._processOutput.Enqueue($line)
                }
                # $null = Add-Member -InputObject $EventArgs -Type NoteProperty -Name self -Value $self -Force
                # $null = $form.BeginInvoke($ProcessOutputDelegate, ($Sender, $EventArgs))
            }.GetNewClosure())
            $this._process.add_OutputDataReceived($this._outputCallback)
        }

        if ($LogErrors) {
            $this._processError = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        }

        $this._errorCallback = New-RunspacedDelegate ([System.Diagnostics.DataReceivedEventHandler] {
            param($Sender, $EventArgs)
            # $self = $EventArgs.self
            $line = $EventArgs.Data
            if ($line -eq $null) {  # IMPORTANT
                $self._processErrorEnded = $true
            } elseif ($self._LogErrors) {
                $null = $self._processError.Enqueue($line)
            }
            # $null = Add-Member -InputObject $EventArgs -Type NoteProperty -Name self -Value $self -Force
            # $null = $form.BeginInvoke($ProcessErrorDelegate, ($Sender, $EventArgs))
        }.GetNewClosure())
        $this._process.add_ErrorDataReceived($this._errorCallback)

        if ($LogOutput -or $LogErrors) {
            $this._timer = [System.Windows.Forms.Timer]::new()

            $delegate = New-RunspacedDelegate ([EventHandler] {
                param([System.Windows.Forms.Timer] $Sender)
                # $Sender.Enabled = $false
                $self = $Sender.Tag
                $count = $self.FlushBuffersToLog()

                if ($self._hasFinished -and $self._processOutputEnded -and -$self._processErrorEnded -and ($count -eq 0)) {
                    WriteLog "Timer exiting normally <5>"
                    $self._ConfirmExit($null)
                } else {
                    if ($count) {
                        $Sender.Interval = 75
                        $self._timerIdleThreshold = 1
                    } elseif ((--$self._timerIdleThreshold) -lt 0) {
                        $throttlingInterval = 1000
                        if ($Sender.Interval -eq $throttlingInterval) { # already throttling, is it alive?
                            WriteLog "Timer throttle enter <100>"
                            # Process.Refresh() wipes its state entirely then possibly gets filled in from alive proc; WaitForExit() may lock and is not an option
                            $p = try { [System.Diagnostics.Process]::GetProcessById($self._pid) } catch { $null }
                            $actuallyDead = $p -eq $null
                            if ($p) { $p.Dispose() }
                            $p = $null

                            WriteLog "Throttling status: started=$($self._hasStarted) exited_evt=$($self._hasFinished) now_dead=$actuallyDead code=$($self._exitCode) out_end=$($self._processOutputEnded) err_end=$($self._processErrorEnded)" -Background LightPink

                            $self._missedExitEvent = $actuallyDead -ne $self._hasFinished

                            if ($actuallyDead) {
                                if (-not $self._hasFinished) {
                                    $self._exitCode = 0  # lean towards false positive on successful termination
                                }

                                # it is dead and pipes have been dry for seconds; we'll dispose potentially hung async readers

                                $self._processOutputEnded = $true
                                $self._processErrorEnded = $true
                            }

                            $self._hasFinished = $actuallyDead
                            WriteLog "Timer throttle EXIT <101>"
                        } else {
                            $Sender.Interval = $throttlingInterval  # our child process is silent, we'll throttle
                        }
                    }
                    # $Sender.Enabled = $true
                }
            })

            $this._timer.Tag = $self
            $this._timer.Interval = 75
            $this._timer.add_Tick($delegate)
        }

        WriteLog "StartWithLogging <2>"
        $null = $this._Start()
        WriteLog "StartWithLogging <3>"

        if ($this._hasStarted) {
            if ($LogOutput) {
                $this._process.BeginOutputReadLine()
            }
            if ($LogErrors) {
                $this._process.BeginErrorReadLine()
            }
            if ($LogOutput -or $LogErrors) {
                $this._timer.Start()
                WriteLog "Timer START <0>"
            }
        }

        WriteLog "StartWithLogging <4>"
        return $this._taskCompletionSource.Task
    }

    hidden _ConfirmExit($exception) {
        WriteLog "_ConfirmExit E='$exception' OUT=$($this._processOutputEnded) ERR=$($this._processErrorEnded)"

        if ($this._timer) {
            $this._timer.Stop()
            $this._timer = $null
        }

        $this._process.EnableRaisingEvents = $false
        $this._process.remove_Exited($this._exitedCallback)
        $this._exitedCallback = $null
        if ($this._errorCallback) {
             $this._process.remove_ErrorDataReceived($this._errorCallback)
             $this._errorCallback = $null
         }
        if ($this._outputCallback) {
             $this._process.remove_OutputDataReceived($this._outputCallback)
             $this._outputCallback = $null
         }

        $delegate = New-RunspacedDelegate ([Action[object]] {
            param([object] $locals)
            $self = $locals.self
            $exception = $locals.exception
            if (($exception -ne $null) -or ($self._exitCode -eq -1)) {
                $null = $self._taskCompletionSource.TrySetException([Exception]::new($exception))
            } else {
                $null = $self._taskCompletionSource.TrySetResult($self._exitCode)
            }
            if (-not $self._hasFinished) {
                WriteLog "Killing" -Background Red
                try { $null = $self._process.Kill() } catch { }
            }
            if ($self._missedExitEvent) {
                WriteLog "We are in trouble, missed exit event, this shouldn't happen!" -Background Magenta
            }

            $processClenup = New-RunspacedDelegate([Action[object]] {
                param([ProcessWithPipedIO] $self)
                $self._process.WaitForExit()
                try {
                    $self._process.Dispose()
                } finally {
                    $self._process = $null
                }
            })
            $token = [System.Threading.CancellationToken]::None
            $options = ([System.Threading.Tasks.TaskCreationOptions]::DenyChildAttach)
            $taskGetBuiltinPackages = [System.Threading.Tasks.Task]::Factory.StartNew($processClenup, $self,
                $token, $options, [System.Threading.Tasks.TaskScheduler]::Default)

            WriteLog "Exiting _ConfirmExit" -Background DarkOrange
        })
        $token = [System.Threading.CancellationToken]::None
        $options = ([System.Threading.Tasks.TaskCreationOptions]::AttachedToParent)
        $null = [System.Threading.Tasks.Task]::Factory.StartNew($delegate, @{self=$this; exception=$exception}, $token,
            $options, $global:UI_SYNCHRONIZATION_CONTEXT)
    }

    Kill() {
        $this._process.CancelOutputRead()
        $this._process.CancelErrorRead()
        TryTerminateGracefully($this._process)
    }

    hidden [int] _FlushContainerToLog($container, $color) {
        $count = 0
        if ($container -and (-not $container.IsEmpty)) {
            $buffer = [System.Text.StringBuilder]::new()
            while (-not $container.IsEmpty) {
                [ref] $lineRef = [ref] [string]::Empty
                try {
                    if (-not $container.TryDequeue($lineRef)) {
                        break
                    }
                } catch { }
                [string] $line = $lineRef.Value

                if ($buffer.Length -gt 0) {
                    $null = $buffer.Append([Environment]::NewLine)
                }
                $null = $buffer.Append($line)
                ++$count
            }
            $EventArgs = global:MakeEvent @{Text=$buffer.ToString(); Color=$color}
            $null = $global:form.BeginInvoke($global:ProcessFlushBuffersDelegate, ($this, $EventArgs))
        }
        return $count
    }

    hidden [int] FlushBuffersToLog () {
        $outLines = $this._FlushContainerToLog($this._processOutput, 'LightBlue')
        $errLines = $this._FlushContainerToLog($this._processError, 'LightSalmon')
        return $outLines + $errLines
    }

    [System.Threading.Tasks.Task[string]] ReadOutputToEndAsync() {
        $continuationReadingDone = New-RunspacedDelegate ([Action[System.Threading.Tasks.Task, object]] {
            param([System.Threading.Tasks.Task] $task, [object] $locals)
            $locals.self._processOutputEnded = $true
        })
        try {
            $taskRead = $this._process.StandardOutput.ReadToEndAsync()
            $options = ([System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent -bor `
                [System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously)
            $null = $taskRead.ContinueWith($continuationReadingDone, @{ self=$this; },
                [System.Threading.CancellationToken]::None, $options, $global:UI_SYNCHRONIZATION_CONTEXT)
            return $taskRead
        } catch {
            $this._processOutputEnded = $true
        }
        return [System.Threading.Tasks.Task[string]]::FromResult($null)
    }

    WaitForExit() {
        $this._process.WaitForExit()
    }
}


class WidgetStateTransition {

    hidden [System.Collections.Generic.Stack[hashtable]] $_states
    hidden [System.Collections.Generic.List[System.Windows.Forms.Control]] $_controls
    hidden [System.Collections.Generic.HashSet[Action]] $_actions
    hidden [System.Collections.Generic.Stack[object]] $_modes

    WidgetStateTransition () {
        $this._states = [System.Collections.Generic.Stack[hashtable]]::new()
        $this._controls = [System.Collections.Generic.List[System.Windows.Forms.Control]]::new()
        $this._modes = [System.Collections.Generic.Stack[object]]::new()
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
                # Write-Information $property.Key '=' $property.Value ' of ' $property.Value.GetType()
                if (($property.Value -isnot [ScriptBlock]) -and ($property.Value -isnot [delegate])) {
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

                { ($_ -is [System.Windows.Forms.Control]) -or ($_ -is [System.Windows.Forms.Form]) } {
                    $control, $properties = $widgetState.Key, $widgetState.Value
                    foreach ($property in $properties.GetEnumerator()) {
                        # Write-Information 'REV ' $property.Key '=' $property.Value ' of ' $property.Value.GetType()
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
        $delegate = New-RunspacedDelegate ([Action[System.Threading.Tasks.Task, object]]{
            param([System.Threading.Tasks.Task] $task, [object] $widgetStateTransition)
            $null = ($widgetStateTransition -as [WidgetStateTransition]).ReverseAll()
        }.GetNewClosure())
        return $delegate
    }

    [WidgetStateTransition] Debounce([Action] $action) {
        if (-not $this._actions.Contains($action)) {
            $this._actions.Add($action)
        }
        return $this
    }

    static hidden [string] FormatInternalEventName([type] $type, [string] $event) {
        $name = switch ($type) {
            { $_ -eq [System.Windows.Forms.Form] } { "EVENT_$event".ToUpper() ; break }
            { $_ -eq [System.Windows.Forms.Control] } { "Event$event" ; break }
        }
        return $name
    }

    static [delegate[]] SpliceEventHandlers([object] $control, [string] $event, $handlers) {

        $old = [System.Collections.Generic.List[delegate]]::new()
        $type = $control.GetType()
        $baseType = switch ($control) {
            { $_ -is [System.Windows.Forms.Form] } { [System.Windows.Forms.Form] ; break }
            { $_ -is [System.Windows.Forms.Control] } { [System.Windows.Forms.Control] ; break }
        }
        $internalEventName = [WidgetStateTransition]::FormatInternalEventName($baseType, $event)

        # WriteLog $type -Background Pink
        # WriteLog $baseType -Background Pink

        $propertyInfo = $type.GetProperty('Events',
            [System.Reflection.BindingFlags]::Instance -bor
            [System.Reflection.BindingFlags]::NonPublic -bor
            [System.Reflection.BindingFlags]::Static)

        $eventHandlerList = $propertyInfo.GetValue($control)

        $fieldInfo = $baseType.GetField($internalEventName,
            [System.Reflection.BindingFlags]::Static -bor
            [System.Reflection.BindingFlags]::NonPublic)

        $eventKey = $fieldInfo.GetValue($control)

        $eventHandler = $eventHandlerList[$eventKey]

        if ($eventHandler) {
            $invocationList = @($eventHandler.GetInvocationList())

            foreach ($handler in $invocationList) {
                $null = $old.Add($handler)
                $control."remove_$event"($handler)
                # WriteLog "EVENT REMOVE $event $control $handler" -Background Red
            }
        }

        foreach ($handler in $handlers) {
            $control."add_$event"($handler)
            # WriteLog "EVENT ADD $event $control $handler" -Background Green
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

    [WidgetStateTransition] PostIrreversibleTransform([hashtable] $properties) {
        return $this
    }

    [WidgetStateTransition] PostIrreversibleTransformWithMethodCall($methodName, [object[]] $methodArguments) {
        return $this
    }

    [WidgetStateTransition] CommitIrreversibleTransformations() {
        return $this
    }

    [WidgetStateTransition] AllocateStatusLine([ref] $statusLineToken) {
        return $this
    }

    [WidgetStateTransition] FreeStatusLine([object] $statusLineToken) {
        return $this
    }

    [WidgetStateTransition] SetStatusLineText([object] $statusLineToken, [string] $text) {
        return $this
    }

    [WidgetStateTransition] AppendStatusLineText([object] $statusLineToken, [string] $text) {
        return $this
    }

    [WidgetStateTransition] EnterMode([object] $mode, [ref] $successRef) {
        $success = ($this._modes.Count -eq 0) -or ($this._modes.Peek() -ne $mode)
        if ($success) {
            $this._modes.Push($mode)
        }
        if ($successRef -ne $null) {
            $successRef.Value = $success
        }
        return $this
    }

    [WidgetStateTransition] IsModeActive([object] $mode, [ref] $result) {
        $result.Value = ($this._modes.Count -gt 0) -and ($this._modes.Peek() -eq $mode)
        return $this
    }

    [WidgetStateTransition] ExitMode() {
        $null = ($this._modes.Count -eq 0) -or $this._modes.Pop()
        return $this
    }

}


Function global:ShowDocView($packageName, $SetContent = $null, $Highlight = $null, [switch] $NoDefaultHighlighting) {
    if (-not $SetContent) {
        $content = (GetPyDoc $packageName) -join "`n"
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
    InitPackageSearchColumns $selected

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
    $pip_exe = GetCurrentInterpreter 'PipExe' -Executable
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
    $conda_exe = GetCurrentInterpreter 'CondaExe' -Executable
    if (-not $conda_exe) {
        WriteLog 'conda is not found!'
        return 0
    }
    $arch = GetCurrentInterpreter 'Arch'

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
    $json = DownloadString ($github_search_url -f [System.Web.HttpUtility]::UrlEncode($request))
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
            (GetCurrentInterpreter 'Version'),
            (GetCurrentInterpreter 'Bits'))

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
        $null = DownloadString $url -ContinueWith {
            param($json)
            if (-not [string]::IsNullOrEmpty($json)) {
                $tags = $json | ConvertFrom-Json | ForEach-Object { $_.Name }
                $null = & $ContinueWith $tags
                return
            }
        }.GetNewClosure()
        return
    }

    $json = DownloadString $url
    if (-not [string]::IsNullOrEmpty($json)) {
        $tags = $json | ConvertFrom-Json | Select-Object -ExpandProperty Name # ForEach-Object { $_.Name }
        return $tags
    }

    return $null
}

Function Get-SearchResults($request) {
    $previousSelected = Store-CheckedPipSearchResults
    ClearRows
    InitPackageSearchColumns $global:dataModel

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

Function global:AddPackagesToTable {
    param($packages, $defaultType = [String]::Empty)

    $global:dataModel.BeginLoadData()
    $headersSizeMode = $dataGridView.RowHeadersWidthSizeMode
    $columnsSizeMode = $dataGridView.AutoSizeColumnsMode
    $dataGridView.RowHeadersWidthSizeMode = [System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode]::DisableResizing
    $dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None

    foreach ($package in $packages) {
        $row = $global:dataModel.NewRow()
        $row.Select = $false
        $row.Package = $package.Package

        $availableKeys = $package.PSObject.Properties.Name

        # write-host $availableKeys
        # Write-Information $package.Package $availableKeys

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

    $dataGridView.RowHeadersWidthSizeMode = $headersSizeMode
    $dataGridView.AutoSizeColumnsMode = $columnsSizeMode
    $global:dataModel.EndLoadData()
}

Function global:GetPythonPackages($outdatedOnly = $true) {
    ClearRows
    InitPackageUpdateColumns $global:dataModel
    $global:outdatedOnly = $outdatedOnly

    $continuationAllDone = New-RunspacedDelegate ([Action[System.Threading.Tasks.Task]] {
        param([System.Threading.Tasks.Task] $task)

        [object[]] $tuplesPackagesType = $task.Result
        $stats = @{}

        foreach ($tuple in $tuplesPackagesType) {
            $packages, $type = $tuple.Item1, $tuple.Item2
            AddPackagesToTable $packages $type
            $null = $stats.Add($type, $packages.Count)
        }

        HighlightPackages

        WriteLog 'Double click or [Ctrl+Enter] a table row to open package''s home page in browser'
        $total = $global:dataModel.Rows.Count
        $stats = ($stats.GetEnumerator() | ForEach-Object { "$($_.Value) $($_.Key)"  }) -join ', '
        WriteLog "Total $total packages: $stats" -Background LightGreen
    })

    $allTasks = [System.Collections.Generic.List[System.Threading.Tasks.Task[object]]]::new()

    $python_exe = GetCurrentInterpreter 'PythonExe' -Executable
    if ($python_exe) {
        # Func [Task[string], Tuple`2[System.Management.Automation.PSObject, System.String]]
        $continuationParsePipOutput = New-RunspacedDelegate ([Func[System.Threading.Tasks.Task, object]] {
            param([System.Threading.Tasks.Task] $task)

            $csv = $task.Result -replace ' +',' '  # ConvertFrom-Csv understands only one space between columns

            $pipPackages = $csv | ConvertFrom-Csv -Header $csv_header -Delimiter ' '

            if ($pipPackages -eq $null) {
                throw [Exception]::new("Failed to parse CSV from pip: $csv")
            }

            $pipPackages = $pipPackages | Select-Object -Skip 2  # Ignore a header line and a separator line

            return [Tuple]::Create($pipPackages, 'pip')
        })

        $arguments = New-Object System.Collections.ArrayList
        $null = $arguments.AddRange(('-m', 'pip'))
        $null = $arguments.Add('list')
        $null = $arguments.Add('--format=columns')

        if ($outdatedOnly) {
            $null = $arguments.Add('--outdated')
        }

        $process = [ProcessWithPipedIO]::new($python_exe, $arguments)
        $taskProcessDone  = $process.StartWithLogging($false, $true)
        $taskReadOutput = $process.ReadOutputToEndAsync()
        $taskPipList = $taskReadOutput.ContinueWith($continuationParsePipOutput,
            [System.Threading.CancellationToken]::None,
            ([System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent -bor
                [System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously),
            $global:UI_SYNCHRONIZATION_CONTEXT)
        $null = $allTasks.Add($taskPipList)
    }

    $conda_exe = GetCurrentInterpreter 'CondaExe' -Executable
    if ($conda_exe) {
        $continuationAddCondaPackages = New-RunspacedDelegate([Func[System.Threading.Tasks.Task, object]] {
            param([System.Threading.Tasks.Task] $task)
            $condaPackages = $task.Result
            if ($condaPackages -eq $null) {
                throw [Exception]::new("Empty response from conda.")
            }
            return [Tuple]::Create($condaPackages, 'conda')
        })
        $task = GetCondaPackagesAsync $outdatedOnly
        $taskCondaList = $task.ContinueWith($continuationAddCondaPackages,
            [System.Threading.CancellationToken]::None,
            ([System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent -bor
                [System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously),
            $global:UI_SYNCHRONIZATION_CONTEXT)
        $null = $allTasks.Add($taskCondaList)
    }

    $continuationAddBuiltinPackages = New-RunspacedDelegate([Func[System.Threading.Tasks.Task, object]] {
        param([System.Threading.Tasks.Task] $task)
        $builtinPackages = $task.Result
        if ($builtinPackages -eq $null) {
            throw [Exception]::new("Empty response from builtin packages.")
        }
        return [Tuple]::Create($builtinPackages, 'builtin')
    })
    $taskGetBuiltinPackages = GetPythonBuiltinPackagesAsync
    $taskAddBuiltinPackages = $taskGetBuiltinPackages.ContinueWith($continuationAddBuiltinPackages,
        [System.Threading.CancellationToken]::None,
        ([System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent -bor
            [System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously),
        $global:UI_SYNCHRONIZATION_CONTEXT)
    $null = $allTasks.Add($taskAddBuiltinPackages)

    # other packages are packages that were found but do not belong to any other list

    $continuationAddOtherPackages = New-RunspacedDelegate([Func[System.Threading.Tasks.Task, object]] {
        param([System.Threading.Tasks.Task] $task)
        $otherPackages = $task.Result
        if ($otherPackages -eq $null) {
            throw [Exception]::new("Empty response from other packages.")
        }
        return [Tuple]::Create($otherPackages, 'other')
    })
    $taskGetOtherPackages = GetPythonOtherPackagesAsync
    $taskAddOtherPackages = $taskGetOtherPackages.ContinueWith($continuationAddOtherPackages,
        [System.Threading.CancellationToken]::None,
        ([System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent -bor
            [System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously),
        $global:UI_SYNCHRONIZATION_CONTEXT)
    $null = $allTasks.Add($taskAddOtherPackages)

    # WriteLog $allTasks -Background DarkCyan -Foreground Yellow

    try {
        $gathered = [System.Threading.Tasks.Task]::WhenAll($allTasks)
        $taskAllDone = $gathered.ContinueWith($continuationAllDone,
            [System.Threading.CancellationToken]::None,
            ([System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent -bor
                [System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously),
            $global:UI_SYNCHRONIZATION_CONTEXT)
    } catch [System.AggregateException] {
        WriteLog "One or more tasks have failed" -Background DarkRed
    }

    return $taskAllDone

    if (-not $outdatedOnly) {
        $builtinPackages = GetPythonBuiltinPackages
        Add-PackagesToTable $builtinPackages 'builtin'


        $builtinCount = $builtinPackages.Count
        $otherCount = $otherPackages.Count
    }
}

Function global:SetVisiblePackageCheckboxes {
    [CmdletBinding()]
    param(
        [bool] $Value,
        [AllowNull()] [string[]] $Filter = $null,
        [switch] $Inverse)

    $global:dataModel.BeginLoadData()

    $headersSizeMode = $dataGridView.RowHeadersWidthSizeMode
    $columnsSizeMode = $dataGridView.AutoSizeColumnsMode
    $dataGridView.RowHeadersWidthSizeMode = [System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode]::DisableResizing
    $dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None

    for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
        $type = $dataGridView.Rows[$i].DataBoundItem.Row.Type
        if ($type -in @('builtin', 'other') ) {
            continue
        }
        if ($Filter -and ($type -notin $Filter)) {
            continue
        }
        if ($Inverse) {
            $Value = -not $dataGridView.Rows[$i].DataBoundItem.Row.Select
        }
        $dataGridView.Rows[$i].DataBoundItem.Row.Select = $Value
    }

    $dataGridView.RowHeadersWidthSizeMode = $headersSizeMode
    $dataGridView.AutoSizeColumnsMode = $columnsSizeMode

    $global:dataModel.EndLoadData()
}

Function global:SetAllPackageCheckboxes($value) {
    $global:dataModel.BeginLoadData()

    $headersSizeMode = $dataGridView.RowHeadersWidthSizeMode
    $columnsSizeMode = $dataGridView.AutoSizeColumnsMode
    $dataGridView.RowHeadersWidthSizeMode = [System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode]::DisableResizing
    $dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None

    for ($i = 0; $i -lt $global:dataModel.Rows.Count; $i++) {
       $global:dataModel.Rows[$i].Select = $value
    }

    $dataGridView.RowHeadersWidthSizeMode = $headersSizeMode
    $dataGridView.AutoSizeColumnsMode = $columnsSizeMode

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

Function global:ClearRows() {
    $global:outdatedOnly = $true
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

Function global:SetSelectedRow($selectedRow) {
    $global:dataGridView.ClearSelection()
    foreach ($vRow in $global:dataGridView.Rows) {
        if ($vRow.DataBoundItem.Row -eq $selectedRow) {
            $vRow.Selected = $true
            $dataGridView.FirstDisplayedScrollingRowIndex = $vRow.Index
            $dataGridView.CurrentCell = $dataGridView[0, $vRow.Index]
            break
        }
    }
}

Function global:CheckDependencies {
    $python_exe = GetCurrentInterpreter 'PythonExe' -Executable
    if (-not $python_exe) {
        WriteLog 'Python is not found!'
        throw [Exception]::new("Python is not found!")
    }

    WriteLog 'Checking dependencies...'

    $process = [ProcessWithPipedIO]::new($python_exe, @('-m', 'pip', 'check'))
    $taskProcessDone = $process.StartWithLogging($false, $true)
    $taskReadOutput = $process.ReadOutputToEndAsync()

    $continuationReport = New-RunspacedDelegate ([Action[System.Threading.Tasks.Task[string]]] {
        param([System.Threading.Tasks.Task[string]] $task)
        $result = $task.Result

        if ($result -match 'No broken requirements found') {
            WriteLog "Dependencies are OK" -Background ([Drawing.Color]::LightGreen)
        } else {
            WriteLog "Dependencies are NOT OK" -Background ([Drawing.Color]::LightSalmon)
            WriteLog $result
        }
    })
    $taskReportLogged = $taskReadOutput.ContinueWith($continuationReport,
            [System.Threading.CancellationToken]::None,
            ([System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent -bor
                [System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously),
            $global:UI_SYNCHRONIZATION_CONTEXT)

    $allTasks = [System.Collections.Generic.List[System.Threading.Tasks.Task]]::new()
    $null = $allTasks.Add($taskProcessDone)
    $null = $allTasks.Add($taskReportLogged)

    return [System.Threading.Tasks.Task]::WhenAll($allTasks)
}

Function global:SelectPipAction($actionName) {
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

    $RequestUserAppExit = [System.Windows.Forms.FormClosingEventHandler] {
        param([object] $Sender, [System.Windows.Forms.FormClosingEventArgs] $EventArgs)
        $response = [System.Windows.Forms.MessageBox]::Show('Sure?', 'Cancel', [System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($response -eq [System.Windows.Forms.DialogResult]::Yes) {
            $null = $script:widgetStateTransition.ReverseAll()
        } else {
            $EventArgs.Cancel = $true
        }
    }.GetNewClosure()

    $RequestUserConfirmCancel = {
        param([object] $Sender, [EventArgs] $EventArgs)
        $response = [System.Windows.Forms.MessageBox]::Show('Sure?', 'Cancel', [System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($response -eq [System.Windows.Forms.DialogResult]::Yes) {
            $null = $script:widgetStateTransition.ReverseAll()
        }
    }.GetNewClosure()

    $null = $widgetStateTransition.Add($form).Transform(@{FormClosing=$RequestUserAppExit})
    $disableThemWhileWorking = @(
        $WIDGET_GROUP_COMMAND_BUTTONS ;
        $WIDGET_GROUP_ENV_BUTTONS ;
        $WIDGET_GROUP_INSTALL_BUTTONS ;
        $interpretersComboBox ;
        $actionListComboBox
    )
    $null = $widgetStateTransition.AddRange($disableThemWhileWorking).Transform(@{Enabled=$false})
    $null = $widgetStateTransition.Add($button).Transform(@{Text='Cancel';Enabled=$true;Click=$RequestUserConfirmCancel})
    $null = $widgetStateTransition.GlobalVariables(@{
            APP_MODE=([AppMode]::Working);
        })
    return $widgetStateTransition
}

Function global:ExecuteAction {
    [CmdletBinding()]
    param($Sender, $EventArgs,
        <#
        .DESCRIPTION
        Optional parameters for this function are supposed to be passed when using alternative button modes
        #>
        [Parameter(Mandatory=$false)] [switch] $Serial = $false,
        [Parameter(Mandatory=$false)] [switch] $ShowCommand = $false,
        [AllowEmptyCollection()][AllowNull()][Parameter(Mandatory=$false)] [string[]] $CustomArguments = $null)

    $fireAction = New-RunspacedDelegate ([Action[object]] {
        param([object] $fireActionLocals)
        $action = $fireActionLocals.action
        $actionId = $action.Id
        $execActionTaskCompletionSource = $fireActionLocals.execActionTaskCompletionSource
        $Serial = $fireActionLocals.Serial

        $reportBatchResults = New-RunspacedDelegate([Action[System.Threading.Tasks.Task, object]] {
            param([System.Threading.Tasks.Task] $task, [object] $FunctionContext)
            $tasksOkay, $tasksFailed = $FunctionContext.tasksOkay, $FunctionContext.tasksFailed

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

            $null = $FunctionContext.execActionTaskCompletionSource.TrySetResult($null)
        })

        $continuationReportIteration = New-RunspacedDelegate([Action[System.Threading.Tasks.Task[int], object]] {
            param([System.Threading.Tasks.Task[int]] $task, [object] $locals)

            $exitCode = -1
            $success = $false

            if ($task.IsCompleted -and (-not ($task.IsFaulted -or $task.IsCanceled))) {
                $exitCode = $task.Result
                $success = $exitCode -eq 0
            }

            if ($exitCode -ge 0) {
                $color = if ($success) { 'DarkGreen' } else { 'DarkRed' }
                WriteLog "Exited with code $exitCode" -Background $color -Foreground White
            } else {
                $message = $task.Exception.InnerException
                WriteLog "Failed: $message" -Background DarkRed -Foreground White
            }

            if ($success) {
                $locals.functionContext.tasksOkay += 1
            } else {
                $locals.functionContext.tasksFailed += 1
            }

            # PostIrreversibleTransformWithMethodCall()
            # $global:dataModel.Columns['Status'].ReadOnly = $false
            # $global:dataModel.Columns['Status'].ReadOnly = $true
            # $logTo = (GetLogLength) - $locals.logFrom
            # $locals.dataRow | Add-Member -Force -MemberType NoteProperty -Name LogFrom -Value $locals.logFrom
            # $locals.dataRow | Add-Member -Force -MemberType NoteProperty -Name LogTo -Value $logTo

            #$widgetStateTransition.CommitIrreversibleTransformations()
        })

        $function = New-RunspacedDelegate ([Func[object, object, System.Threading.Tasks.Task]] {
            param([object] $element, [object] $FunctionContext)
            $actionId = $FunctionContext.actionId  # what should we do with the specific package type(s)
            $interpreter = $FunctionContext.interpreter  # current Python
            $logFrom = GetLogLength
            $executableArguments = $null
            $executableCommand = $null
            $variables = $null
            $functions = @{ py={ param($property) $interpreter."$property" }; }  # funcs to be called from ActionCommands

            trap {
                $null = WriteLog "Failed to ${actionId}: $_" -Background LightSalmon
                return [System.Threading.Tasks.Task]::FromException($_.Exception)
            }

            if ($element -is [System.Collections.DictionaryEntry]) {  # Bulk operation on [String type, DataRow[] packages]
                $type, $dataRows = $element.Key, $element.Value

                $action = GetActionCommand $type $actionId
                if (-not $action) {
                    throw [Exception]::new("action $actionId is unavailable for type $type")
                }

                $packages = [System.Collections.Generic.List[string]]::new()
                foreach ($dataRow in $dataRows) {
                    $null = $packages.Add("{0}=={1}" -f ($dataRow.Package,$dataRow.Installed))
                }
                $packages = $packages -join ' '
                WriteLog "Running $($actionId) on $packages" -Background LightPink

                $variables = @{ package=$packages; count=$dataRows.Count; }  # vars for ActionCommands
            } else {
                $dataRow = $element
                $package, $installed, $type = $dataRow.Package, $dataRow.Installed, $dataRow.Type

                $action = GetActionCommand $type $actionId
                if (-not $action) {
                    throw [Exception]::new("action $actionId is unavailable for type $type")
                }

                WriteLog "Running $($actionId) on $package" -Background LightPink

                $variables = @{ package=$package; version=$installed; }
            }

            $isInternalCommand = $action.Command -is [ScriptBlock]  # if it's not an external process then handle it right away (causes hang on long op)
            if ($isInternalCommand) {
                if ($FunctionContext.ShowCommand) {
                    WriteLog "Can't show a command line for $actionId because it's implemented in pips."
                } else {
                    try {
                        $null = InvokeWithContext $action.Command $functions $variables @()
                        $FunctionContext.tasksOkay += 1
                    } catch {
                        $FunctionContext.tasksFailed += 1
                        throw $_.Exception
                    }
                }
                return [System.Threading.Tasks.Task]::FromResult(@{})
            } else {
                $executableCommand = $interpreter."$($action.Command)"
                if ($FunctionContext.CustomArguments) {
                    $executableArguments = "$($FunctionContext.CustomArguments) $($variables.package)"
                } else {
                    $executableArguments = InvokeWithContext $action.Args $functions $variables @()  # substitute actual params
                }
                WriteLog "$executableCommand $executableArguments" -Background Green -Foreground White
            }

            if ($FunctionContext.ShowCommand) {
                return [System.Threading.Tasks.Task]::FromResult(@{})
            }

            $process = [ProcessWithPipedIO]::new($executableCommand, $executableArguments)
            $taskProcessDone = $process.StartWithLogging($true, $true)

            $reportLocals = @{
                dataRow=$dataRow;
                process=$process;
                logFrom=$logFrom;
                functionContext=$FunctionContext;
            }

            $taskReport = $taskProcessDone.ContinueWith($FunctionContext.continuationReportIteration, $reportLocals,
                [System.Threading.CancellationToken]::None,
                ([System.Threading.Tasks.TaskContinuationOptions]::AttachedToParent -bor
                    [System.Threading.Tasks.TaskContinuationOptions]::ExecuteSynchronously),
                $global:UI_SYNCHRONIZATION_CONTEXT)
            return $taskReport
        })


        $functionContext = @{
            actionId=$actionId;
            interpreter=(GetCurrentInterpreter);
            tasksOkay=0;
            tasksFailed=0;
            execActionTaskCompletionSource=$execActionTaskCompletionSource;
            ShowCommand=$fireActionLocals.ShowCommand;
            CustomArguments=$fireActionLocals.CustomArguments;
            continuationReportIteration=$continuationReportIteration;
        }

        $queue = [System.Collections.Generic.Queue[object]]::new()

        for ($i = 0; $i -lt $global:dataModel.Rows.Count; $i++) {
            if ($global:dataModel.Rows[$i].Select -eq $true) {
                $null = $queue.Enqueue($global:dataModel.Rows[$i])
            }
        }

        if ($action.Bulk -and (-not $Serial) -and (-not $action.AllTypes) -and ($queue.Count -gt 0)) {
            $queueTypePackages = $queue | Group-Object -Property @{Expression={ $_.Type }} -AsString -AsHashTable
            $queue.Clear()
            foreach ($pair in $queueTypePackages.GetEnumerator()) {
                $queue.Enqueue($pair)
            }
        } elseif ($action.Singleton) {
            $firstElement = $queue.Dequeue()
            $queue.Clear()
            $queue.Enqueue($firstElement)
        } elseif ($action.Bulk -and $action.AllTypes) {
            $everything = $queue.ToArray()
            $queue.Clear()
            $queue.Enqueue([System.Collections.DictionaryEntry]::new('common', $everything))
        }

        $null = ApplyAsync $functionContext $queue $function $reportBatchResults
    })

    $execActionTaskCompletionSource = [System.Threading.Tasks.TaskCompletionSource[object]]::new()

    $fireActionLocals = @{
        action=$global:actionsModel[$actionListComboBox.SelectedIndex];
        execActionTaskCompletionSource=$execActionTaskCompletionSource;
        Serial=$Serial;
        ShowCommand=$ShowCommand;
        CustomArguments=$CustomArguments;
    }
    $token = [System.Threading.CancellationToken]::None
    $options = ([System.Threading.Tasks.TaskCreationOptions]::AttachedToParent -bor
        [System.Threading.Tasks.TaskCreationOptions]::PreferFairness)

    $null = [System.Threading.Tasks.Task]::Factory.StartNew($fireAction, $fireActionLocals,
        $token, $options, $global:UI_SYNCHRONIZATION_CONTEXT)

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

    # test if each key in the collection is pressed
    $Result = foreach ($Key in $Keys)
    {
        [bool]($WinAPI::GetAsyncKeyState($Key) -eq -32767)
    }

    # if all are pressed, return true, if any are not pressed, return false
    $Result -notcontains $false
}

Function global:GetReverseDependencies {
    $di = GetPackageDistributionInfo
    $packages = $di | Get-Member -Type NoteProperty | Select-Object -ExpandProperty Name
    foreach ($p in $packages) {
        $deps = $di."$p".deps | Get-Member -Type NoteProperty | Select-Object -ExpandProperty Name
        if ($package -in $deps) {
            WriteLog "$p"
        }
    }
}

Function global:GetPackageDistributionInfo {
    $python_code = @'
import pkg_resources
import json
pkgs = pkg_resources.working_set
info = {p.key: {'deps': {r.name: [str(s) for s in r.specifier] for r in p.requires()}, 'extras': p.extras} for p in pkgs}
print(json.dumps(info))
'@ -join ';'

    $output = & (GetCurrentInterpreter 'PythonExe') -c "`"$python_code`""
    $pkgs = $output | ConvertFrom-Json
    return $pkgs
}

Function global:GetAsciiTree($output, $name,
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

    $null = $output.AppendLine("${prefix}${name}${suffix}")  # Add a line to the Return Stack

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
            GetAsciiTree $output $child `
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

Function global:GetDependencyAsciiGraph($name) {
    $null = PreparePackageAutoCompletion  # for checking presence of pkg in the index
    $distributionInfo = GetPackageDistributionInfo
    $output = [System.Text.StringBuilder]::new()
    $null = GetAsciiTree $output $name.ToLower() $distributionInfo
    return $output.ToString()
}

Function global:SetConsoleVisibility {
    [CmdletBinding()]
    param([bool] $Visible)

    $cp = [System.CodeDom.Compiler.CompilerParameters]::new()
    $cp.GenerateInMemory = $true
    $cp.WarningLevel = 0

    $global:SW_HIDE = [int] 0
    $global:SW_MINIMIZE = [int] 6
    $global:SW_SHOWNOACTIVATE = [int] 4
    $global:SW_SHOWNA = [int] 8
    $global:SW_RESTORE = [int] 9

    $consolePtr = $WinAPI::GetConsoleWindow()
    $value = if ($Visible) { $SW_SHOWNOACTIVATE } else { $SW_HIDE }
    $WinAPI::ShowWindow($consolePtr, $value)
}

Function global:SaveSettings {
    $settingsPath = "$($env:LOCALAPPDATA)\pips"
    $null = New-Item -Force -ItemType Directory -Path $settingsPath
    $userInterpreterRecords = $interpreters | Where-Object { $_.User }
    $settings."envs" = @($userInterpreterRecords)
    try {
        $settings | ConvertTo-Json -Depth 25 | Out-File "$settingsPath\settings.json"
    } catch {
    }
}

Function global:LoadSettings {
    $settingsFile = "$($env:LOCALAPPDATA)\pips\settings.json"
    $global:settings = @{ envs=@(); condaChannels=@(); }
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

Function LoadPlugins() {
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
                [string[]] $output = & (GetCurrentInterpreter 'PythonExe') -m pip download --no-deps --no-index --progress-bar off --dest $destination $url
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
            ${function:PreparePackageAutoCompletion},
            ${function:Recode})
        [void] $global:plugins.Add($instance)
        [void] $global:packageTypes.AddRange($instance.GetSupportedPackageTypes())
        WriteLog $instance.GetDescription()
    }
}

Function global:UninstallDebugHelpers {
   [System.AppDomain]::CurrentDomain.remove_FirstChanceException($firstChanceExceptionHandler)
   [System.AppDomain]::CurrentDomain.remove_UnhandledException($unhandledExceptionHandler)
}

Function global:InstallDebugHelpers {
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
        [System.Management.Automation.PSNotSupportedException],
        [System.TimeoutException],
        [System.InvalidOperationException]
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
            ForegroundColor=((5 + $color) % 15);
        }
        Write-Host ('=' * 70) @color
        Write-Host 'Managed TID=' ([System.Threading.Thread]::CurrentThread.ManagedThreadId) ', is POOL=' ([System.Threading.Thread]::CurrentThread.IsThreadPoolThread) ', is BACKGRND=' ([System.Threading.Thread]::CurrentThread.IsBackground) -BackgroundColor White -ForegroundColor Black ', UI TID=' $global:UI_THREAD_ID
        Write-Host ('-' * 70) @color
        Write-Host $Exception.GetType() @color
        Write-Host ('-' * 70) @color
        Write-Host $Exception.Message @color
        Write-Host ('-' * 70) @color

        if ((($Exception.GetType() -in $exceptionsWithScriptBacktrace) -or
            ($Exception.GetType().BaseType -in $exceptionsWithScriptBacktrace)) -and
            -not [string]::IsNullOrWhiteSpace($Exception.ErrorRecord.ScriptStackTrace) )
            {
            $ScriptStackTrace = $Exception.ErrorRecord.ScriptStackTrace
            if (-not [string]::IsNullOrWhiteSpace($ScriptStackTrace)) {
                Write-Host $ScriptStackTrace @color
            }
        } else {
            Write-Host (Get-PSCallStack | Format-Table -AutoSize | Out-String -Width 4096) @color
        }
        Write-Host ('-' * 70) @color

        Write-Host $Exception.StackTrace @color
        Write-Host ('=' * 70) @color

        SetConsoleVisibility $true
    }.GetNewClosure()

    $global:firstChanceExceptionHandler = New-RunspacedDelegate ([ System.EventHandler`1[System.Runtime.ExceptionServices.FirstChanceExceptionEventArgs]] {
        param($Sender, $EventArgs)
        $null = $appExceptionHandler.Invoke($EventArgs.Exception)
    }.GetNewClosure())

    $global:unhandledExceptionHandler = New-RunspacedDelegate ([UnhandledExceptionEventHandler] {
        param($Sender, $EventArgs)
        $null = $appExceptionHandler.Invoke($EventArgs.Exception)
    }.GetNewClosure())

    [System.AppDomain]::CurrentDomain.Add_FirstChanceException($firstChanceExceptionHandler)
    [System.AppDomain]::CurrentDomain.Add_UnhandledException($unhandledExceptionHandler)
}

Function SetPSLogging([bool] $value) {
    $LogCommandHealthEvent = $value
    $LogCommandLifecycleEvent = $value
    $LogEngineHealthEvent = $value
    $LogEngineLifecycleEvent = $value
    $LogProviderHealthEvent = $value
    $LogProviderLifecycleEvent = $value
    $ProgressPreference = 'SilentlyContinue'
    $VerbosePreference = 'SilentlyContinue'
}

Function SetPSLimits {
    $MaximumVariableCount = 32767
    $MaximumFunctionCount = 32767
}

Function global:AtExit {
    foreach ($plugin in $global:plugins) {
        $plugin.Release()
    }

    SaveSettings

    Write-Information 'AtExit has finished.'
}

Function global:Main {
    [CmdletBinding()]
    param([switch] $HideConsole)

    $Debug = $PSBoundParameters['Debug']
    SetPSLogging $false
    SetPSLimits

    if (CheckPipsAlreadyRunning) {
        Exit
    }

    $null = Import-Module -Global .\PSRunspacedDelegate\PSRunspacedDelegate

    $null = SetConsoleVisibility ((-not $HideConsole) -or $Debug)
    if ($Debug) {
        Set-PSDebug -Strict -Trace 0  # -Trace ∈ (0, 1=lines, 2=lines+vars+calls)
        $ConfirmPreference = 'None'
        $DebugPreference = 'Continue'
        $ErrorActionPreference = 'Continue'
        $WarningPreference = 'Continue'
        $InformationPreference = 'Continue'
        InstallDebugHelpers
    } else {
        $null = Set-PSDebug -Off
        $ConfirmPreference = 'None'
        $DebugPreference = 'SilentlyContinue'
        $ErrorActionPreference = 'SilentlyContinue'
        $WarningPreference = 'SilentlyContinue'
        $InformationPreference = 'SilentlyContinue'
    }

    $null = StartPipsSpellingServer

    $null = LoadSettings
    $null = LoadPlugins

    $null = SetWebClientWorkaround

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = CreateMainForm
    $form.Show()
    $form.Activate()

    $appContext = [System.Windows.Forms.ApplicationContext]::new($form)
    $null = [System.Windows.Forms.Application]::Run($appContext)
    [System.Windows.Forms.Application]::Exit()

    AtExit

    if ($Debug) {
        UninstallDebugHelpers
    }

    Write-Information 'Exiting.'
}

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
