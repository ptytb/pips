[Void][Reflection.Assembly]::LoadWithPartialName("System")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Drawing.Size")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Drawing.Point")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.MessageBox")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.FontStyle")
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


Function global:Get-Bin($command, $all = $false) {
    $commands = Get-Command -All -ErrorAction SilentlyContinue $command
    $found = $null
    if ($commands) {
        $commands = $commands | ForEach-Object { $_.Source }
        if ($all) {
            $found = $commands
        } else {
            $found = $commands | Select-Object -Index 0
        }
    }
    return $found
}

Function global:Get-CurrentInterpreter($item) {
    return $Script:interpretersComboBox.SelectedItem."$item"
}

Function Exists-File($path) {
    return [System.IO.File]::Exists($path)
}

Function global:Exists-Directory($path) {
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
$dataGridView = $null
$inputFilter = $null
$logView = $null
$actionsModel = $null
$isolatedCheckBox = $null
$header = ("Select", "Package", "Installed", "Latest", "Type", "Status")
$csv_header = ("Package", "Installed", "Latest", "Type", "Status")
$search_columns = ("Select", "Package", "Version", "Description", "Type", "Status")
$formLoaded = $false
$outdatedOnly = $true
$interpreters = $null

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


Function Set-WebClientWorkaround {
    Function Set-UseUnsafeHeaderParsing    {
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

    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    
    Set-UseUnsafeHeaderParsing -Enable
}

Set-WebClientWorkaround
Function Download-String($url) {
    try {
        $wc = New-Object System.Net.WebClient
        $wc.Headers["User-Agent"] = "Mozilla/5.0 (compatible; MSIE 6.0;)"
        $result = $wc.DownloadString($url)
    } catch {
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

Function global:Write-PipLog() {
    foreach ($obj in $args) {
        $logView.AppendText("$obj")
    }
    $logView.AppendText("`n")
    $logView.ScrollToCaret()
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

Function Add-Button ($name, $handler) {
    $button = New-Object Windows.Forms.Button
    $button.Text = $name
    $button.Add_Click({ $handler.Invoke( @($Script:button) ) }.GetNewClosure())
    Add-TopWidget $button
    return $button
}

Function Add-ButtonMenu ($text, $tools, $onclick) {
    $form = $script:form  # to be captured by $handler's closure
    $interpretersComboBox = $Script:interpretersComboBox  # the same reason
    
    $handler = {
        param($button)
        $menuStrip = New-Object System.Windows.Forms.ContextMenuStrip
        foreach ($tool in $tools) {
            if ($tool.Persistent) {
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
            if ($tool.Persistent) {
                $menuStrip.Items.Add($tool.MenuText)
            }
        }
        
        $tools = $Script:tools
        $onclick = $Script:onclick
        $menuStrip.add_ItemClicked({
            foreach ($tool in $tools) {
                if ($tool.MenuText -eq $_.ClickedItem) {
                    $Script:menuStrip.Hide()
                    $Script:onclick.Invoke( @($tool) )
                }
            }
        }.GetNewClosure())
        
        $point = New-Object System.Drawing.Point ($button.Location.X, $button.Bottom)
        $menuStrip.Show($Script:form.PointToScreen($point))
    }.GetNewClosure()

    $button = Add-Button $text ($handler)
    $button.ImageAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $button.Image = Convert-Base64ToBMP $iconBase64_DownArrow
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
    $input.Add_TextChanged({ $handler.Invoke( @($Script:input) ) }.GetNewClosure())
    Add-TopWidget $input
    return $input
}


Function Add-Buttons {
    Add-Button "Check Updates" { Get-PythonPackages }
    Add-Button "List Installed" { Get-PythonPackages($false) }
    Add-Button "Sel All Visible" { Select-VisiblePipPackages($true) }
    Add-Button "Select None" { Select-PipPackages($false) }
    Add-Button "Check Deps" { Check-PipDependencies }
    Add-Button "Execute" { Execute-PipAction }
}

Function global:Get-PyDoc($request) {
    $requestNormalized = $request -replace '-','_'
    $output = & (Get-CurrentInterpreter 'PythonExe') -m pydoc $requestNormalized

    if ("$output".StartsWith('No Python documentation found')) {
        $output = & (Get-CurrentInterpreter 'PythonExe') -m pydoc ($requestNormalized).ToLower()
    }
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
        $null = $builtinLibs.Add(@{Package=$packageName; Type='builtin'})
    }

    $getBuiltinsScript = "import sys; print(','.join(sys.builtin_module_names))"
    $sys_builtin_module_names = & (Get-CurrentInterpreter 'PythonExe') -c $getBuiltinsScript
    $modules = $sys_builtin_module_names.Split(',')
    foreach ($builtinModule in $modules) {
        if ($trackDuplicates.Contains("$builtinModule")) {
            continue
        }
        $null = $builtinLibs.Add(@{Package=$builtinModule; Type='builtin'})
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
            $null = $otherLibs.Add(@{Package=$packageName; Type='other'})
        }
    }

    return ,$otherLibs
}

Function Get-CondaPackages() {
    $condaPackages = New-Object System.Collections.ArrayList
    $conda_exe = Get-CurrentInterpreter 'CondaExe'

    if ($conda_exe) {
        $arguments =New-Object System.Collections.ArrayList
        $null = $arguments.Add('list')
        $null = $arguments.Add('--json')
        $null = $arguments.Add('--no-pip')
        $null = $arguments.Add('--show-channel-urls')

        # This one sounds nice but could give versions older than installed
        # conda update --dry-run --json --all
        
        # This one sounds nice but could give versions older than installed
        # conda search --outdated

        $items = & $conda_exe $arguments | ConvertFrom-Json 

        foreach ($item in $items) {
            $null = $condaPackages.Add(@{Type='conda'; Package=$item.name; Version=$item.version})
        }
    }

    return ,$condaPackages
}

$actionCommands = @{
    pip=@{
        info          = { return (& (Get-CurrentInterpreter 'PipExe') show              $args 2>&1) };
        documentation = { $null = (Show-DocView $pkg).Show(); return ''    };
        files         = { return (& (Get-CurrentInterpreter 'PipExe') show    --files   $args 2>&1) };
        update        = { return (& (Get-CurrentInterpreter 'PipExe') install -U        $args 2>&1) };
        install       = { return (& (Get-CurrentInterpreter 'PipExe') install           $args 2>&1) };
        install_dry   = { return 'Not supported on pip'                        };
        install_nodep = { return (& (Get-CurrentInterpreter 'PipExe') install --no-deps $args 2>&1) };
        download      = { return (& (Get-CurrentInterpreter 'PipExe') download          $args 2>&1) };
        uninstall     = { return (& (Get-CurrentInterpreter 'PipExe') uninstall --yes   $args 2>&1) };
    };
    conda=@{
        info          = { return (& (Get-CurrentInterpreter 'CondaExe') list      -v --json                     $args 2>&1) };        
        documentation = { return ''                                                                    };
        files         = { return ''                                                                    };
        update        = { return (& (Get-CurrentInterpreter 'CondaExe') update  --yes                           $args 2>&1) };
        install       = { return (& (Get-CurrentInterpreter 'CondaExe') install --yes --no-shortcuts            $args 2>&1) };
        install_dry   = { return (& (Get-CurrentInterpreter 'CondaExe') install --dry-run                       $args 2>&1) };
        install_nodep = { return (& (Get-CurrentInterpreter 'CondaExe') install --yes --no-shortcuts --no-deps  $args 2>&1) };
        download      = { return ''                                                                    };
        uninstall     = { return (& (Get-CurrentInterpreter 'CondaExe') uninstall --yes                         $args 2>&1) };
    };
}
$actionCommands.wheel   = $actionCommands.pip
$actionCommands.sdist   = $actionCommands.pip
$actionCommands.builtin = $actionCommands.pip
$actionCommands.other   = $actionCommands.pip
$actionCommands.git     = $actionCommands.pip

Function Copy-AsRequirementsTxt($list) {
    $requirements = New-Object System.Text.StringBuilder
    foreach ($item in $list) {
        $null = $requirements.AppendLine("$($item.Package)==$(if ($item.Installed) { $item.Installed } else { $item.Version })")
    }
    Set-Clipboard $requirements.ToString()
    Write-PipLog "Copied $($list.Count) items to clipboard."
}

Function Add-ComboBoxActions {
    $actionItemCount = 0
    
    Function Make-PipActionItem($name, $code, $validator, $takesList = $false) {
        $action = New-Object psobject -Property @{Name=$name; TakesList=$takesList; Id=(++$Script:actionItemCount);}
        $action | Add-Member ScriptMethod ToString { "$($this.Name) [F$($this.Id)]" } -Force
        $action | Add-Member ScriptMethod Execute  $code
        $action | Add-Member ScriptMethod Validate $validator
        return $action
    }

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
        { param($pkg,$out); $out -match ('Successfully installed |Installing collected packages:\s*(\s*\S*,\s*)*' + $pkg) } )

    & $Add (Make-PipActionItem 'Install (Dry Run)' `
        { param($pkg,$type); $actionCommands[$type].install_dry.Invoke($pkg) } `
        { param($pkg,$out); $out -match ('Successfully installed |Installing collected packages:\s*(\s*\S*,\s*)*' + $pkg) } )

    & $Add (Make-PipActionItem 'Install (No Deps)' `
        { param($pkg,$type); $actionCommands[$type].install_nodep.Invoke($pkg) } `
        { param($pkg,$out); $out -match ('Successfully installed |Installing collected packages:\s*(\s*\S*,\s*)*' + $pkg) } )

    & $Add (Make-PipActionItem 'Install' {
            param($pkg,$type,$version)
            $git_url = Validate-GitLink $pkg
            if ($git_url) {
                $pkg = $git_url
            }
            if (-not [string]::IsNullOrEmpty($version) -and ($type -ne 'git')) {  # as git version is a timestamp
                $pkg = "$pkg==$version"
            }
            $actionCommands[$type].install.Invoke($pkg) } `
        { param($pkg,$out); $out -match ('Successfully installed |Installing collected packages:\s*(\s*\S*,\s*)*' + $pkg) } )

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

    & $Add (Make-PipActionItem 'Dependency tree' `
        { param($list); Get-DependencyAsciiGraph $list } `
        { param($pkg,$out); $out -match '.*' } )

    $Script:actionsModel = $actionsModel

    $actionListComboBox = New-Object System.Windows.Forms.ComboBox
    $actionListComboBox.DataSource = $actionsModel
    $actionListComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $Script:actionListComboBox = $actionListComboBox
    Add-TopWidget($actionListComboBox)    
    
    return $actionListComboBox
}


$trackDuplicateInterpreters = New-Object System.Collections.Generic.HashSet[String]

Function Get-InterpreterRecord($path, $items, $user = $false) {
    if ($trackDuplicateInterpreters.Contains($path)) {
        return
    }

    Function Guess-EnvPath ($fileName, [switch] $directory) {
        $subdirs = @('\'; '\Scripts\'; '\.venv\Scripts\'; '\.venv\'; '\env\Scripts\'; '\env\'; '\bin\')
        foreach ($tryPath in $subdirs) {
            $target = "${path}${tryPath}${fileName}"
            if ($directory) {
                if (Exists-Directory $target) {
                    return $target
                }
            } else {
                if (Exists-File $target) {
                    return $target
                }
            }
        }
        return $null
    }

    $python = Guess-EnvPath 'python.exe'
    if (-not $python) {
        return
    }
    $versionString = & $python --version 2>&1
    $version = [regex]::Match($versionString, '\s+(\d+\.\d+)').Groups[1]

    $action = New-Object psobject -Property @{
        Path            = $path;
        Version            = "$version";
        Arch            = (Test-is64Bit $python).FileType;
        PythonExe        = $python;
        PipExe            = Guess-EnvPath 'pip.exe';
        CondaExe        = Guess-EnvPath 'conda.exe';
        VirtualenvExe   = Guess-EnvPath 'virtualenv.exe';
        VenvActivate    = Guess-EnvPath 'activate.bat';
        PipenvExe        = Guess-EnvPath 'pipenv.exe';
        RequirementsTxt = Guess-EnvPath 'requirements.txt';
        Pipfile          = Guess-EnvPath 'Pipfile';
        PipfileLock     = Guess-EnvPath 'Pipfile.lock';
        SitePackagesDir = Guess-EnvPath 'Lib\site-packages' -directory;
        User            = $user;
    }
    $action | Add-Member ScriptMethod ToString {
        "{2} [{0}] {1}" -f $this.Arch, $this.PythonExe, $this.Version
    } -Force

    $null = $items.Add($action)
    $null = $trackDuplicateInterpreters.Add($path)

    if ($Script:interpretersComboBox) {
        $interpretersComboBox.DataSource = $null
        $interpretersComboBox.DataSource = $interpreters
    }
}

Function Find-Interpreters {
    $items = New-Object System.Collections.ArrayList

    $list = @((Get-Bin 'python' $true); (dir $env:SystemDrive\Python*) | foreach { "$_\python.exe" })
    foreach ($path in $list) {
        if (-not [string]::IsNullOrEmpty($path)) {
            Get-InterpreterRecord (Split-Path -Parent $path) $items
        }
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
        $reg_pythons_u = dir HKCU:Software\Python | foreach { dir HKCU:$_ } | foreach { (Get-ItemProperty -Path HKCU:$_\InstallPath).'(default)' }
    }
    if (Test-Path HKLM:Software\Python) {
        $reg_pythons_m = dir HKLM:Software\Python | foreach { dir HKLM:$_ } | foreach { (Get-ItemProperty -Path HKLM:$_\InstallPath).'(default)' }
    }
    if (Test-Path HKLM:Software\Wow6432Node\Python) {
        $reg_WoW_pythons = dir HKLM:Software\Wow6432Node\Python | foreach { dir HKLM:$_ } | foreach { (Get-ItemProperty -Path HKLM:$_\InstallPath).'(default)' }
    }
    foreach ($d in @($reg_pythons_u; $reg_pythons_m; $reg_WoW_pythons)) {
        Get-InterpreterRecord ($d -replace '\\$','') $items
    }

    return ,$items  # keep comma to prevent conversion to an @() array
}


Function Add-ComboBoxInterpreters {
    $interpreters = Find-Interpreters

    foreach ($interpreter in $Global:settings.envs) {
        if ($interpreter.User) {
            $interpreter | Add-Member ScriptMethod ToString {
                "{2} [{0}] {1}" -f $this.Arch, $this.PythonExe, $this.Version
            } -Force
            $null = $interpreters.Add($interpreter)
        }
    }
    
    $Script:interpreters = $interpreters
    $interpretersComboBox = New-Object System.Windows.Forms.ComboBox
    $interpretersComboBox.DataSource = $interpreters
    $interpretersComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $Script:interpretersComboBox = $interpretersComboBox
    Add-TopWidget $interpretersComboBox 4
    return $interpretersComboBox
}

Function Add-CheckBox($text, $code) {
    $checkBox = New-Object System.Windows.Forms.CheckBox
    $checkBox.Text = $text
    $checkBox.Add_Click($code)
    Add-TopWidget($checkBox)
    return ($checkBox)
}

Function Toggle-VirtualEnv ($state) {
    if (! $Script:formLoaded) {
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
        Write-PipLog 'virtualenv not found. Run me from where "pip -m virtualenv env" command has been executed.'
        return
    }

    if ($state) {
        $env:_OLD_VIRTUAL_PROMPT = "$env:PROMPT"
        $env:_OLD_VIRTUAL_PYTHONHOME = "$env:PYTHONHOME"
        $env:_OLD_VIRTUAL_PATH = "$env:PATH"

        Write-PipLog ('Activating: ' + $pipEnvActivate)
        . $pipEnvActivate
    }
    else {
        Write-PipLog ('Deactivating: "' + $pipEnvDeactivate + '" and unsetting environment')
        
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

    #Write-PipLog "PROMPT=" $env:PROMPT
    #Write-PipLog "PYTHONHOME=" $env:PYTHONHOME
    #Write-PipLog "PATH=" $env:PATH
}

Function Load-KnownPackageIndex {
    $fs = New-Object System.IO.FileStream "$PSScriptRoot\known-packages.bin", ([System.IO.FileMode]::Open)
    $bf = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
    $index = $bf.Deserialize($fs)
    $fs.Close()
    return ,$index
}

Function ConvertFrom-RegexGroupsToObject($groups) {
    $gitLinkInfo = New-Object PSCustomObject
        
    foreach ($group in $groups) {
        if (-not [string]::IsNullOrEmpty($group.Value) -and $group.Name -ne "0") {
            $gitLinkInfo | Add-Member NoteProperty "$($group.Name)" "$($group.Value)"
        }
    }
    
    return $gitLinkInfo
}

Function global:Validate-GitLink ($url, $AsObject = $false) {
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
    @('Package'; 'Version'; 'Latest'; 'Type'; 'Status') | ForEach-Object {
        if ($dataModel.Columns.Contains($_)) {
            $dataModel.Columns[$_].ReadOnly = -not $enable
        }
    }
}

Function Generate-FormInstall {
    Function Prepare-PackageAutoCompletion {
        $Script:packageIndex = Load-KnownPackageIndex

        $Script:autoCompleteIndex = New-Object System.Windows.Forms.AutoCompleteStringCollection
        if ($packageIndex.Keys -eq $null) {
            $keys = $packageIndex
        } else {
            $keys = $packageIndex.Keys
        }
        # $dbg_n = 0
        [int] $maxLength = 0
        foreach ($item in $keys) {
            # $dbg_n++
            # if ($dbg_n -gt 1000) {
            #     break
            # }
            $autoCompleteIndex.Add($item)
            $maxLength = [Math]::Max($maxLength, $item.Length)
        }
        
        # Write-PipLog $maxLength         
        # Populate the Dictionary<WordLength: int, WordList: List[string]>
        # for faster selection within Levenshtein distance (+/-) range
        $Global:TypoErrorTable = New-Object 'System.Collections.Generic.Dictionary[[string],[System.Collections.Generic.List[string]]]'
        for ($i = 1; $i -le $maxLength; $i++) {
            $list = New-Object 'System.Collections.Generic.List[string]'
            $null = $Global:TypoErrorTable.Add($i, $list)             
        }
        foreach ($item in $keys) {
            $itemLength = $item.Length
            $list = $Global:TypoErrorTable[$itemLength]
            $list.Add($item)
        }
    }

    if ($Script:packageIndex -eq $null) {
        Prepare-PackageAutoCompletion
    }

    $form = New-Object System.Windows.Forms.Form
    $form.KeyPreview = $true
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Type package names, git urls, local paths and hit Enter after each."
    $label.Location = New-Object Drawing.Point 7,7
    $label.Size = New-Object Drawing.Point 340,24
    $form.Controls.Add($label)

    $cb = New-Object System.Windows.Forms.TextBox
    
    $autoCompleteIndex = $Script:autoCompleteIndex
    $FuncGuessAutoCompleteMode = {
        $text = $cb.Text
        $n = $text.LastIndexOfAny('\/')
        
        if ($text.Contains('==') -or $text.Contains('@')) {  # in cases when file path contains these, won't work; need state tracking
            if (($cb.AutoCompleteSource -eq [System.Windows.Forms.AutoCompleteSource]::FileSystemDirectories) `
                -or $cb.AutoCompleteCustomSource.Equals($Script:autoCompleteIndex) `
                -or ($cb.AutoCompleteCustomSource.Equals($null))) {
                $cb.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
                $cb.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource                
                
                if ($text.Contains('==')) {  # completion for pip package version
                    $packageName = $text -replace '==.*',''
                    
                    if (-not $Global:PyPiPackageJsonCache.ContainsKey($packageName)) {
                        Download-PythonPackageDetails $packageName
                    }
                    
                    $releases = $Global:PyPiPackageJsonCache[$packageName].'releases' | Get-Member -Type Properties | ForEach-Object { $_.Name }
                    $completions_format = "{0}=={1}"
                }
                else {  # completion for git repo tags
                    $packageName = $text -replace '@.*',''
                    $gitLinkInfo = Validate-GitLink $packageName -AsObject $true
                    $releases = Get-GithubRepoTags $gitLinkInfo
                    $completions_format = "{0}@{1}"
                }
                
                $autoCompletePackageVersion = New-Object System.Windows.Forms.AutoCompleteStringCollection                
                foreach ($release in $releases) {
                    $autoCompletePackageVersion.Add($completions_format -f @($packageName,$release))
                }
                $cb.AutoCompleteCustomSource = $autoCompletePackageVersion
            }
        } elseif (($n -gt -1) -and (Exists-Directory $text.Substring(0, $n + 1))) {
            if ($cb.AutoCompleteSource -ne [System.Windows.Forms.AutoCompleteSource]::FileSystemDirectories) {
                $cb.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
                $cb.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::FileSystemDirectories
                $cb.AutoCompleteCustomSource = $null
            }
        } else {
            if (-not ($cb.AutoCompleteCustomSource.Equals($Script:autoCompleteIndex))) {
                $cb.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
                $cb.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource
                $cb.AutoCompleteCustomSource = $Script:autoCompleteIndex
            }
        }
    }.GetNewClosure()
    
    $hint = $null
    $FuncShowToolTip = {
        param($title, $text)
        $hint = New-Object System.Windows.Forms.ToolTip
        $hint.IsBalloon = $true
        $hint.ToolTipTitle = $title
        $hint.ToolTipIcon = [System.Windows.Forms.ToolTipIcon]::Error
        $hint.Show([string]::Empty, $cb, 0);
        $hint.Show($text, $cb, 0, $cb.Height);
        $Script:hint = $hint
    }

    $FuncCleanupToolTip = {
        if ($Script:hint) {
            $Script:hint.Dispose()
            $Script:hint = $null
            return $true
        }
        return $false
    }
    
    $FuncAddInstallSource = {
        param($package)
        
        if ($dataModel.Rows.Count -eq 0) {
            Init-PackageSearchColumns $dataModel
        }
        
        $link = Validate-GitLink $package
        if ($link) {
            if ((Test-PackageInList $link) -ne -1) {
                & $FuncShowToolTip "$package" "Repository '$link' is already in the list"
                return $false
            }
            $row = $dataModel.NewRow()
            $row.Package = $link
            $row.Type = 'git'
            $row.Status = 'Pending'
            $row.Select = $true
            $dataModel.Rows.InsertAt($row, 0)
            return $true
        }
        
        $version = [string]::Empty
        $package_with_version = [regex] '^(?<Name>[^=]+)==(?<Version>[^=]+)$'
        $pv_match = $package_with_version.Match($package)
        $pv_group = $pv_match.Groups
        if ($pv_group.Count -gt 1) {
            ($package, $version) = ($pv_group['Name'].Value, $pv_group['Version'].Value)
        }

        if (-not ($packageIndex.Contains($package))) {
            return $false
        }
        
        $nAlreadyInList = Test-PackageInList $package        
        if ($nAlreadyInList -ne -1) {
            $oldRow = $dataModel.Rows[$nAlreadyInList]
            
            if (      -not [string]::IsNullOrEmpty($oldRow.Version)) {
                $oldVersion = $oldRow.Version
            } elseif (-not ($oldRow.Latest -eq $null)) {
                $oldVersion = $oldRow.Latest
            } elseif (-not [string]::IsNullOrEmpty($oldRow.Installed)) {
                $oldVersion = $oldRow.Installed
            } else {
                $oldVersion = [string]::Empty
            }
            
            #Write-Host "old='$oldVersion', new='$version'"
            
            $IsDifferentVersion = $version -ne $oldVersion
        } else {
            $IsDifferentVersion = $true
        }

        if ($nAlreadyInList -ne -1) {
            if (-not $IsDifferentVersion) {            
                & $FuncShowToolTip "$package" "Package '$package' is already in the list"
                return $false
            } else {
                $row = $oldRow
                Set-PackageListEditable $true
            }
        } else {
            $row = $dataModel.NewRow()
        }
        
        $row.Select = $true
        $row.Package = $package

        if ($dataModel.Columns.Contains('Version')) {
            $row.Version = $version
        } else {
            $row.Latest = $version
        }

        $row.Type = 'pip'
        $row.Status = 'Pending'
        
        if ($nAlreadyInList -eq -1) {
            $dataModel.Rows.InsertAt($row, 0)
        } else {
            Set-PackageListEditable $false
        }
        
        return $true
    }

    $cb.add_TextChanged({
        & $FuncGuessAutoCompleteMode
    })
    
    $cb.Location = New-Object Drawing.Point 7,35
    $cb.Size = New-Object Drawing.Point 330,32
    $form.Controls.Add($cb)

    $form.add_KeyDown({
        $null = & $FuncCleanupToolTip

        if ($_.KeyCode -eq 'Escape') {
            if (($cb.Text.Length -gt 0) -or (& $FuncCleanupToolTip)) {
                $cb.Text = [string]::Empty
            } else {
                $form.Close()
            }
        }
            
        if ($_.KeyCode -eq 'Enter') { 
            $text = $cb.Text.Trim()
            
            if ([string]::IsNullOrEmpty($text)) {
                return
            }
            
            if (-not (Test-KeyPress -Keys ShiftKey)) {
                $okay = & $FuncAddInstallSource $text
                if ($okay) {
                   $cb.Text = [string]::Empty
                   return
                }
            } else {
                $response = ([System.Windows.Forms.MessageBox]::Show(
                    "Search for similar package names?`n`nShould take around 15 sec.",
                    "Fuzzy name search", [System.Windows.Forms.MessageBoxButtons]::YesNo))
                if ($response -ne 'Yes') {
                    return
                }
            
                $candidates = Get-TypoErrorCandidates $text
                if ($candidates.Count -le 10) {
                    $candidatesToolTipText = "$($candidates -join "`n")"
                } else {                     
                    $candidatesToolTipText = "$(($candidates | Select-Object -First 10) -join "`n")`n...`n`nfull list in the log"
                }
                & $FuncShowToolTip "$text" "Packages with similar names found in the index.`n`nDid you mean:`n`n$candidatesToolTipText"
                Write-PipLog "Fuzzy search candidates: $($candidates -join ', ')`n"             
            }
        }        
    }.GetNewClosure())

    $install = New-Object System.Windows.Forms.Button
    $install.Text = "Install"
    $install.Location = New-Object Drawing.Point 140,65
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

    $form.Size = New-Object Drawing.Point 360,140
    $form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Hide
    $form.FormBorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $form.Text = 'Install packages | [Shift+Enter] fuzzy name search'
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false
    $form.Icon = $Script:form.Icon
    $form.ShowDialog()
}

Function Request-UserString($message, $title, $default, $completionItems = $null) {
    $Form                            = New-Object system.Windows.Forms.Form
    $Form.ClientSize                 = '421,247'
    $Form.text                       = $title
    $Form.TopMost                    = $false
    $Form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Hide
    $Form.FormBorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
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
        return $TextBox1.Text         
    } else {
        return $null
    }         
}

Function Generate-FormSearch {
    $message = "Enter keywords to search PyPi, Conda, Github`n`nChecked items will be kept in the search list"
    $title = "pip, conda, github search"
    $default = ""
    $input = Request-UserString $message $title $default    
    if (-not $input) {
        return
    }

    Write-PipLog "Searching for $input"
    Write-PipLog 'Double click or [Ctrl+Enter] a table row to open PyPi, Anaconda.com or github.com in browser'
    $stats = Get-SearchResults $input
    Write-PipLog "Found $($stats.Total) packages: $($stats.PipCount) pip, $($stats.CondaCount) conda, $($stats.GithubCount) github. Total $($dataModel.Rows.Count) packages in list."
    Write-PipLog
}

Function Init-PackageGridViewProperties() {
    $dataGridView.MultiSelect = $false
    $dataGridView.SelectionMode = [System.Windows.Forms.SelectionMode]::One
    $dataGridView.ColumnHeadersVisible = $true
    $dataGridView.RowHeadersVisible = $false
    $dataGridView.ReadOnly = $false
    $dataGridView.AllowUserToResizeRows = $false
    $dataGridView.AllowUserToResizeColumns = $false
    $dataGridView.VirtualMode = $true
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

Function Highlight-PythonPackages {
    if (! $outdatedOnly) {
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
    if ($url -match '^https?://') {
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
               $Global:PyPiPackageJsonCache.Add($packageName, $info)
        }
    }
}

Function global:Update-PythonPackageDetails {
    $_ = $args[0]
    $viewRow = $dataGridView.Rows[$_.RowIndex]
    $rowItem = $viewRow.DataBoundItem
    $cells = $dataGridView.Rows[$_.RowIndex].Cells
    $packageName = $rowItem.Row.Package

    if (! [String]::IsNullOrEmpty($cells['Package'].ToolTipText) -or (!($rowItem.Row.Type -in @('pip', 'wheel', 'sdist')))) {
        return
    }

    if ($Global:PyPiPackageJsonCache.ContainsKey($packageName)) {
        $info = $Global:PyPiPackageJsonCache[$packageName]
        $cells['Package'].ToolTipText = Format-PythonPackageToolTip $info
        return
    }

    if (!(Test-KeyPress -Keys ShiftKey)) {
        return
    }

    $dataModel.Columns['Status'].ReadOnly = $false
    $cells['Status'].Value = 'Fetching...'
    $dataModel.Columns['Status'].ReadOnly = $true
    $viewRow.DefaultCellStyle.BackColor = [Drawing.Color]::Gray

    Run-SubProcessWithCallback ({
        # Worker: Separate process
        param($params)
        [Void][Reflection.Assembly]::LoadWithPartialName("System.Net")
        [Void][Reflection.Assembly]::LoadWithPartialName("System.Net.WebClient")
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
        $Global:dataModel.Columns['Status'].ReadOnly = $false
        $row.Status = if ($okay) { 'OK' } else { 'Failed' }
        $Global:dataModel.Columns['Status'].ReadOnly = $true        
    }) @{PackageName=$packageName; RowIndex=$_.RowIndex}
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
            $Script:interpretersComboBox.SelectedIndex = $i
            Write-PipLog "Switching to env '$path'"
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
            IsAccessible = { [bool] (Get-CurrentInterpreter 'VirtualenvExe') };
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
            IsAccessible = { [bool] (Get-CurrentInterpreter 'PipenvExe') };
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
                Write-PipLog "Firewall rule '$ruleName' was successfully created."
            } else {
                Write-PipLog "Error while creating firewall rule '$ruleName'."
            }
        }
    }
    
    $FuncGetPythonInfo = {
        return (Get-CurrentInterpreter 'Version')
    }

    $menuclick = {
        param($tool)

        $path = Request-FolderPathFromUser `
            "New python environment with active version $($FuncGetPythonInfo.Invoke()) will be created"
        if ($path -eq $null) { return }
        Write-PipLog "Create $($tool.MenuText), please wait..."        
        $output = $tool.Code.Invoke( @($path) )
        Write-PipLog (Tidy-Output $output)

        $FuncUpdateInterpreters.Invoke($path)
    }.GetNewClosure()
    
    $createEnvButton = Add-ButtonMenu 'env: Create' $tools $menuclick
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
    )
    
    $menuclick = {
        param($item)
        $output = $item.Code.Invoke()
    }
    
    $envToolsButton = Add-ButtonMenu 'env: Tools' $menu $menuclick
}

Function global:Get-PyDocTopics() {
    $pythonExe = Get-CurrentInterpreter 'PythonExe'
    if ([string]::IsNullOrEmpty($pythonExe)) {
        Write-PipLog 'No python executable found.'
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
        Write-PipLog 'No python executable found.'
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
                Write-PipLog "Searching apropos for $input"
                $apropos = Get-PyDocApropos $input
                if ($apropos -and $apropos.Count -gt 0) {
                    Write-PipLog "Found $($apropos.Count) topics"
                    $docView = Show-DocView -SetContent ($apropos -join "`n") -Highlight $Script:pyRegexNameChain -NoDefaultHighlighting
                    $docView.Show()
                } else {
                    Write-PipLog 'Nothing found.'
                }
            };
        };
    )
    
    $menuclick = {
        param($item)
        $output = $item.Code.Invoke()
    }
    
    $envToolsButton = Add-ButtonMenu 'Tools' $menu $menuclick
}

Function global:Show-CurrentPackageInBrowser() {
    $view_row = $dataGridView.CurrentRow
    if ($view_row) {
        $row = $view_row.DataBoundItem.Row
        $packageName = $row.Package
        $urlName = [System.Web.HttpUtility]::UrlEncode($packageName)
        if ($row.Type -eq 'conda') {
            $url = $anaconda_url
        } elseif ($row.Type -eq 'git') {
            $gitLinkInfo = Validate-GitLink $packageName -AsObject $true
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
    $form.topmost = $false
    $form.KeyPreview = $true
    $form.Icon = Convert-Base64ToICO $iconBase64_Snakes
    $Script:form = $form

    $null = Add-Buttons

    $actionListComboBox = Add-ComboBoxActions    

    $group = New-Object System.Windows.Forms.Panel
    $group.Location = New-Object System.Drawing.Point 502,2
    $group.Size = New-Object System.Drawing.Size 202,28
    $group.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $form.Controls.Add($group)
    
    $Script:isolatedCheckBox = Add-CheckBox 'isolated' { Toggle-VirtualEnv $Script:isolatedCheckBox.Checked }
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.SetToolTip($isolatedCheckBox, "Ignore environmental variables, user configuration and global packages.`n`n--isolated`n--local")

    $null = Add-Button "Search..." { Generate-FormSearch }
    $null = Add-Button "Install..." { Generate-FormInstall }    
    Add-ToolsButtonMenu

    $null = NewLine-TopLayout

    $null = Add-Label "Filter results:"
    
    $inputFilter = Add-Input {  # TextChanged Handler here
        param($input)
        
        if ($Script:dataGridView.CurrentRow) {
            # Keep selection while filter is being changed
            $selectedRow = $Script:dataGridView.CurrentRow.DataBoundItem.Row
        }

        $searchText = $input.Text -replace "'","''"
        
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
        $isInstallMode = $dataModel.Columns.Contains('Description')
        if ($isInstallMode) {
            $query = "($subQueryPackage) OR ($subQueryDescription)"
        } else {
            $query = "$subQueryPackage"
        }

        #Write-Host $query
        
        if ($searchText.Length -gt 0) {
            $Script:dataModel.DefaultView.RowFilter = $query
        } else {
            $Script:dataModel.DefaultView.RowFilter = $null
        }

        if ($selectedRow) {
            Set-SelectedRow $selectedRow
        }

        Highlight-PythonPackages
    }
    $Script:inputFilter = $inputFilter
    $toolTipFilter = New-Object System.Windows.Forms.ToolTip
    $toolTipFilter.SetToolTip($inputFilter, "Esc to clear")

    $searchMethodComboBox = New-Object System.Windows.Forms.ComboBox
    $searchMethods = New-Object System.Collections.ArrayList
    $null = $searchMethods.Add('Whole Phrase')
    $null = $searchMethods.Add('Any Word')
    $null = $searchMethods.Add('All Words')
    $null = $searchMethods.Add('Exact Match')
    $searchMethodComboBox.DataSource = $searchMethods
    $searchMethodComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $Script:searchMethodComboBox = $searchMethodComboBox
    Add-TopWidget($searchMethodComboBox)
    $searchMethodComboBox.add_SelectionChangeCommitted({
        $flt = $Script:inputFilter
        $t = $flt.Text
        $flt.Text = [String]::Empty
        $flt.Text = $t
        })

    $labelInterp   = Add-Label "Active Interpreter:"
    $toolTipInterp = New-Object System.Windows.Forms.ToolTip
    $toolTipInterp.SetToolTip($labelInterp, "Ctrl+C to copy selected path")

    $interpretersComboBox = Add-ComboBoxInterpreters
    $null = Add-Button "env: Open..." {
        $path = Request-FolderPathFromUser ("Choose a folder with python environment, created by either Virtualenv or pipenv`n`n" +
            "Typically it contains dirs: Include, Lib, Scripts")
        if ($path) {
            $oldCount = $interpreters.Count
            Get-InterpreterRecord $path $interpreters -user $true
            if (($interpreters.Count -gt $oldCount) -or ($trackDuplicateInterpreters.Contains($path))) {
                if ($interpreters.Count -gt $oldCount) {
                    Write-PipLog "Added virtual environment location: $path"
                }
                Set-ActiveInterpreterWithPath $path                
            } else {
                Write-PipLog "No python found in $path"
            }
        }
    }
    
    Add-CreateEnvButtonMenu
    Add-EnvToolButtonMenu
    
    $interpreters = $Script:interpreters
    $trackDuplicateInterpreters = $Script:trackDuplicateInterpreters
    $form.add_KeyDown({
        if ($Script:interpretersComboBox.Focused) {
            if (($_.KeyCode -eq 'C') -and $_.Control) {
                $python_exe = Get-CurrentInterpreter 'PythonExe'
                Set-Clipboard $python_exe
                Write-PipLog "Copied to clipboard: $python_exe"
            }
            if ($_.KeyCode -eq 'Delete') {
                if (-not (Get-CurrentInterpreter 'User')) {
                    Write-PipLog 'Can only delete venv which was added manually with env:Open or env:Create.'
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
                        Write-PipLog "Cannot delete '$path'"
                        $failedToRemove = $true
                    } else {
                        Write-PipLog "Removed $($stats.Count) items."
                    }
                }
                
                if ((($response -eq 'No') -or ($response -eq 'Yes')) -and -not $failedToRemove) {
                    $interpreters.Remove($interpretersComboBox.SelectedItem)                    
                    $Script:trackDuplicateInterpreters.Remove($path)
                    $interpretersComboBox.DataSource = $null
                    $interpretersComboBox.DataSource = $interpreters                     
                    Write-PipLog "Removed venv '${path}' from list."
                    
                    $interpretersComboBox.SelectedIndex = 0
                    Write-PipLog "Switching to '$(Get-CurrentInterpreter 'Path')'"
                }
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
            
            if ($inputFilter.Focused) {
                if ([string]::IsNullOrEmpty($inputFilter.Text)) {
                    $dataGridView.Focus()                     
                } else {
                    $inputFilter.Text = [String]::Empty                     
                }
            } else {
                $inputFilter.Focus()
            }
        }
        if ($_.KeyCode -eq 'Return') {
            if ($inputFilter.Focused) {
                $_.Handled = $true
                $dataGridView.Focus()
            }
        }
    }.GetNewClosure())
    Init-PackageGridViewProperties
    
    $dataModel = New-Object System.Data.DataTable
    $dataGridView.DataSource = $dataModel    
    $Script:dataModel = $dataModel
    $Global:dataModel = $dataModel
    Init-PackageUpdateColumns $dataModel

    $form.Controls.Add($dataGridView)

    $logView = New-Object System.Windows.Forms.RichTextBox
    $logView.Location = New-Object Drawing.Point 7,520
    $logView.Size = New-Object Drawing.Point 800,270
    $logView.ReadOnly = $true
    $logView.Multiline = $true
    $logView.Font = New-Object System.Drawing.Font("Consolas", 11)
    $Script:logView = $logView
    $form.Controls.Add($logView)

    $FuncHighlightLogFragment = {
        if ($Script:dataModel.Rows.Count -eq 0) {
            return
        }
        
        $viewRow = $Script:dataGridView.CurrentRow
        if (! $viewRow) {
            return
        }
        $row = $viewRow.DataBoundItem.Row        

        $Script:logView.SelectAll()
        $Script:logView.SelectionBackColor = $Script:logView.BackColor

        if (Get-Member -inputobject $row -name "LogFrom" -Membertype Properties) {
            $Script:logView.Select($row.LogFrom, $row.LogTo)
            $Script:logView.SelectionBackColor = [Drawing.Color]::Yellow
            $Script:logView.ScrollToCaret()
        }
    }
    
    $dataGridView.Add_CellMouseClick({ & $FuncHighlightLogFragment }.GetNewClosure())
    $dataGridView.Add_SelectionChanged({ & $FuncHighlightLogFragment }.GetNewClosure())

    $dataGridView.Add_CellMouseDoubleClick({
            if (($_.RowIndex -gt -1) -and ($_.ColumnIndex -gt 0)) {
                Show-CurrentPackageInBrowser
            }
        }.GetNewClosure())
    $form.Add_Load({ $Script:formLoaded = $true })
    
    $FuncResizeForm = {
        $dataGridView.Width = $form.ClientSize.Width - 15
        $dataGridView.Height = $form.ClientSize.Height / 2
        $logView.Top = $dataGridView.Bottom + 15
        $logView.Width = $form.ClientSize.Width - 15
        $logView.Height = $form.ClientSize.Height - $dataGridView.Bottom - $lastWidgetTop
    }

    & $FuncResizeForm
    $form.Add_Resize({ & $FuncResizeForm }.GetNewClosure())
    $form.Add_Shown({
        Write-PipLog 'Hold Shift and hover the rows to fetch the detailed info'
        $form.BringToFront()
        })
        
    $form.add_KeyDown({
        if ($_.KeyCode -in (1..12 | ForEach-Object { "F$_" })) {  # Handle F1..F12 functional keys
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
                Execute-PipAction
                $_.Handled = $true
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

        $this.docView = New-Object System.Windows.Forms.RichTextBox
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

        $this.formDoc.add_KeyDown({
                if ($_.KeyCode -eq 'Escape') {
                    $self.formDoc.Close()
                }
                if ($_.KeyCode -eq 'Enter') {
                    $charIndex = $self.docView.SelectionStart
                    $jumpToWord.Invoke($charIndex)
                }
                if ($_.KeyCode -eq 'Space') {
                    [System.Windows.Forms.SendKeys]::Send('{PGDN}')
                }
                if (($_.KeyCode -eq 'F') -and $_.Control) {
                    [System.Windows.Forms.SendKeys]::Send('{PGDN}')
                    $_.Handled = $true
                }
                if (($_.KeyCode -eq 'B') -and $_.Control) {
                    [System.Windows.Forms.SendKeys]::Send('{PGUP}')
                    $_.Handled = $true
                }
            }.GetNewClosure())

        $this.docView.add_LinkClicked({
                Open-LinkInBrowser $_.LinkText
            }.GetNewClosure())

        $this.docView.Add_MouseClick({
                $clickedIndex = $self.docView.GetCharIndexFromPosition($_.Location)
                $jumpToWord.Invoke($clickedIndex)
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

        foreach ($match in $matches.Groups) {
            if ($match.Name -eq 0) {
                continue
            }
            $this.docView.Select($match.Index + $this.modifiedTextLengthDelta, $match.Length)
            $selectionAlteringCode.Invoke($match.Index + $this.modifiedTextLengthDelta, $match.Length, $match.Value)
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

Function global:Show-DocView($packageName, $SetContent = $null, $Highlight = $null, [switch] $NoDefaultHighlighting) {
    if (-not $SetContent) {
        $content = (Get-PyDoc $packageName)
    } else {
        $content = $SetContent
    }
    
    $viewer = New-Object DocView -ArgumentList @((Tidy-Output $content), $packageName, $NoDefaultHighlighting)
    
    if ($Highlight) {
        $viewer.Highlight_Text($Highlight, ([System.Drawing.Color]::Navy), $false, $true)
    }
    
    return $viewer
}

Function Write-PipPackageCounter {
    $count = $dataModel.Rows.Count
    Write-PipLog "Now $count packages in the list."
}

Function Store-CheckedPipSearchResults() {
    $selected = New-Object System.Data.DataTable
    Init-PackageSearchColumns $selected

    $isInstallMode = $dataModel.Columns.Contains('Description')
    if ($isInstallMode) {
        foreach ($row in $dataModel) {
            if ($row.Select) {
                $selected.ImportRow($row)
            }
        }
    }

    return ,$selected
}

Function Get-PipSearchResults($request) {
    $pip_exe = Get-CurrentInterpreter 'PipExe'
    if (!$pip_exe) {
        Write-PipLog 'pip is not found!'
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
        $row = $dataModel.NewRow()
        $row.Select = $false
        $row.Package = $m.Groups[1].Value
        $row.Version = $m.Groups[2].Value
        $row.Description = $m.Groups[3].Value
        $row.Type = 'pip'
        $row.Status = ''
        $dataModel.Rows.Add($row)
        $count += 1
    }

    return $count
}

Function Get-CondaSearchResults($request) {
    $conda_exe = Get-CurrentInterpreter 'CondaExe'
    if (! $conda_exe) {
        Write-PipLog 'conda is not found!'
        return 0
    }
    $arch = Get-CurrentInterpreter 'Arch'

    # --info should give better details but not supported on every conda
    $items = & $conda_exe search --json $request | ConvertFrom-Json

    $count = 0
    $items.PSObject.Properties | ForEach-Object {
        $name = $_.Name
        $item = $_.Value

        $row = $dataModel.NewRow()
        $row.Select = $false
        $row.Package = $name

        if ($item.GetType().BaseType -eq [System.Array]) {
            # If we've a list of versions, take only the first for now
            $item = $item[0]
        }

        $row.Version = $item.version
        $row.Description = $item.license_family
        $row.Type = 'conda'
        $row.Status = ''
        $dataModel.Rows.Add($row)

        $count += 1
    }

    return $count
}

Function Get-GithubSearchResults ($request) {
    $json = Download-String ($github_search_url -f [System.Web.HttpUtility]::UrlEncode($request))
    $info = $json | ConvertFrom-Json
    $items = $info.'items'
    $count = 0
    $items | ForEach-Object {
        $row = $dataModel.NewRow()
        $row.Select = $false
        $row.Package = $_.'full_name'
        $row.Version = ($_.'pushed_at' -replace 'T',' ') -replace 'Z',''
        $row.Description = "$($_.'stargazers_count') $([char] 0x2729) $($_.'forks') $([char] 0x2442) $($_.'open_issues') $([char] 0x2757) $($_.'description')"
        $row.Type = 'git'
        $row.Status = ''
        $dataModel.Rows.Add($row)
        $count++
    }
    return $count
}

Function global:Get-GithubRepoTags($gitLinkInfo) {
    if (-not $gitLinkInfo) {
        return $null
    }
    
    if ($gitLinkInfo.Path) {
        $git = Get-Bin 'git'
        if (-not $git) {
            return $null
        }
        return (& $git -C $gitLinkInfo.Path tag)
    }
    
    $github_tags_url = 'https://api.github.com/repos/{0}/{1}/tags'
    $url = $github_tags_url -f $gitLinkInfo.User,$gitLinkInfo.Repo
    $json = Download-String $url
    if (-not [string]::IsNullOrEmpty($json)) {
        $tags = $json | ConvertFrom-Json | ForEach-Object { $_.Name }
        return $tags
    }
    
    return $null
}

Function Get-SearchResults($request) {
    $previousSelected = Store-CheckedPipSearchResults    
    Clear-Rows
    Init-PackageSearchColumns $dataModel

    $dataGridView.BeginInit()
    $dataModel.BeginLoadData()
    
    foreach ($row in $previousSelected) {
        $dataModel.ImportRow($row)
    }
    
    $pipCount = Get-PipSearchResults $request
    $condaCount = Get-CondaSearchResults $request
    $githubCount = Get-GithubSearchResults $request

    $dataModel.EndLoadData()
    $dataGridView.EndInit()

    return @{PipCount=$pipCount; CondaCount=$condaCount; GithubCount=$githubCount; Total=($pipCount + $condaCount + $githubCount)}
}

 Function Get-PythonPackages($outdatedOnly = $true) {
    Write-PipLog
    Write-PipLog 'Updating package list... '
    
    $python_exe = Get-CurrentInterpreter 'PythonExe'
    $pip_exe = Get-CurrentInterpreter 'PipExe'
    $conda_exe = Get-CurrentInterpreter 'CondaExe'
    
    if ($python_exe) {
        Write-PipLog (& $python_exe --version 2>&1)
    } else {
        Write-PipLog 'Python is not found!'
    }
    
    if ($pip_exe) {
        Write-PipLog (& $pip_exe --version 2>&1)
    } else {
        Write-PipLog 'pip is not found!'
    }
    
    Write-PipLog

    Clear-Rows
    Init-PackageUpdateColumns $dataModel

    
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
            $row = $dataModel.NewRow()        
            $row.Select = $false
            $row.Package = $packages[$n].Package
            $row.Installed = $packages[$n].Installed
            $row.Latest = $packages[$n].Latest
            $row.Type = $packages[$n].Type
            if (! [String]::IsNullOrEmpty($packages[$n].Version)) {
                $row.Installed = $packages[$n].Version
            }
            if ([String]::IsNullOrEmpty($row.Type)) {
                $row.Type = $defaultType
            } 
            $dataModel.Rows.Add($row)
        }        
    }

    $dataModel.BeginLoadData()
    if ($pip_exe) {
        Add-PackagesToTable $pipPackages 'pip'
    }
    if (! $outdatedOnly) {
        $builtinPackages = Get-PythonBuiltinPackages
        Add-PackagesToTable $builtinPackages 'builtin'

        $otherPackages = Get-PythonOtherPackages
        Add-PackagesToTable $otherPackages 'other'
    }
    if ($conda_exe) {
        $condaPackages = Get-CondaPackages
        Add-PackagesToTable $condaPackages 'conda'
    }
    $dataModel.EndLoadData()

    $Script:outdatedOnly = $outdatedOnly
    Highlight-PythonPackages

    Write-PipLog 'Package list updated.'
    Write-PipLog 'Double click or [Ctrl+Enter] a table row to open PyPi, Anaconda.com or github.com in browser'
    
    $count = $dataModel.Rows.Count
    $pipCount = $pipPackages.Count
    $builtinCount = $builtinPackages.Count
    $condaCount = $condaPackages.Count
    $otherCount = $otherPackages.Count
    Write-PipLog "Total $count packages: $builtinCount builtin, $pipCount pip, $condaCount conda, $otherCount other"
    Write-PipLog
}

Function Select-VisiblePipPackages($value) {
    $dataModel.BeginLoadData()
    for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
        if ($dataGridView.Rows[$i].DataBoundItem.Row.Type -in @('builtin', 'other') ) {
            continue
        }
        $dataGridView.Rows[$i].DataBoundItem.Row.Select = $value
    }
    $dataModel.EndLoadData()
}

Function Select-PipPackages($value) {
    $dataModel.BeginLoadData()
    for ($i = 0; $i -lt $dataModel.Rows.Count; $i++) {
       $dataModel.Rows[$i].Select = $value
    }
    $dataModel.EndLoadData()
}

Function Set-Unchecked($index) {
    $dataModel.Rows[$index].Select = $false
}

Function Test-PackageInList($name) {
    $n = 0
    foreach ($item in $Script:dataModel.Rows) {
        if ($item.Package -eq $name) {
            return $n
        }
        $n++
    }
    return -1
}

Function global:Tidy-Output($text) {
    return ($text -replace '$', "`n")
}

Function Clear-Rows() {
    $Script:outdatedOnly = $true
    $Script:inputFilter.Clear()
    $dataModel.DefaultView.RowFilter = $null
    $dataGridView.ClearSelection()
    
    #if ($dataGridView.SortedColumn) {
    #    $dataGridView.SortedColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    #}

    $dataModel.DefaultView.Sort = [String]::Empty

    $dataGridView.BeginInit()
    $dataModel.BeginLoadData()
    
    $dataModel.Rows.Clear()

    $dataModel.EndLoadData()
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

Function Set-SelectedRow($selectedRow) {
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
    Write-PipLog 'Checking dependencies...'

    $pip_exe = Get-CurrentInterpreter 'PipExe'
    if (!$pip_exe) {
        Write-PipLog 'pip is not found!'
        return
    }

    $result = & $pip_exe check 2>&1
    $result = Tidy-Output $result
    
    if ($result -match 'No broken requirements found') {
        Write-PipLog "OK"
        Write-PipLog $result
    } else {
        Write-PipLog "NOT OK"
        Write-PipLog $result
    }
}

Function global:Select-PipAction($actionName) {
    $n = 0
    foreach ($item in $actionsModel) {
        if ($item.Name -eq $actionName) {
            $actionListComboBox.SelectedIndex = $n
            return
        }
        $n++
    }
}

Function global:Execute-PipAction {
    $action = $Script:actionsModel[$actionListComboBox.SelectedIndex]
    
    if ($action.TakesList) {
        $checkedList = New-Object System.Collections.ArrayList
    }
    
    $tasksOkay = 0
    $tasksFailed = 0
    
    for ($i = 0; $i -lt $dataModel.Rows.Count; $i++) {
       if ($dataModel.Rows[$i].Select -eq $true) {
            Set-SelectedRow $dataModel.Rows[$i]
            [System.Windows.Forms.Application]::DoEvents()
            
            $package = $dataModel.Rows[$i]
            
            if (-not $action.TakesList) {
                Write-PipLog ""
                Write-PipLog $action.Name ' ' $package.Package
                
                $version = if ($package.Version) { $package.Version } else { $package.Latest }
                $result = $action.Execute($package.Package, $package.Type, $version)
                
                $logFrom = $Script:logView.TextLength 
                Write-PipLog (Tidy-Output $result)
                $logTo = $Script:logView.TextLength - $logFrom
                $dataModel.Rows[$i] | Add-Member -Force -MemberType NoteProperty -Name LogFrom -Value $logFrom
                $dataModel.Rows[$i] | Add-Member -Force -MemberType NoteProperty -Name LogTo -Value $logTo

                $dataModel.Columns['Status'].ReadOnly = $false
                if ($action.Validate($package.Package, $result)) {
                    $dataModel.Rows[$i].Status = "OK"
                    Set-Unchecked $i
                    $tasksOkay++
                } else {
                    $dataModel.Rows[$i].Status = "Failed"
                    $tasksFailed++
                }
                $dataModel.Columns['Status'].ReadOnly = $true
            } else {
                $null = $checkedList.Add($package)
            }
            
            Set-SelectedRow $dataModel.Rows[$i]
            [System.Windows.Forms.Application]::DoEvents()
       }
    }
    
    if ($action.TakesList) {
        $null = $action.Execute($checkedList)
        return
    }
    
    if (($tasksOkay -eq 0) -and ($tasksFailed -eq 0)) {
        Write-PipLog 'Nothing is selected.'
    } else {
        Write-PipLog ''
        Write-PipLog '----'
        Write-PipLog "All tasks finished, $tasksOkay ok, $tasksFailed failed."
        Write-PipLog 'Select a row to highlight the relevant log piece'
        Write-PipLog 'Double click or [Ctrl+Enter] a table row to open PyPi, Anaconda.com or github.com in browser'
        Write-PipLog '----'
        Write-PipLog ''
    }
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

$Global:FuncCalculateLevenshteinDistance = {
    <#
        .SYNOPSIS
        Returns Levenshtein distance of two strings

        .LINK
        https://en.wikipedia.org/wiki/Damerau%E2%80%93Levenshtein_distance
    
        .OUTPUTS
        System.Int32
    #>

    param([string] $word1, [string] $word2)
    
    [int] $len1 = $word1.Length
    [int] $len2 = $word2.Length
    
    $v = [array]::CreateInstance([int], $len1 + 1, $len2 + 1)
    
    [int] $i = 0
    [int] $j = 0

    for ( ; $i -le $len1; $i++) {
        $v[$i, 0] = $i
    }
    
    for ( ; $j -le $len2; $j++) {
        $v[0, $j] = $j
    }
    
    [int] $im1 = 0
    for ($i = 1; $i -le $len1; $i++, $im1++) {
        [int] $rowMin = $i
        [int] $jm1 = 0
        for ($j = 1; $j -le $len2; $j++, $jm1++) {           
            [char] $c1m1 = $word1[$im1]
            [char] $c1m2 = $word1[$i - 2]
            [char] $c2m1 = $word2[$jm1]
            [char] $c2m2 = $word2[$j - 2]

            if ($c1m1 -ceq $c2m1) {
                [int] $cost = 0
            } else {
                [int] $cost = 1
            }

            [int] $v1 = $v[($im1), $j] + 1  # deletion
            [int] $v2 = $v[$i, ($jm1)] + 1  # insertion
            [int] $v3 = $v[($im1), ($jm1)] + $cost  # subtraction
            
            [int] $v_ij = [Math]::Min([Math]::Min($v1, $v2), $v3)             
            
            if (($i -gt 1) -and ($j -gt 1) -and (
                ($c1m1 -ceq $c2m2) -and ($c1m2 -ceq $c2m1)
            )) {
                [int] $v4 = $v_ij
                [int] $v5 = $v[($i - 2), ($j - 2)] + $cost
                $v_ij = [Math]::Min($v4, $v5)  # transposition
            }
            
            $rowMin = [Math]::Min($rowMin, $v_ij)
            $v[$i, $j] = $v_ij
        }
        if ($rowMin -gt 2) {
            return 3
        }
    }

    return $v[$len1, $len2]
}

Function global:Get-TypoErrorCandidates([string] $text, [int] $threshold = 2) {
    $candidates = New-Object System.Collections.ArrayList
    $text = $text.ToLower()
    $textLength = $text.Length
    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()
    # $dbg_n = 0
    foreach ($n in ($textLength - $threshold)..($textLength + $threshold)) {
        $items = $Global:TypoErrorTable[$n]
        foreach ($item in $items) {
            # $dbg_n++
            # if ($dbg_n -gt 1000) {
            #     break
            # }
            # if ([Math]::Abs($textLength - $item.Length) -gt $threshold) {
            #     continue
            # }
            [int] $distance = & $Global:FuncCalculateLevenshteinDistance $text $item.ToLower()
            if ($distance -le $threshold) {
                $null = $candidates.Add(@{Text=$item; Distance=$distance})
            }
        }
    }
    $stopWatch.Stop()
    Write-PipLog "Search finished in $($stopWatch.Elapsed.TotalSeconds.ToString(".00")) sec."
    
    return ,($candidates | Sort-Object -Property Distance -Descending | ForEach-Object { $_.Text })
}

Function Get-PipDistributionInfo {
    $python_code = @'
import pip
import json
pkgs = pip.get_installed_distributions(local_only=False, user_only=False, skip=tuple())
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
                       $isExtra = $false) {
    # $output = & pip.exe show $name 2>&1
    # 
    # if ([string]::IsNullOrEmpty($output)) {
    #     return
    # }
    # 
    # $reqs = $output[$output.Count - 1]
    # 
    # if ([string]::IsNullOrEmpty($reqs) -or -not $reqs.StartsWith('Requires:')) {
    #     return
    # }
    # 
    # $n = "Requires:".Length
    # $children = $reqs.Substring($n).Split(', ', [System.StringSplitOptions]::RemoveEmptyEntries)
    
    $children = $distributionInfo."$name"."deps".PSObject.Properties.Name  # dep name list from JSON from pip                 
    if ($children -eq $null) {
        $children = @()
    } else {
        $children = @($children)
    }
    
    $extras = @($distributionInfo."$name"."extras")     
    if (($extras -ne $null) -and ($extras.Length -gt 0)) {
        $children = @($children; $extras)
    }
    
    if ($children.Length -gt 1) {
        $children = @($children) | Sort-Object
    }
    
    # Write-PipLog "$name children: '$children' $($children.Length) sibl=$hasSibling"
    $hasChildren = $children.Length -gt 0
    
    $prefix = if ($indent -gt 0) {
        if ($hasChildren) {
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
    
    $suffix = if ($isExtra) { ' (*)' } else { '' }

    "${prefix}${name}${suffix}"  # Add a line to the Return Stack

    if ($hasSibling) {
           [void]$dangling.Push($indent * 4)
    } else {
        [void]$dangling.Push(-1)
    }

    $i = 1
    foreach ($child in $children) {
        $childHasSiblings = ($i -lt $children.Length) -and ($children.Length -gt 1)
        Get-AsciiTree $child $distributionInfo ($indent + 1) $childHasSiblings $dangling ($child -in $extras)
        $i++
    }

    [void]$dangling.Pop()
}

Function global:Get-DependencyAsciiGraph($name) {
    $distributionInfo = Get-PipDistributionInfo
    $asciiTree = Get-AsciiTree $name.ToLower() $distributionInfo
    return $asciiTree
}

Function Show-ConsoleWindow([bool] $show) {
    Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
    
    $consolePtr = [Console.Window]::GetConsoleWindow()
    $value = if ($show) { 2 } else { 0 }
    [Console.Window]::ShowWindow($consolePtr, $value)
}

Function Save-PipsSettings {
    $settingsPath = "$($env:LOCALAPPDATA)\pips"
    $null = New-Item -Force -ItemType Directory -Path $settingsPath
    $userInterpreterRecords = $interpreters | Where-Object { $_.User }
    $settings = @{
        "envs"=@($userInterpreterRecords);
    }
    try {
        $settings | ConvertTo-Json | Out-File "$settingsPath\settings.json"
    } catch {
    }
}

Function Load-PipsSettings {
    $settingsFile = "$($env:LOCALAPPDATA)\pips\settings.json"
    if (Exists-File $settingsFile) {
        try {
            $Global:settings = Get-Content $settingsFile | ConvertFrom-Json
        } catch {
        }
    }
}

Function Start-Main([switch] $HideConsole, [switch] $Debug) {
    $env:PYTHONIOENCODING="utf-8"
    $env:LC_CTYPE="utf-8"
    
    if (-not $Debug) {
       Set-StrictMode -Off
       Set-PSDebug -Off
    } else {
       Set-StrictMode -Version latest
       Set-PSDebug -Strict         
    }
    
    if ($HideConsole) {
        Show-ConsoleWindow $false
    }
    
    Load-PipsSettings
    
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form = Generate-Form

    $form.Add_Closing({
        Save-PipsSettings
        [System.Windows.Forms.Application]::Exit()
        if (-not [Environment]::UserInteractive) {
            Stop-Process $pid
        }
    })

    $form.Show()
    $form.Activate()

    $appContext = New-Object System.Windows.Forms.ApplicationContext 
    try {
        [void][System.Windows.Forms.Application]::Run($appContext)
    } catch {
        Show-ConsoleWindow $true
    }
}
