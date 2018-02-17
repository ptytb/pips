[Void][Reflection.Assembly]::LoadWithPartialName("System")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Drawing.Size")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Drawing.Point")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.FontStyle")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Text")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Text.RegularExpressions")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Collections")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Collections.ArrayList")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Collections.Hashtable")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Collections.Generic")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Collections.Generic.HashSet")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Web")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Web.HttpUtility")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Net")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Net.WebClient")


Function Get-Bin($command) {
    (where.exe $command) | Select-Object -Index 0
}

Function Get-PythonPath() {
    $Script:interpretersComboBox.SelectedItem.Path
}

Function Get-PythonArchitecture() {
    $Script:interpretersComboBox.SelectedItem.Arch
}

Function Get-PythonExe() {
    $Script:interpretersComboBox.SelectedItem.PythonExe
}

Function Get-PipExe() {
    $Script:interpretersComboBox.SelectedItem.PipExe
}

Function Get-CondaExe() {
    $Script:interpretersComboBox.SelectedItem.CondaExe
}

Function Exists-File($path) {
    return [System.IO.File]::Exists($path)
}

Function Exists-Directory($path) {
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
    $button.Add_Click($handler)
    Add-TopWidget $button
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
    Add-Button "Execute:" { Execute-PipAction }
}

Function global:Get-PyDoc($request) {
    $output = & (Get-PythonExe) -m pydoc $request
    return $output
}

Function Get-PythonBuiltinPackages() {
    $builtinLibs = New-Object System.Collections.ArrayList
    $path = Get-PythonPath
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
        $trackDuplicates.Add("$packageName") | Out-Null        
        $builtinLibs.Add(@{Package=$packageName; Type='builtin'}) | Out-Null
    }

    $getBuiltinsScript = "import sys; print(','.join(sys.builtin_module_names))"
    $sys_builtin_module_names = & (Get-PythonExe) -c $getBuiltinsScript
    $modules = $sys_builtin_module_names.Split(',')
    foreach ($builtinModule in $modules) {
        if ($trackDuplicates.Contains("$builtinModule")) {
            continue
        }
        $builtinLibs.Add(@{Package=$builtinModule; Type='builtin'}) | Out-Null
    }

    return ,$builtinLibs
}

Function Get-CondaPackages() {
    $condaPackages = New-Object System.Collections.ArrayList
    $conda_exe = Get-CondaExe

    if ($conda_exe) {
        $arguments =New-Object System.Collections.ArrayList
        $arguments.Add('list') | Out-Null
        $arguments.Add('--json') | Out-Null
        $arguments.Add('--no-pip') | Out-Null
        $arguments.Add('--show-channel-urls') | Out-Null

        # This one sounds nice but could give versions older than installed
        # conda update --dry-run --json --all
        
        # This one sounds nice but could give versions older than installed
        # conda search --outdated

        $items = & $conda_exe $arguments | ConvertFrom-Json 

        foreach ($item in $items) {
            $condaPackages.Add(@{Type='conda'; Package=$item.name; Version=$item.version}) | Out-Null
        }
    }

    return ,$condaPackages
}

$actionCommands = @{
    pip=@{
        info          = { return (& (Get-PipExe) show              $args 2>&1) };
        documentation = { (Show-DocView $pkg).Show() | Out-Null; return ''     };
        files         = { return (& (Get-PipExe) show    --files   $args 2>&1) };
        update        = { return (& (Get-PipExe) install -U        $args 2>&1) };
        install       = { return (& (Get-PipExe) install           $args 2>&1) };
        install_dry   = { return 'Not supported on pip'                        };
        install_nodep = { return (& (Get-PipExe) install --no-deps $args 2>&1) };
        download      = { return (& (Get-PipExe) download          $args 2>&1) };
        uninstall     = { return (& (Get-PipExe) uninstall --yes   $args 2>&1) };
    };
    conda=@{
        info          = { return (& (Get-CondaExe) list      -v --json                     $args 2>&1) };        
        documentation = { return ''                                                                    };
        files         = { return ''                                                                    };
        update        = { return (& (Get-CondaExe) update  --yes                           $args 2>&1) };
        install       = { return (& (Get-CondaExe) install --yes --no-shortcuts            $args 2>&1) };
        install_dry   = { return (& (Get-CondaExe) install --dry-run                       $args 2>&1) };
        install_nodep = { return (& (Get-CondaExe) install --yes --no-shortcuts --no-deps  $args 2>&1) };
        download      = { return ''                                                                    };
        uninstall     = { return (& (Get-CondaExe) uninstall --yes                         $args 2>&1) };
    };
}
$actionCommands.wheel   = $actionCommands.pip
$actionCommands.sdist   = $actionCommands.pip
$actionCommands.builtin = $actionCommands.pip

Function Add-ComboBoxActions {
    Function Make-PipActionItem($name, $code, $validator) {
        $action = New-Object psobject -Property @{Name=$name}
        $action | Add-Member ScriptMethod ToString { $this.Name } -Force
		$action | Add-Member ScriptMethod Execute  $code
		$action | Add-Member ScriptMethod Validate $validator
        return $action
    }

    $actionsModel = New-Object System.Collections.ArrayList
    $Add = { param($a) $actionsModel.Add($a) | Out-Null }

    & $Add (Make-PipActionItem 'Show Info' `
		{ param($pkg,$type); $actionCommands[$type].info.Invoke($pkg) } `
        { param($pkg,$out); $out -match $pkg } )
	
    & $Add (Make-PipActionItem 'Documentation' `
		{ param($pkg,$type); $actionCommands[$type].documentation.Invoke($pkg) } `
        { param($pkg,$out); $out -match '.*' } )

    & $Add (Make-PipActionItem 'List Files' `
		{ param($pkg,$type); $actionCommands[$type].files.Invoke($pkg) } `
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

    & $Add (Make-PipActionItem 'Install' `
		{ param($pkg,$type); $actionCommands[$type].install.Invoke($pkg) } `
        { param($pkg,$out); $out -match ('Successfully installed |Installing collected packages:\s*(\s*\S*,\s*)*' + $pkg) } )

    & $Add (Make-PipActionItem 'Download' `
		{ param($pkg,$type); $actionCommands[$type].download.Invoke($pkg) } `
        { param($pkg,$out); $out -match 'Successfully downloaded ' } )

    & $Add (Make-PipActionItem 'Uninstall' `
		{ param($pkg,$type); $actionCommands[$type].uninstall.Invoke($pkg) } `
        { param($pkg,$out); $out -match ('Successfully uninstalled ' + $pkg) } )

    $Script:actionsModel = $actionsModel

    $actionListComboBox = New-Object System.Windows.Forms.ComboBox
    $actionListComboBox.DataSource = $actionsModel
    $actionListComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $Script:actionListComboBox = $actionListComboBox
    Add-TopWidget($actionListComboBox)    
}


$trackDuplicates = New-Object System.Collections.Generic.HashSet[String]

Function Get-InterpreterRecord($path, $items) {
    if ($trackDuplicates.Contains($path)) {
        continue
    }

    $python = Get-ExistingFilePathOrNull "${path}\python.exe"
    if (! $python) {
        $python = Get-ExistingFilePathOrNull "${path}\Scripts\python.exe"
    }    
    if (! $python) {
        return
	}	
	$versionString = & $python --version 2>&1
	$version = [regex]::Match($versionString, '\s+(\d+\.\d+)').Groups[1]

	$action = New-Object psobject -Property @{
		Path		    = $path;
		Version		    = $version;
		Arch		    = Test-is64Bit $python;
		PythonExe	    = $python;
		PipExe		    = Get-ExistingFilePathOrNull "${path}\Scripts\pip.exe";
		CondaExe	    = Get-ExistingFilePathOrNull "${path}\Scripts\conda.exe";
		VirtualenvExe   = Get-ExistingFilePathOrNull "${path}\Scripts\virtualenv.exe";
		PipenvExe	    = Get-ExistingFilePathOrNull "${path}\Scripts\pipenv.exe";
		RequirementsTxt = Get-ExistingFilePathOrNull "${path}\requirements.txt";
		Pipfile  	    = Get-ExistingFilePathOrNull "${path}\Pipfile";
	}
	$action | Add-Member ScriptMethod ToString {
		"{2} [{0}] {1}" -f $this.Arch.FileType, $this.PythonExe, $this.Version
	} -Force

    $items.Add($action) | Out-Null
    $trackDuplicates.Add($path) | Out-Null
}

Function Find-Interpreters {
    $items = New-Object System.Collections.ArrayList

    $list = @(where.exe 'python'; (dir $env:SystemDrive\Python*) | foreach{ "$_\python.exe" })
    foreach ($path in $list) {
        Get-InterpreterRecord (Split-Path -Parent $path) $items
    }

    foreach ($d in dir "$env:LOCALAPPDATA\Programs\Python") {
        if ($d -is [System.IO.DirectoryInfo]) {
            Get-InterpreterRecord (${d}.FullName) $items
        }
    }

    return ,$items  # keep comma to prevent conversion to an @() array
}


Function Add-ComboBoxInterpreters {
    $interpreters = Find-Interpreters
    $Script:interpreters = $interpreters
    $interpretersComboBox = New-Object System.Windows.Forms.ComboBox
    $interpretersComboBox.DataSource = $interpreters
    $interpretersComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $Script:interpretersComboBox = $interpretersComboBox
    Add-TopWidget $interpretersComboBox 4
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
        $paths = @('.\env\Scripts\', '.\Scripts\')
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
    return $index
}

Function Generate-FormInstall {
    Function Prepare-PackageAutoCompletion {
        $Script:packageIndex = Load-KnownPackageIndex
        $Script:autoCompleteIndex = New-Object System.Windows.Forms.AutoCompleteStringCollection
        foreach ($item in $packageIndex.Keys) {
            $autoCompleteIndex.Add($item)
        }
    }

    if ($Script:packageIndex -eq $null) {
        Prepare-PackageAutoCompletion
    }

    $form = New-Object System.Windows.Forms.Form
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = 'Type package name and hit Enter to enqueue'
    $label.Location = New-Object Drawing.Point 7,7
    $label.Size = New-Object Drawing.Point 270,24
    $form.Controls.Add($label)

    $cb = New-Object System.Windows.Forms.TextBox
    $cb.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
    $cb.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource
    $cb.AutoCompleteCustomSource = $Script:autoCompleteIndex
    $cb.add_KeyDown({
        if ($_.KeyCode -eq 'Escape') {
            if ($cb.Text.Length -gt 0) {
                $cb.Text = [string]::Empty
            } else {
                $form.Close()
            }
        }
        if ($_.KeyCode -eq 'Enter') {
            if ($dataModel.Rows.Count -eq 0) {
                Init-PackageSearchColumns $dataModel
            }

            $package = $cb.Text
            $record = $packageIndex[$package]
            if (-not $record) {
                return
            }
            $row = $dataModel.NewRow()
            $row.Select = $true
            $row.Package = $package
            if ($dataModel.Columns.Contains('Version')) {
                $row.Version = $record.Version
            } else {
                $row.Latest  = $record.Version
            }
            if ($dataModel.Columns.Contains('Description')) {
                $row.Description = $record.Description
            }
            $row.Type = 'pip'
            $row.Status = ''
            $dataModel.Rows.InsertAt($row, 0)

            $cb.Text = [string]::Empty
        }
    })
    $cb.Location = New-Object Drawing.Point 7,40
    $cb.Size = New-Object Drawing.Point 240,32
    $form.Controls.Add($cb)

    $form.Size = New-Object Drawing.Point 280,130
    $form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Hide
    $form.FormBorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $form.Text = 'Install packages'
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false
    $form.Icon = $Script:form.Icon
    $form.ShowDialog()
}

Function Generate-FormSearch {
    $message = "Enter keywords to search PyPi and Conda`n`n* = list all packages`n`nChecked items will be kept in the search list"
    $title = "pip search ... & conda search ..."
    $default = "*"

    $input = $(
        Add-Type -AssemblyName Microsoft.VisualBasic
        [Microsoft.VisualBasic.Interaction]::InputBox($message, $title, $default)
    )
    
    if (! $input) {
        return
    }

    Write-PipLog ("Searching for " + $input)
    Write-PipLog 'Double click a table row to open PyPi in browser (online)'
    
    Write-PipLog
    $stats = Get-SearchResults $input
    Write-PipLog "Found $($stats.Total) packages: $($stats.PipCount) pip, $($stats.CondaCount) conda. Total $($dataModel.Rows.Count) packages in list."
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

Function Highlight-PythonBuiltinPackages {
    if (! $outdatedOnly) {
        $dataGridView.BeginInit()
        foreach ($row in $dataGridView.Rows) {
            if ($row.DataBoundItem.Row.Type -eq 'builtin') {
                $row.DefaultCellStyle.BackColor = [Drawing.Color]::LightGreen
            }
        }
        $dataGridView.EndInit()
    }    
}

Function global:Open-LinkInBrowser($url) {
    if ($url -match '^https://') {
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
        Write-Host 'run timer'
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
                    Write-Host 'stop timer'
                }
                Write-Host "*** Cleaned up $n sem=$Global:jobSemaphore"
            }           
        }
}

$Global:PyPiPackageJsonCache = New-Object System.Collections.Hashtable

Function global:Format-PythonPackageToolTip($info) {
     $tt = "Summary: $($info.info.summary)`nAuthor: $($info.info.author)`nRelease: $($info.info.version)`n"
     $lr = ($info.releases."$($info.info.version)")[0]
     $tr = "Release Uploaded $($lr.upload_time)"
     $tags = "`n`n$($info.info.classifiers -join "`n")"
     return "${tt}${tr}${tags}"
}

Function global:Update-PythonPackageDetails {
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

    $Global:dataModel = $dataModel
    $Global:dataGridView = $dataGridView

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
        $Global:PyPiPackageJsonCache.Add($message.Params.PackageName, $message.Result)
        
        $viewRow = $Global:dataGridView.Rows[$message.Params.RowIndex]
        $viewRow.DefaultCellStyle.BackColor = [Drawing.Color]::Empty
        $row = $viewRow.DataBoundItem.Row        
        $Global:dataModel.Columns['Status'].ReadOnly = $false
        $row.Status = 'OK'
        $Global:dataModel.Columns['Status'].ReadOnly = $true        
    }) @{PackageName=$packageName; RowIndex=$_.RowIndex}
}

Function Generate-Form {
    $form = New-Object Windows.Forms.Form
    $form.Text = "pip package browser"
    $form.Size = New-Object Drawing.Point 1050, 840
	$form.topmost = $false
	$form.KeyPreview = $true
    $iconPath = Get-Bin 'pip'
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
	$Script:form = $form
	
	$form.add_KeyDown({
		$gridViewActive = $form.ActiveControl -is [System.Windows.Forms.DataGridView]
		if ($dataGridView.Focused -and $dataGridView.RowCount -gt 0) {			
			if ($_.KeyCode -eq 'Home') {
				Set-SelectedNRow 0
				$_.Handled = $true
			}
			if ($_.KeyCode -eq 'End') {
				Set-SelectedNRow ($dataGridView.RowCount - 1)
				$_.Handled = $true
			}
		}
	})

    Add-Buttons | Out-Null
	
	Add-ComboBoxActions	
	$form.add_KeyDown({
		$comboActive = $form.ActiveControl -is [System.Windows.Forms.ComboBox]
		if (($_.KeyCode -eq 'C') -and ($_.Control) -and $comboActive) {
			$python_exe = Get-PythonExe
			Set-Clipboard $python_exe
			Write-PipLog "Copied to clipboard: $python_exe"
		}
	})
    
    $Script:isolatedCheckBox = Add-CheckBox 'isolated' { Toggle-VirtualEnv $Script:isolatedCheckBox.Checked }
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.SetToolTip($isolatedCheckBox, 'Pip ignores environment variables and user configuration.')

    $null = Add-Button "Search..." { Generate-FormSearch }

    $null = Add-Button "Install..." { Generate-FormInstall }

    NewLine-TopLayout | Out-Null

    Add-Label "Filter results:" | Out-Null
    
    $Script:inputFilter = Add-Input {  # TextChanged Handler here
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

        Highlight-PythonBuiltinPackages
    }
    $inputFilter.Add_KeyDown({
            if ($_.KeyCode -eq 'Escape') {
                $inputFilter.Text = [String]::Empty
            }
        })
    $toolTipFilter = New-Object System.Windows.Forms.ToolTip
    $toolTipFilter.SetToolTip($inputFilter, "Esc to clear")

    #dd-HorizontalSpacer | Out-Null

    $searchMethodComboBox = New-Object System.Windows.Forms.ComboBox
    $searchMethods = New-Object System.Collections.ArrayList
    $searchMethods.Add('Whole Phrase') | Out-Null
    $searchMethods.Add('Any Word') | Out-Null
    $searchMethods.Add('All Words') | Out-Null
    $searchMethods.Add('Exact Match') | Out-Null
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

    Add-ComboBoxInterpreters | Out-Null
    $null = Add-Button "Add venv path..." {
        $selectFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $selectFolderDialog.ShowDialog()
        $path = $selectFolderDialog.SelectedPath
        $path = Get-ExistingPathOrNull $path
        if ($path) {
            $oldCount = $interpreters.Count
            Get-InterpreterRecord $path $interpreters
            if ($interpreters.Count -gt $oldCount) {
                $interpretersComboBox.DataSource = $null
                $interpretersComboBox.DataSource = $interpreters
                Write-PipLog "Added virtual environment location: $path"
            } else {
                Write-PipLog "No python found in $path"
            }
        }
    }
    
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $Script:dataGridView = $dataGridView
    $dataGridView.Location = New-Object Drawing.Point 7,($Script:lastWidgetTop + $Script:widgetLineHeight)
    $dataGridView.Size = New-Object Drawing.Point 800,450
    $dataGridView.ShowCellToolTips = $true
    $dataGridView.Add_Sorted({ Highlight-PythonBuiltinPackages })
    $dataGridView.Add_CellMouseEnter({
            if (($_.RowIndex -gt -1)) {
                Update-PythonPackageDetails
            }
        })
    Init-PackageGridViewProperties
    
    $dataModel = New-Object System.Data.DataTable
    $dataGridView.DataSource = $dataModel    
    $Script:dataModel = $dataModel
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

    $FuncShowPackageInBrowser = {
        $view_row = $dataGridView.CurrentRow
        if ($view_row) {
            $row = $view_row.DataBoundItem.Row
            $packageName = $row.Package
            $urlName = [System.Web.HttpUtility]::UrlEncode($packageName)
            if ($row.Type -eq 'conda') {
                $url = $anaconda_url
            } else {
                $url = $pypi_url
            }
            Open-LinkInBrowser "${url}${urlName}"
        }
    }

    $dataGridView.Add_CellMouseDoubleClick({
            if (($_.RowIndex -gt -1) -and ($_.ColumnIndex -gt 0)) {
                & $FuncShowPackageInBrowser
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
        Write-PipLog 'Hold Shift and hover the rows to fetch detailed info, then hover the Package column'
        $form.BringToFront()
        })
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

$DrawingColor = New-Object Drawing.Color  # Workaround for PowerShell, which parses script classes before loading types

class DocView {
    
    [System.Windows.Forms.Form]        $formDoc
    [System.Windows.Forms.RichTextBox] $docView
    [int]           $modifiedTextLengthDelta  # We've found all matches with immutual Text, but will be changing actual document
	[String]        $packageName

    DocView($content, $packageName) {
		$self = $this
		$this.packageName = $packageName
        $this.modifiedTextLengthDelta = 0
        
        $this.formDoc = New-Object Windows.Forms.Form
        $this.formDoc.Text = "PyDoc for [$packageName] *** Click/Enter goto $packageName.WORD | Esc back | Ctrl+Wheel font | Space scroll"
        $this.formDoc.Size = New-Object Drawing.Point 830, 840
        $this.formDoc.Topmost = $false
        $this.formDoc.KeyPreview = $true
        $this.formDoc.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent

        $this.formDoc.Icon = $Global:form.Icon

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
	                if ($self.docView.Text[$charIndex] -notmatch '\w') {
	                    $begin = $charIndex + 1
	                    break
	                }
	            }
	        
	            for ($charIndex = $clickedIndex; $charIndex -lt $self.docView.Text.Length; $charIndex++) {
	                if ($self.docView.Text[$charIndex] -notmatch '\w') {
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

        $this.Highlight_Links()
        $this.Highlight_PyDocSyntax()
	    $this.docView.Select(0, 0)
        $this.docView.ScrollToCaret()	    
    }
	
	[void] Show() {
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

    [void] Highlight_Text($pattern, $foreground, $useBold) {
        $this.Alter_MatchingFragments($pattern, {
		    $this.docView.SelectionColor = $foreground
            if ($useBold) {
                $this.docView.SelectionFont = $fontBold
            }
	    })
    }

    [void] Highlight_PyDocSyntax() {
        $this.Highlight_Text($Script:pyRegexStrQuote,  ($Script:DrawingColor::DarkGreen),   $false)
        $this.Highlight_Text($Script:pyRegexNumber,    ($Script:DrawingColor::DarkMagenta), $false)
        $this.Highlight_Text($Script:pyRegexSection,   ($Script:DrawingColor::DarkCyan),    $true)
        $this.Highlight_Text($Script:pyRegexKeyword,   ($Script:DrawingColor::DarkRed),     $true)
        $this.Highlight_Text($Script:pyRegexSpecial,   ($Script:DrawingColor::DarkOrange),  $true)
        $this.Highlight_Text($Script:pyRegexSpecialU,  ($Script:DrawingColor::DarkOrange),  $true)
    }
    
    [void] Highlight_Links() {
	    # PEP Links
	    $this.Alter_MatchingFragments($Script:pyRegexPEP, {
            param($index, $length, $match)
            $pep_url = "${peps_url}$(($match -replace ' ', '-').ToLower())"
            $this.modifiedTextLengthDelta -= $this.docView.TextLength
            $this.docView.SelectedText = "$match [$pep_url]"
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

Function global:Show-DocView($packageName) {
	$content = (Get-PyDoc $packageName)
    $viewer  = New-Object DocView -ArgumentList @((Tidy-Output $content), $packageName)
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
    $pip_exe = Get-PipExe
    if (!$pip_exe) {
        Write-PipLog 'pip is not found!'
        return 0
    }

    $args = New-Object System.Collections.ArrayList
    $args.Add('search') | Out-Null
    $args.Add("$request") | Out-Null
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
    $conda_exe = Get-CondaExe
    if (! $conda_exe) {
        Write-PipLog 'conda is not found!'
        return 0
    }
    $arch = Get-PythonArchitecture

    # --info should give better details but not supported on every conda
    $items = & $conda_exe search --json $request | ConvertFrom-Json

    $count = 0
    $items.PSObject.Properties | forEach-Object {
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

    $dataModel.EndLoadData()
    $dataGridView.EndInit()

    return @{PipCount=$pipCount; CondaCount=$condaCount; Total=($pipCount + $condaCount)}
}

 Function Get-PythonPackages($outdatedOnly = $true) {
    Write-PipLog
    Write-PipLog 'Updating package list... '
    
    $python_exe = Get-PythonExe
    $pip_exe = Get-PipExe
    $conda_exe = Get-CondaExe
    
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
        $args.Add('list') | Out-Null

        if ($outdatedOnly) {
            $args.Add('--outdated') | Out-Null
        }

        $args.Add('--format=columns') | Out-Null
        if ($Script:isolatedCheckBox.Checked) {
            $args.Add('--isolated') | Out-Null
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
    }
    if ($conda_exe) {
        $condaPackages = Get-CondaPackages
        Add-PackagesToTable $condaPackages 'conda'
    }
    $dataModel.EndLoadData()

    $Script:outdatedOnly = $outdatedOnly
    Highlight-PythonBuiltinPackages

    Write-PipLog 'Package list updated.'
    Write-PipLog 'Double click a table row to open PyPi in browser (online)'
    
    $count = $dataModel.Rows.Count
    $pipCount = $pipPackages.Count
    $builtinCount = $builtinPackages.Count
    $condaCount = $condaPackages.Count
    Write-PipLog "Total $count packages: $builtinCount builtin, $pipCount pip, $condaCount conda"
}

Function Select-VisiblePipPackages($value) {
    $dataModel.BeginLoadData()
    for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
        if ($dataGridView.Rows[$i].DataBoundItem.Row.Type -eq 'builtin' ) {
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

Function Tidy-Output($text) {
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

function Set-SelectedNRow($n) {
	$dataGridView.ClearSelection()
	$dataGridView.FirstDisplayedScrollingRowIndex = $n
	if ($n -lt $dataGridView.RowCount) {
        $dataGridView.Rows[$n].Selected = $true
    }
}

Function Set-SelectedRow($selectedRow) {
    $Script:dataGridView.ClearSelection()
    foreach ($vRow in $Script:dataGridView.Rows) {
        if ($vRow.DataBoundItem.Row -eq $selectedRow) {
            $vRow.Selected = $true
            $dataGridView.FirstDisplayedScrollingRowIndex = $vRow.Index
            break
        }
	}
}

Function Check-PipDependencies {
    Write-PipLog 'Checking dependencies...'

    $pip_exe = Get-PipExe
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

Function Execute-PipAction($action) {
    for ($i = 0; $i -lt $dataModel.Rows.Count; $i++) {
       if ($dataModel.Rows[$i].Select -eq $true) {
            $package =  $dataModel.Rows[$i]
            $action = $Script:actionsModel[$actionListComboBox.SelectedIndex]
            
            Write-PipLog ""
            Write-PipLog $action.Name ' ' $package.Package

            $result = $action.Execute($package.Package, $package.Type)

            $logFrom = $Script:logView.TextLength 
            Write-PipLog (Tidy-Output $result)
            $logTo = $Script:logView.TextLength - $logFrom
            $dataModel.Rows[$i] | Add-Member -Force -MemberType NoteProperty -Name LogFrom -Value $logFrom
            $dataModel.Rows[$i] | Add-Member -Force -MemberType NoteProperty -Name LogTo -Value $logTo

            $dataModel.Columns['Status'].ReadOnly = $false
            if ($action.Validate($package.Package, $result)) {
                 $dataModel.Rows[$i].Status = "OK"
                 Set-Unchecked $i
            } else {
                 $dataModel.Rows[$i].Status = "Failed"
            }
            $dataModel.Columns['Status'].ReadOnly = $true
            Set-SelectedRow $dataModel.Rows[$i]
       }
    }

    Write-PipLog ''
    Write-PipLog '----'
    Write-PipLog 'All tasks finished.'
    Write-PipLog 'Select a row to highlight the relevant log piece'
    Write-PipLog 'Double click a table row to open PyPi or Anaconda.com in browser (online)'
    Write-PipLog '----'
    Write-PipLog ''
}

# by https://superuser.com/users/243093/megamorf
function Test-is64Bit {
    param($FilePath)

    [int32]$MACHINE_OFFSET = 4
    [int32]$PE_POINTER_OFFSET = 60

    [byte[]]$data = New-Object -TypeName System.Byte[] -ArgumentList 4096
    $stream = New-Object -TypeName System.IO.FileStream -ArgumentList ($FilePath, 'Open', 'Read')
    $stream.Read($data, 0, 4096) | Out-Null

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
function Test-KeyPress
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


Function Start-Main() {
    $env:PYTHONIOENCODING="utf-8"
    $env:LC_CTYPE="utf-8"
    
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form = Generate-Form
    $form.ShowDialog()
}
