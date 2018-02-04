[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Web")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Web.HttpUtility")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Text")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Text.RegularExpressions")
[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.FontStyle")


Function Get-Bin($command) {
    (where.exe $command) | Select-Object -Index 0
}

Function Get-PythonPath() {
    $Script:interpretersComboBox.SelectedItem.Path
}

Function Get-PythonExe() {
    $Script:interpretersComboBox.SelectedItem.PythonExe
}

Function Get-PipExe() {
    $Script:interpretersComboBox.SelectedItem.PipExe
}

Function Exists-File($path) {
    return [System.IO.File]::Exists($path)
}

Function Get-ExistingFilePathOrNull($path) {
    if (Exists-File $path) {
        return $path
    } else {
        return $null
    }
}

$pypi_path = 'https://pypi.python.org/pypi/'
$lastWidgetLeft = 5
$lastWidgetTop = 5
$widgetLineHeight = 23
$dataGridView = $null
$inputFilter = $null
$logView = $null
$actionsModel = $null
$virtualenvCheckBox = $null
$header = ("Select", "Package", "Installed", "Latest", "Type", "Status")
$csv_header = ("Package", "Installed", "Latest", "Type", "Status")
$search_columns = ("Select", "Package", "Version", "Description", "Status")
$formLoaded = $false
$outdatedOnly = $true


Function Write-PipLog() {
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
    $Script:lastWidgetLeft = $lastWidgetLeft + 100
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
}

Function Add-Label ($name) {
    $label = New-Object Windows.Forms.Label
    $label.Text = $name
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    Add-TopWidget $label
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

Function Get-PyDoc($request) {
    $output = & (Get-PythonExe) -m pydoc $request
    return $output
}

Function Get-PythonBuiltinPackages() {
    $builtinLibs = New-Object System.Collections.ArrayList
    $path = Get-PythonPath
    $libs = "${path}\Lib"
    $ignore = [regex] '^__'

    $trackDuplicates = New-Object System.Collections.Generic.HashSet[String]

    foreach ($item in dir $libs) {
        if ($item -cmatch $ignore) {
            continue
        }

        $trackDuplicates.Add("$item") | Out-Null
        
        $fullItem = "$libs\$item"
        
        if ($item -is [System.IO.DirectoryInfo]) {
            $packageName = "$item"
        } elseif ($item -is [System.IO.FileInfo]) {
            $packageName = "$item" -replace '.py$',''
        }
        
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

    return $builtinLibs
}

Function Add-ComboBoxActions {
    Function Make-PipActionItem($name, $code, $validation) {
        $action = New-Object psobject -Property @{Name=$name; Validation=$validation}
        $action | Add-Member ScriptMethod ToString { $this.Name } -Force
        $action | Add-Member ScriptMethod Execute $code  # $code takes $args array which will have only package name
        return $action
    }

    $actionsModel = New-Object System.Collections.ArrayList
    $Add = { param($a) $actionsModel.Add($a) | Out-Null }

    & $Add (Make-PipActionItem 'Show Info'      {return (& (Get-PipExe) show  $args 2>&1)} `
        '.*' )

    & $Add (Make-PipActionItem 'Documentation' {Show-DocView (Get-PyDoc $args) $args[0] | Out-Null; return ''} `
        '.*' )

    & $Add (Make-PipActionItem 'Update'    {return (& (Get-PipExe) install -U $args 2>&1)} `
        'Successfully installed |Installing collected packages:\s*(\s*\S*,\s*)*' )

    & $Add (Make-PipActionItem 'Install'    {return (& (Get-PipExe) install   $args) 2>&1} `
        'Successfully installed |Installing collected packages:\s*(\s*\S*,\s*)*' )

    & $Add (Make-PipActionItem 'Download'  {return (& (Get-PipExe) download   $args) 2>&1} `
        'Successfully downloaded ' )

    & $Add (Make-PipActionItem 'Uninstall' {return (& (Get-PipExe) uninstall  $args) 2>&1} `
        'Successfully uninstalled ' )

    $Script:actionsModel = $actionsModel

    $actionList = New-Object System.Windows.Forms.ComboBox
    $actionList.DataSource = $actionsModel
    $actionList.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $Script:actionList = $actionList
    Add-TopWidget($actionList)    
}

Function Find-Interpreters {
    $items = New-Object System.Collections.ArrayList
    $trackDuplicates = New-Object System.Collections.Generic.HashSet[String]

    Function Get-InterpreterRecord($path) {
        if ($trackDuplicates.Contains($path)) {
            continue
        }
        $trackDuplicates.Add($path) | Out-Null
        
        $python = Get-ExistingFilePathOrNull "${path}\python.exe"
        $pip = Get-ExistingFilePathOrNull "${path}\Scripts\pip.exe"
        $conda = Get-ExistingFilePathOrNull "${path}\Scripts\conda.exe"
        
        if ($python) {
            $arch = Test-is64Bit $python
        } else {
            $arch = $null
        }

        $action = New-Object psobject -Property @{Path=$path; Arch=$arch; PythonExe=$python; PipExe=$pip; CondaExe=$conda}
        $action | Add-Member ScriptMethod ToString { "[{0}] {1}" -f $this.Arch.FileType,$this.PythonExe } -Force

        $items.Add($action) | Out-Null
    }

    $list = (where.exe 'python')
    foreach ($path in $list) {
        Get-InterpreterRecord (Split-Path -Parent $path)
    }

    foreach ($d in dir "$env:LOCALAPPDATA\Programs\Python") {
        if ($d -is [System.IO.DirectoryInfo]) {
            Get-InterpreterRecord (${d}.FullName)
        }
    }

    return $items
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

Function Generate-FormInstall {
    $message = "Enter keywords to search PyPi`n`n* = list all packages`n`nChecked items will be kept in the search list"
    $title = "pip search ..."
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
    
    Get-PipSearchResults $input
    
    Write-PipPackageCounter
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

Function Generate-Form {
    $form = New-Object Windows.Forms.Form
    $form.Text = "pip package browser"
    $form.Size = New-Object Drawing.Point 1000, 840
    $form.topmost = $false
    $iconPath = Get-Bin 'pip'
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
    $Script:form = $form

    Add-Buttons
    Add-ComboBoxActions
    $Script:virtualenvCheckBox = Add-CheckBox 'virtualenv' { Toggle-VirtualEnv $Script:virtualenvCheckBox.Checked }
    Add-Button "Search..." { Generate-FormInstall }

    NewLine-TopLayout

    Add-Label "Filter results:"
    
    $Script:inputFilter = Add-Input {
        param($input)
        
        if ($Script:dataGridView.CurrentRow) {
            # Keep selection while filter is being changed
            $selectedRow = $Script:dataGridView.CurrentRow.DataBoundItem.Row
        }

        $isInstallMode = $dataModel.Columns.Contains('Description')
        $searchText = $input.Text
        if ($isInstallMode) {
            $query = "Package LIKE '{0}%' OR Package LIKE '%{0}%' OR Description LIKE '%{0}%'" -f $searchText
        } else {
            $query = "Package LIKE '{0}%' OR Package LIKE '%{0}%'" -f $searchText
        }
        
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

    Add-HorizontalSpacer
    Add-Label "Active Interpreter:"
    Add-ComboBoxInterpreters
    Add-Button "Add venv dir..." { Write-PipLog 'Not implemented yet...' }
    
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $Script:dataGridView = $dataGridView
    $dataGridView.Location = New-Object Drawing.Point 7,($Script:lastWidgetTop + $Script:widgetLineHeight)
    $dataGridView.Size = New-Object Drawing.Point 800,450
    $dataGridView.Add_Sorted({ Highlight-PythonBuiltinPackages })
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

    Function Highlight-LogFragment() {
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
    
    $dataGridView.Add_CellMouseClick({ Highlight-LogFragment })
    $dataGridView.Add_SelectionChanged({ Highlight-LogFragment })

    Function Show-PackageInBrowser() {
        $row = $dataGridView.CurrentRow
        if ($row) {
            $packageName = $row.DataBoundItem.Row['Package']
            $urlName = [System.Web.HttpUtility]::UrlEncode($packageName)
            Start-Process -FilePath "${pypi_path}${urlName}"
        }
    }

    $dataGridView.Add_CellMouseDoubleClick({ Show-PackageInBrowser })
    $form.Add_Load({ $Script:formLoaded = $true })
    
    Function Resize-Form() {
        $dataGridView.Width = $form.ClientSize.Width - 15
        $dataGridView.Height = $form.ClientSize.Height / 2
        $logView.Top = $dataGridView.Bottom + 15
        $logView.Width = $form.ClientSize.Width - 15
        $logView.Height = $form.ClientSize.Height - $dataGridView.Bottom - $lastWidgetTop
    }

    $form.Add_Closed({ if ($formDoc) { $formDoc.Close() } })

    Resize-Form
    $form.Add_Resize({ Resize-Form })
    $form.ShowDialog()
    $form.BringToFront()
}

Function Resize-FormDoc() {
    $docView.Width = $formDoc.ClientSize.Width - 15
    $docView.Height = $formDoc.ClientSize.Height - 15
}

Function Generate-FormDocView($title) {
    $formDoc = New-Object Windows.Forms.Form
    $formDoc.Text = $title
    $formDoc.Size = New-Object Drawing.Point 830, 840
    $formDoc.topmost = $false
    $iconPath = Get-Bin 'pip'
    $formDoc.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
    $Script:formDoc = $formDoc

    $docView = New-Object System.Windows.Forms.RichTextBox
    $docView.Location = New-Object Drawing.Point 7,7
    $docView.Size = New-Object Drawing.Point 800,810
    $docView.ReadOnly = $true
    $docView.Multiline = $true
    $docView.Font = New-Object System.Drawing.Font("Consolas", 11)
    $docView.WordWrap = $false
    $Script:docView = $docView
    $formDoc.Controls.Add($docView)

    Resize-FormDoc
    $formDoc.Add_Resize({ Resize-FormDoc })
}

Function Highlight-Output() {    
    Function Highlight-Text($pattern, $foreground = [Drawing.Color]::DarkCyan) {
        $regexOptions = [System.Text.RegularExpressions.RegexOptions]::ExplicitCapture 
                      + [System.Text.RegularExpressions.RegexOptions]::Compiled
        $matches = [regex]::Matches($docView.Text, "$pattern")

        $fontBold = New-Object System.Drawing.Font("Consolas",11,[System.Drawing.FontStyle]::Bold)

        foreach ($match in $matches.Groups) {
            if ($match.Name -eq 0) {
                continue
            }
            $docView.Select($match.Index, $match.Length)
            $docView.SelectionColor = $foreground
            $docView.SelectionFont = $fontBold
        }
    }
    
    $pydocSections = @('NAME', 'DESCRIPTION', 'PACKAGE CONTENTS', 'CLASSES', 'FUNCTIONS', 'DATA', 'VERSION', 'AUTHOR', 'FILE')
    $pythonKeywords = @('False', 'None', 'True', 'and', 'as', 'assert', 'break', 'class', 'continue', 'def', 'del', 'elif', 'else', 'except', 'finally', 'for', 'from', 'global', 'if', 'import', 'in', 'is', 'lambda', 'nonlocal', 'not', 'or', 'pass', 'raise', 'return', 'try', 'while', 'with', 'yield')
    $pythonSpecialMethods = @('self', '__all__', '__abs__', '__add__', '__and__', '__call__', '__class__', '__cmp__', '__coerce__', '__complex__', '__contains__', '__del__', '__delattr__', '__delete__', '__delitem__', '__delslice__', '__dict__', '__div__', '__divmod__', '__eq__', '__float__', '__floordiv__', '__ge__', '__get__', '__getattr__', '__getattribute__', '__getitem__', '__getslice__', '__gt__', '__hash__', '__hex__', '__iadd__', '__iand__', '__idiv__', '__ifloordiv__', '__ilshift__', '__imod__', '__imul__', '__index__', '__init__', '__instancecheck__', '__int__', '__invert__', '__ior__', '__ipow__', '__irshift__', '__isub__', '__iter__', '__itruediv__', '__ixor__', '__le__', '__len__', '__long__', '__lshift__', '__lt__', '__metaclass__', '__mod__', '__mro__', '__mul__', '__ne__', '__neg__', '__new__', '__nonzero__', '__oct__', '__or__', '__pos__', '__pow__', '__radd__', '__rand__', '__rcmp__', '__rdiv__', '__rdivmod__', '__repr__', '__reversed__', '__rfloordiv__', '__rlshift__', '__rmod__', '__rmul__', '__ror__', '__rpow__', '__rrshift__', '__rshift__', '__rsub__', '__rtruediv__', '__rxor__', '__set__', '__setattr__', '__setitem__', '__setslice__', '__slots__', '__str__', '__sub__', '__subclasscheck__', '__truediv__', '__unicode__', '__weakref__', '__xor__')

    foreach ($text in $pydocSections) {
        Highlight-Text "\W($text)\W"
    }

    foreach ($text in $pythonKeywords) {
        Highlight-Text "\W($text)\W" ([Drawing.Color]::DarkRed)
    }

    foreach ($text in $pythonSpecialMethods) {
        Highlight-Text "\W($text)\W" ([Drawing.Color]::DarkOrange)
    }
}

Function Show-DocView($text, $packageName) {
    Generate-FormDocView "PyDoc for $packageName"
    $Script:docView.Text = (Tidy-Output $text)
    Highlight-Output
    $docView.Select(0, 0)
    $docView.ScrollToCaret()
    $formDoc.ShowDialog()
    $formDoc.BringToFront()
}

Function Write-PipPackageCounter {
    $count = $dataModel.Rows.Count
    Write-PipLog "Now $count packages in the list."
}

Function Store-CheckedPipSearchResults() {
    $selected = New-Object System.Data.DataTable
    Init-PackageSearchColumns $selected
    foreach ($row in $dataModel) {
        if ($row.Select) {
            $selected.ImportRow($row)
        }
    }
    return $selected
}

Function Get-PipSearchResults($request) {
    $pip_exe = Get-PipExe
    if (!$pip_exe) {
        Write-PipLog 'pip is not found!'
        return
    }

    $args = New-Object System.Collections.ArrayList
    $args.Add('search') | Out-Null
    $args.Add("$request") | Out-Null
    $output = & $pip_exe $args

    $results = $dataModel
       
    $previousSelected = Store-CheckedPipSearchResults
    
    Clear-Rows
    Init-PackageSearchColumns $dataModel
    
    $dataGridView.BeginInit()
    $results.BeginLoadData()

    $r = [regex] '^(.*?)\s*\((.*?)\)\s+-\s+(.*?)$'

    foreach ($row in $previousSelected) {
        $results.ImportRow($row)
    }

    foreach ($line in $output) {
        $m = $r.Match($line)
        $row = $results.NewRow()
        $row['Select'] = $false
        $row['Package'] = $m.Groups[1].Value
        $row['Version'] = $m.Groups[2].Value
        $row['Description'] = $m.Groups[3].Value
        $results.Rows.Add($row)
    }    
    
    $results.EndLoadData()
    $dataGridView.EndInit()
}

 Function Get-PythonPackages($outdatedOnly = $true) {
    Write-PipLog
    Write-PipLog 'Updating package list... '
    
    $python_exe = Get-PythonExe
    $pip_exe = Get-PipExe
    
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
        if ($Script:virtualenvCheckBox.Checked) {
            $args.Add('--isolated') | Out-Null
        }

        $pip_list = & $pip_exe $args
        $packages = $pip_list | Select-Object -Skip 2 | % { $_ -replace '\s+', ' ' }  | ConvertFrom-Csv -Header $csv_header -Delimiter ' '
    }

    Function Add-PackagesToTable($packages) {        
        for ($n = 0; $n -lt $packages.Count; $n++) {
            $row = $dataModel.NewRow()        
            $row['Select'] = $false
            $row['Package'] = $packages[$n].Package
            $row['Installed'] = $packages[$n].Installed
            $row['Latest'] = $packages[$n].Latest
            $row['Type'] = $packages[$n].Type
            $dataModel.Rows.Add($row)
        }        
    }

    $dataModel.BeginLoadData()
    if ($pip_exe) {
        Add-PackagesToTable $packages
    }
    if (! $outdatedOnly) {
        $builtinPackages = Get-PythonBuiltinPackages
        Add-PackagesToTable $builtinPackages
    }
    $dataModel.EndLoadData()

    $Script:outdatedOnly = $outdatedOnly
    Highlight-PythonBuiltinPackages

    Write-PipLog 'Package list updated.'
    Write-PipLog 'Double click a table row to open PyPi in browser (online)'
    
    $count = $dataModel.Rows.Count
    $builtinCount = $builtinPackages.Count
    Write-PipLog "Now $count packages in the list, $builtinCount builtin"
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
    return $dataModel.Rows[$index].Select = $false
}

Function Tidy-Output($text) {
    $result = ($text -replace '\s*$', "`n")
    return $result
}

Function Clear-Rows() {
    $Script:outdatedOnly = $true
    $dataGridView.BeginInit()
    $dataModel.BeginLoadData()    
    $dataGridView.ClearSelection()
    
    $Script:inputFilter.Clear()
    $dataModel.DefaultView.RowFilter = $null    
    
    $dataModel.Clear()

    $dataModel.EndLoadData()
    $dataGridView.EndInit()   
}

Function Set-SelectedRow($selectedRow) {
    $dataModel.BeginLoadData();
    $Script:dataGridView.ClearSelection()
    foreach ($vRow in $Script:dataGridView.Rows) {
        if ($vRow.DataBoundItem.Row -eq $selectedRow) {
            $vRow.Selected = $true
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
            $action = $Script:actionsModel[$actionList.SelectedIndex]
            
            Write-PipLog ""
            Write-PipLog $action.Name ' ' $package.Package

            $args = New-Object System.Collections.ArrayList
            if ($Script:virtualenvCheckBox.Checked) {
                $args.add('--isolated')
            }
            $args.Add($package.Package)
            $result = $action.Execute($args)

            $logFrom = $Script:logView.TextLength 
            Write-PipLog (Tidy-Output $result)
            $logTo = $Script:logView.TextLength - $logFrom
            $dataModel.Rows[$i] | Add-Member -Force -MemberType NoteProperty -Name LogFrom -Value $logFrom
            $dataModel.Rows[$i] | Add-Member -Force -MemberType NoteProperty -Name LogTo -Value $logTo

            $dataModel.Columns['Status'].ReadOnly = $false
            if ($result -match ($action.Validation + $package.Package)) {
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
    Write-PipLog 'Double click a table row to open PyPi in browser (online)'
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

$env:PYTHONIOENCODING="utf-8"
$env:LC_CTYPE="utf-8"
Generate-Form | Out-Null
