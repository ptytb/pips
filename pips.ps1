[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

Function Get-Bin($command) {
    (where.exe $command) | Select-Object -Index 0
}

Function Get-PythonExe() {
    $Script:interpretersComboBox.SelectedItem.PythonExe
    #Get-Bin 'python'
}

Function Get-PipExe() {
    $Script:interpretersComboBox.SelectedItem.PipExe
    #Get-Bin 'pip'
}

$lastWidgetLeft = 5
$lastWidgetTop = 5
$widgetLineHeight = 23
$dataGridView = $null
$logView = $null
$arrayModel = $null
$actionsModel = $null
$virtualenvCheckBox = $null
$header = ("Update", "Package", "Installed", "Latest", "Type", "Status")
$csv_header = ("Package", "Installed", "Latest", "Type", "Status")
$formLoaded = $false


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
}


Function Add-Buttons {
    Add-Button "Check Updates" { Get-PythonPackages }
    Add-Button "Select All" { Select-PipPackages($true) }
    Add-Button "Select None" { Select-PipPackages($false) }
    Add-Button "Check Deps" { Check-PipDependencies }
    Add-Button "Execute:" { Execute-PipAction }
}

Function Add-ComboBoxActions {
    Function Make-PipActionItem($name, $code, $validation) {
        $action = New-Object psobject -Property @{Name=$name; Validation=$validation}
        $action | Add-Member ScriptMethod ToString { $this.Name } -Force
        $action | Add-Member ScriptMethod Execute $code
        return $action
    }

    $actionsModel = New-Object System.Collections.ArrayList
    $Add = { param($a) $actionsModel.Add($a) | Out-Null }

    & $Add (Make-PipActionItem 'Show'      {return (& (Get-PipExe) show       $args)} `
        '.*' ) 

    & $Add (Make-PipActionItem 'Update'    {return (& (Get-PipExe) install -U $args)} `
        'Successfully installed |Installing collected packages:\s*(\s*\S*,\s*)*' )

    & $Add (Make-PipActionItem 'Download'  {return (& (Get-PipExe) download   $args)} `
        'Successfully downloaded ' )

    & $Add (Make-PipActionItem 'Uninstall' {return (& (Get-PipExe) uninstall  $args)} `
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
        
        $arch = Test-is64Bit $path
        $pip = (Split-Path -Path $path) + "\Scripts\pip.exe"

        $action = New-Object psobject -Property @{Arch=$arch; PythonExe=$path; PipExe=$pip}
        $action | Add-Member ScriptMethod ToString { "[{0}] {1}" -f $this.Arch.FileType,$this.PythonExe } -Force

        $items.Add($action) | Out-Null
    }

    $list = (where.exe 'python')
    foreach ($path in $list) {
        Get-InterpreterRecord $path
    }

    foreach ($d in dir "$env:LOCALAPPDATA\Programs\Python") {
        if ($d -is [System.IO.DirectoryInfo]) {
            Get-InterpreterRecord (${d}.FullName + "\python.exe")
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
    
}

Function Generate-Form {
    $form = New-Object Windows.Forms.Form
    $form.Text = "pip package browser"
    $form.Size = New-Object Drawing.Point 830, 840
    $form.topmost = $false
    $iconPath = Get-Bin 'pip'
    $form.Icon = [system.drawing.icon]::ExtractAssociatedIcon($iconPath)
    $Script:form = $form

    Add-Buttons
    Add-ComboBoxActions
    $Script:virtualenvCheckBox = Add-CheckBox 'virtualenv' { Toggle-VirtualEnv $Script:virtualenvCheckBox.Checked }
    Add-Button "Search..." { Clear-Rows; Generate-FormInstall }

    NewLine-TopLayout

    Add-Label "Filter results:"
    Add-Input {
        param($input)
        
        if ($Script:dataGridView.CurrentRow) {
            # Keep selection while filter is being changed
            $selectedRow = $Script:dataGridView.CurrentRow.DataBoundItem.Row
        }

        $searchText = $input.Text
        $query = "Package LIKE '%{0}%'" -f $searchText
        if ($searchText.Length -gt 0) {
            $Script:dataModel.DefaultView.RowFilter = $query
        } else {
            $Script:dataModel.DefaultView.RowFilter = $null
        }

        if ($selectedRow) {
            Set-SelectedRow $selectedRow
        }
    }

    Add-HorizontalSpacer
    Add-Label "Active Interpreter:"
    Add-ComboBoxInterpreters
    
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Location = New-Object Drawing.Point 7,($Script:lastWidgetTop + $Script:widgetLineHeight)
    $dataGridView.Size = New-Object Drawing.Point 800,450
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
    
    $dataModel = New-Object System.Data.DataTable
    $dataGridView.DataSource = $dataModel    
    $Script:dataModel = $dataModel
    
    foreach ($c in $header) {
        if ($c -eq "Update") {
            $column = New-Object System.Data.DataColumn $c,([bool])
        } else {
            $column = New-Object System.Data.DataColumn $c,([string])
            $column.ReadOnly = $true
        }
        $dataModel.Columns.Add($column)
    }

    $Script:dataGridView = $dataGridView
    $form.Controls.Add($dataGridView)

    $Script:arrayModel = New-Object System.Collections.ArrayList

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
        
        $row = $Script:dataGridView.CurrentRow.DataBoundItem.Row

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
    $form.Add_Load({ $Script:formLoaded = $true })
    $form.ShowDialog()
}

 Function Get-PythonPackages {
    Write-PipLog 'Updating package list... '
    Write-PipLog (& (&Get-PythonExe) --version)
    Write-PipLog (& (&Get-PipExe) --version)

    Clear-Rows

    $args = New-Object System.Collections.ArrayList
    $args.Add('list')
    $args.Add('--outdated')
    $args.Add('--format=columns')
    if ($Script:virtualenvCheckBox.Checked) {
        $args.Add('--isolated')
    }

    $pip_list = & (&Get-PipExe) $args
    $packages = $pip_list | Select-Object -Skip 2 | % { $_ -replace '\s+', ' ' }  | ConvertFrom-Csv -Header $csv_header -Delimiter ' '
    $Script:procInfo = $packages
    $arrayModel.AddRange($procInfo)
    $Script:arrayModel = $arrayModel
    
    for ($n = 0; $n -lt $procInfo.Count; $n++) {
        $row = $dataModel.NewRow()        
        $row['Package'] = $arrayModel[$n].Package
        $row['Installed'] = $arrayModel[$n].Installed
        $row['Latest'] = $arrayModel[$n].Latest
        $row['Type'] = $arrayModel[$n].Type
        $dataModel.Rows.Add($row)
    }

    $Script:form.refresh()
    Write-PipLog 'Package list updated.'
}

Function Select-PipPackages($value) {
    for ($i = 0; $i -lt $dataModel.Rows.Count; $i++) {
       $dataModel.Rows[$i].Update = $value
    }
}

Function Set-Unchecked($index) {
    return $dataModel.Rows[$index].Update = $false
}

Function Tidy-Output($text) {
    $result = ($text -replace '(\.?\s*$)', "`n")
    return $result
}

Function Clear-Rows() {
    for ($i = $dataModel.Rows.Count; $i -ge 1; $i--) {
        $dataModel.Rows[$i - 1].Delete()
    }
    $Script:dataModel.Clear()
}

Function Set-SelectedRow($selectedRow) {
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

    $result = & (&Get-PipExe) check
    $result = Tidy-Output $result
    
    if ($result.StartsWith('No broken')) {
        Write-PipLog "OK"
        Write-PipLog $result
    } else {
        Write-PipLog "NOT OK"
        Write-PipLog $result
    }
}

Function Execute-PipAction($action) {
    for ($i = 0; $i -lt $dataModel.Rows.Count; $i++) {
       if ($dataModel.Rows[$i].Update -eq $true) {
            $package = $arrayModel[$i]
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
    Write-PipLog 'All tasks finished'
    Write-PipLog '----'
    Write-PipLog ''
}

function Test-is64Bit {
    param($FilePath="$env:windir\notepad.exe")

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

Generate-Form
