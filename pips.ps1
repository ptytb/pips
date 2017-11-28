[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

Function Get-Bin($command) {
    (where.exe $command) | Select-Object -Index 0
}

$python_path = Get-Bin 'python'
$pip_path = Get-Bin 'pip'

$pos = 5
$dataGridView = $null
$logView = $null
$arrayModel = $null
$actionsModel = $null
$virtualenvCheckBox = $null

$form = New-Object Windows.Forms.Form
$form.Text = "pip package browser"
$form.Size = New-Object Drawing.Point 830, 840
$form.topmost = $false
$form.Icon = [system.drawing.icon]::ExtractAssociatedIcon($python_path)

Function Write-PipLog() {
    foreach ($obj in $args) {
        $logView.AppendText("$obj")
    }
    $logView.AppendText("`n")
    $logView.ScrollToCaret()
}

Function Add-TopWidget($widget) {
    $widget.Location = New-Object Drawing.Point $pos,15
    $widget.size = New-Object Drawing.Point 90,23
    $form.Controls.Add($widget)
    $Script:pos = $pos + 100
}

Function Add-Button ($name, $handler) {
    $button = New-Object Windows.Forms.Button
    $button.Text = $name
    $button.Add_Click($handler)
    Add-TopWidget($button)
}

Function Add-Buttons {
    Add-Button "Refresh" { Get-PythonPackages }
    Add-Button "All" { Select-PipPackages($true) }
    Add-Button "None" { Select-PipPackages($false) }
    Add-Button "Check Deps" { Check-PipDependencies }
    Add-Button "Execute" { Execute-PipAction }
}

Function Add-ComboBox {
    Function Make-PipActionItem($name, $code) {
        $action = New-Object psobject -Property @{Name=$name}
        $action | Add-Member ScriptMethod ToString { $this.Name } -Force
        $action | Add-Member ScriptMethod Execute $code
        return $action
    }

    $actionsModel = New-Object System.Collections.ArrayList
    $actionsModel.Add( (Make-PipActionItem 'Show'   {return (&pip show       $args)}) )
    $actionsModel.Add( (Make-PipActionItem 'Update' {return (&pip install -U $args)}) )
    $Script:actionsModel = $actionsModel

    $actionList = New-Object System.Windows.Forms.ComboBox
    $actionList.DataSource = $actionsModel
    $actionList.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $Script:actionsList = $actionList
    Add-TopWidget($actionList)    
}

Function Add-CheckBox($text, $code) {
    $checkBox = New-Object System.Windows.Forms.CheckBox
    $checkBox.Text = $text
    $checkBox.Add_Click($code)
    Add-TopWidget($checkBox)
    return ($checkBox)
}

Function Toggle-VirtualEnv ($state) {
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

Function Generate-Form {
    Add-Buttons

    Add-ComboBox

    $Script:virtualenvCheckBox = Add-CheckBox 'virtualenv' { Toggle-VirtualEnv $Script:virtualenvCheckBox.Checked }
    
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Location = New-Object Drawing.Point 7,40
    $dataGridView.Size = New-Object Drawing.Point 800,450
    $dataGridView.MultiSelect = $false
    $dataGridView.SelectionMode = [System.Windows.Forms.SelectionMode]::One
    $dataGridView.ColumnHeadersVisible = $true
    $dataGridView.RowHeadersVisible = $false
    $dataGridView.ReadOnly = $false
    $Script:dataGridView = $dataGridView
    $form.Controls.Add($dataGridView)

    $logView = New-Object System.Windows.Forms.RichTextBox
    $logView.Location = New-Object Drawing.Point 7,500
    $logView.Size = New-Object Drawing.Point 800,280
    $logView.ReadOnly = $true
    $Script:logView = $logView
    $form.Controls.Add($logView)
    
  #  $form.Add_Load({ Get-PythonPackages })
    $form.ShowDialog()
}

 Function Get-PythonPackages {
    Write-PipLog 'Updating package list... '
    Write-PipLog (&"$python_path" --version)
    Write-PipLog (&"$pip_path" --version)

    if ($dataGridView.columncount -gt 0) {
        $dataGridView.DataSource = $null
        $dataGridView.Columns.RemoveAt(0) 
    }
    
    $Column1 = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $Column1.width = 100
    $Column1.name = "Update"
    $DataGridView.Columns.Add($Column1) 

    $arrayModel = New-Object System.Collections.ArrayList
    $header = ("Package", "Version", "Latest", "Type")

    $args = New-Object System.Collections.ArrayList
    $args.Add('list')
    $args.Add('--outdated')
    $args.Add('--format=columns')
    if ($virtualenvCheckBox.Checked) {
        $args.Add('--isolated')
    }

    $pip_list = (&pip3 $args)
    $packages = $pip_list | Select-Object -Skip 2 | % { $_ -replace '\s+', ' ' }  | ConvertFrom-Csv -Header $header -Delimiter ' '
    $Script:procInfo = $packages
    $arrayModel.AddRange($procInfo)
    $Script:arrayModel = $arrayModel
    $dataGridView.DataSource = $arrayModel
    
    for ($i = 1; $i -lt $dataGridView.ColumnCount; $i++) {
        $dataGridView.Columns[$i].ReadOnly = $true
    }

    $form.refresh()    
    Write-PipLog 'Package list updated.'
}

Function Select-PipPackages($value) {
    for ($i = 0; $i -lt $dataGridView.RowCount; $i++) {
       $datagridview.Rows[$i].Cells['Update'].Value = $value
    }
}

Function Check-PipDependencies {
    Write-PipLog 'Checking dependencies...'

    $result = (pip3 check)
    if ($result.StartsWith('No broken')) {
        Write-PipLog "OK", $result
    } else {
        Write-PipLog "NOT OK", $result
    }
}

Function Execute-PipAction($action) {
    for ($i = 0; $i -lt $dataGridView.RowCount; $i++) {
       if ($dataGridView.Rows[$i].Cells['Update'].Value -eq $true) {
            $package = $arrayModel[$i]
            $action = $actionsModel[$actionList.SelectedIndex]
            Write-PipLog $action.Name ' ' $package.Package

            $args = New-Object System.Collections.ArrayList
            if ($virtualenvCheckBox.Checked) {
                $args.add('--isolated')
            }
            $args.Add($package.Package)
            $result = $action.Execute($args)

            Write-PipLog $result
       }
    }
}

Generate-Form
