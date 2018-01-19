﻿[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

Function Get-Bin($command) {
    (where.exe $command) | Select-Object -Index 0
}

$python_path = Get-Bin 'python'
$pip_path = Get-Bin 'pip'

$lastWidgetLeft = 5
$dataGridView = $null
$logView = $null
$arrayModel = $null
$actionsModel = $null
$virtualenvCheckBox = $null

$formLoaded = $false

Function Write-PipLog() {
    foreach ($obj in $args) {
        $logView.AppendText("$obj")
    }
    $logView.AppendText("`n")
    $logView.ScrollToCaret()
}

Function Add-TopWidget($widget) {
    $widget.Location = New-Object Drawing.Point $lastWidgetLeft,15
    $widget.size = New-Object Drawing.Point 90,23
    $Script:form.Controls.Add($widget)
    $Script:lastWidgetLeft = $lastWidgetLeft + 100
}

Function Add-Button ($name, $handler) {
    $button = New-Object Windows.Forms.Button
    $button.Text = $name
    $button.Add_Click($handler)
    Add-TopWidget($button)
}

Function Add-Buttons {
    Add-Button "Check Updates" { Get-PythonPackages }
    Add-Button "Select All" { Select-PipPackages($true) }
    Add-Button "Select None" { Select-PipPackages($false) }
    Add-Button "Check Deps" { Check-PipDependencies }
    Add-Button "Execute:" { Execute-PipAction }
}

Function Add-ComboBox {
    Function Make-PipActionItem($name, $code, $validation) {
        $action = New-Object psobject -Property @{Name=$name; Validation=$validation}
        $action | Add-Member ScriptMethod ToString { $this.Name } -Force
        $action | Add-Member ScriptMethod Execute $code
        return $action
    }

    $actionsModel = New-Object System.Collections.ArrayList

    $actionsModel.Add( (Make-PipActionItem 'Show'      {return (&pip show       $args)} `
        '.*' ) )

    $actionsModel.Add( (Make-PipActionItem 'Update'    {return (&pip install -U $args)} `
        'Successfully installed |Installing collected packages:\s*(\s*\S*,\s*)*' ) )

    $actionsModel.Add( (Make-PipActionItem 'Download'  {return (&pip download   $args)} `
        'Successfully downloaded ' ) )

    $actionsModel.Add( (Make-PipActionItem 'Uninstall' {return (&pip uninstall  $args)} `
        'Successfully uninstalled ' ) )

    $Script:actionsModel = $actionsModel

    $actionList = New-Object System.Windows.Forms.ComboBox
    $actionList.DataSource = $actionsModel
    $actionList.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $Script:actionList = $actionList
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
    $form.Icon = [system.drawing.icon]::ExtractAssociatedIcon($python_path)
    $Script:form = $form

    Add-Buttons
    Add-ComboBox
    $Script:virtualenvCheckBox = Add-CheckBox 'virtualenv' { Toggle-VirtualEnv $Script:virtualenvCheckBox.Checked }
    Add-Button "Install..." { Generate-FormInstall }
    
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Location = New-Object Drawing.Point 7,40
    $dataGridView.Size = New-Object Drawing.Point 800,450
    $dataGridView.MultiSelect = $false
    $dataGridView.SelectionMode = [System.Windows.Forms.SelectionMode]::One
    $dataGridView.ColumnHeadersVisible = $true
    $dataGridView.RowHeadersVisible = $false
    $dataGridView.ReadOnly = $false
    $dataGridView.AllowUserToResizeRows = $false
    $dataGridView.AllowUserToResizeColumns = $false
    $dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $Script:dataGridView = $dataGridView
    $form.Controls.Add($dataGridView)

    $logView = New-Object System.Windows.Forms.RichTextBox
    $logView.Location = New-Object Drawing.Point 7,500
    $logView.Size = New-Object Drawing.Point 800,280
    $logView.ReadOnly = $true
    $logView.Multiline = $true
    $logView.Font = New-Object System.Drawing.Font("Consolas", 11)
    $Script:logView = $logView
    $form.Controls.Add($logView)

    Function Highlight-LogFragment() {
        $rowIndex = $Script:dataGridView.CurrentRow.Index
        $row = $Script:dataGridView.Rows[$rowIndex]

        $Script:logView.SelectAll()
        $Script:logView.SelectionBackColor = $Script:logView.BackColor

        if (Get-Member -inputobject $row -name "LogFrom" -Membertype Properties) {
            $Script:logView.Select($row.LogFrom, $row.LogTo)
            $Script:logView.SelectionBackColor = [Drawing.Color]::Yellow
            $Script:logView.ScrollToCaret()
        }
    }
    
    $dataGridView.Add_CellMouseClick({Highlight-LogFragment})

    $form.Add_Load({ $Script:formLoaded = $true })
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
    $header = ("Package", "Installed", "Latest", "Type", "Status")

    $args = New-Object System.Collections.ArrayList
    $args.Add('list')
    $args.Add('--outdated')
    $args.Add('--format=columns')
    if ($Script:virtualenvCheckBox.Checked) {
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

    $Script:form.refresh()    
    Write-PipLog 'Package list updated.'
}

Function Select-PipPackages($value) {
    for ($i = 0; $i -lt $dataGridView.RowCount; $i++) {
       $datagridview.Rows[$i].Cells['Update'].Value = $value
    }
}

Function Set-Unchecked($index) {
    return $datagridview.Rows[$index].Cells['Update'].Value = $false
}

Function Tidy-Output($text) {
    $result = ($text -replace '(\.?\s*$)', "`n")
    return $result
}

Function Check-PipDependencies {
    Write-PipLog 'Checking dependencies...'

    $result = (pip3 check)
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
    for ($i = 0; $i -lt $dataGridView.RowCount; $i++) {
       if ($dataGridView.Rows[$i].Cells['Update'].Value -eq $true) {
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
            $dataGridView.Rows[$i] | Add-Member -Force -MemberType NoteProperty -Name LogFrom -Value $logFrom
            $dataGridView.Rows[$i] | Add-Member -Force -MemberType NoteProperty -Name LogTo -Value $logTo

            if ($result -match ($action.Validation + $package.Package)) {
                 $dataGridView.Rows[$i].Cells['Status'].Value = "OK"
                 Set-Unchecked $i
            } else {
                 $dataGridView.Rows[$i].Cells['Status'].Value = "Failed"
            }
       }
    }

    Write-PipLog ''
    Write-PipLog '----'
    Write-PipLog 'All tasks finished'
    Write-PipLog '----'
    Write-PipLog ''
}

Generate-Form
