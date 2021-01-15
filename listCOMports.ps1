# 1) load assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("String")
[void] [System.Reflection.Assembly]::LoadWithPartialName("Environment.NewLine")
[void] [System.Reflection.Assembly]::LoadWithPartialName("FormBorderStyle")

###
$newLine = "`r`n"

# 2) create window
$objForm = New-Object System.Windows.Forms.Form
$objForm.Backcolor="white"

#$objForm.StartPosition = "CenterScreen"
$objForm.StartPosition = "Manual"
$objForm.Location = New-Object System.Drawing.Point(10,10)

$objForm.Size = New-Object System.Drawing.Size(205, 540)
$objForm.Text = "Tobi"
$objForm.TopMost = $true
#$objForm.FormBorderStyle = 'FixedSingle'
$objForm.FormBorderStyle = 'None'
$objForm.Opacity = 0.40
$objForm.Add_MouseEnter({$objForm.Opacity = 0.85})
$objForm.Add_MouseLeave({$objForm.Opacity = 0.40})

# 4) Label
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,10)
$objLabel.Size = New-Object System.Drawing.Size(180,20)
$objLabel.Text = "COM Ports:"
$objForm.Controls.Add($objLabel)

# 5) Textbox
$objTextBox = New-Object System.Windows.Forms.TextBox
$objTextBox.Location = New-Object System.Drawing.Size(10,40)
$objTextBox.Size = New-Object System.Drawing.Size(180,425)
$objTextBox.Text = "foo1"
$objTextBox.Multiline = $true
$objTextBox.WordWrap = $true
$objTextBox.Enabled = $false
$objForm.Controls.Add($objTextBox)
$objTextBox.AppendText($newLine + "foo2")

# 3) add Close Button
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Location = New-Object System.Drawing.Size(115,500)
$btnClose.Size = New-Object System.Drawing.Size(75,25)
$btnClose.Text = "Close"
$btnClose.Name = "Close"
$btnClose.DialogResult = "Cancel"
$btnClose.Add_Click({$objForm.Close()})
$btnClose.Add_MouseEnter({$objForm.Opacity = 0.85})
$btnClose.Add_MouseLeave({$objForm.Opacity = 0.40})
# add to form
$objForm.Controls.Add($btnClose)

# 6) add refresh Button
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Location = New-Object System.Drawing.Size(10,500)
$btnRefresh.Size = New-Object System.Drawing.Size(75,25)
$btnRefresh.Text = "refresh"
$btnRefresh.Name = "refresh"
# add to form
$objForm.Controls.Add($btnRefresh)
$btnRefresh.Add_MouseEnter({$objForm.Opacity = 0.85})
$btnRefresh.Add_MouseLeave({$objForm.Opacity = 0.40})

###
$RefreshCOMPortsHandler = {
  $objTextBox.Text = [string](chgport.exe)
}
$btnRefresh.Add_Click($RefreshCOMPortsHandler)

## initial data
$objTextBox.Text = [string](chgport.exe)
### finally, show dialog
$objForm.ShowDialog()