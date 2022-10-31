# GETMAC ADDRESS
$network = Get-WmiObject win32_networkadapterconfiguration | select description, macaddress | where { $_.description -Like "*ether*" }
$network_mac  = $network.macaddress
$network_card = $network.description
Set-Clipboard -Value $network_mac
Write-Output $network_card

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = 'MAC Address '
$objForm.Size = New-Object System.Drawing.Size(300,200) 
$objForm.StartPosition = "CenterScreen"

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") 
    {
        $x=$objListBox.SelectedItem
        Set-Clipboard -Value $x
        $objForm.Close()
     }
})


$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "Copy Mac"
$OKButton.Add_Click({

        $x=$objListBox.SelectedItem
        Set-Clipboard -Value $X
        $objForm.Close()
})


$objForm.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton)

$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20) 
$objLabel.Size = New-Object System.Drawing.Size(280,20) 
$objLabel.Text = "Network Card: $network_card"
$objForm.Controls.Add($objLabel) 

$objListBox = New-Object System.Windows.Forms.ListBox 
$objListBox.Location = New-Object System.Drawing.Size(10,40) 
$objListBox.Size = New-Object System.Drawing.Size(260,20) 
$objListBox.Height = 80

[void] $objListBox.Items.Add($network_mac)


$objForm.Controls.Add($objListBox) 
$objListBox.SelectedIndex = 0
$objListBox.Focus()

$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()

