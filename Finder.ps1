#Secret Project Finder

    $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
    $Font = New-Object System.Drawing.Font("Colibri",9,[System.Drawing.FontStyle]::Bold)
    $Image =  [system.drawing.image]::FromFile("\\dubovenko\D\SOFT\wallapers\Fobo_INTERFACE.jpg")

    #Create FOBO FORM
    $FinderForm = New-Object System.Windows.Forms.Form
    $FinderForm.SizeGripStyle = "Hide"
    $FinderForm.BackgroundImage = $Image
    $FinderForm.BackgroundImageLayout = "None"
    $FinderForm.Width = $Image.Width
    $FinderForm.Height = $Image.Height
    $FinderForm.StartPosition = "CenterScreen"
    $FinderForm.TopMost = $true
    $FinderForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $FinderForm.Text = "ЦУП FINDER"
    $FinderForm.TopMost = $True
    $FinderForm.Icon = $Icon

    


    $FinderRadioFormat1 = New-Object System.Windows.Forms.CheckBox
    $FinderRadioFormat1.Location = New-Object System.Drawing.Size('10','5')
    $FinderRadioFormat1.Text = "*.xml"
    $FinderRadioFormat1.BackColor = 'Transparent'
    $FinderRadioFormat1.Name = '.xml'
    $FinderRadioFormat1.AutoSize = $True
    $FinderRadioFormat1.Checked = $True

    $FinderRadioFormat2 = New-Object System.Windows.Forms.CheckBox
    $FinderRadioFormat2.Location = New-Object System.Drawing.Size('60','5')
    $FinderRadioFormat2.Text ="*.txt"
    $FinderRadioFormat2.BackColor = 'Transparent'
    $FinderRadioFormat2.Name = '.txt'
    $FinderRadioFormat2.AutoSize = $True
    $FinderRadioFormat2.Checked = $True

    $FinderRadioFormat3 = New-Object System.Windows.Forms.CheckBox
    $FinderRadioFormat3.Location = New-Object System.Drawing.Size('105','5')
    $FinderRadioFormat3.Text ="*.html"
    $FinderRadioFormat3.BackColor = 'Transparent'
    $FinderRadioFormat3.Name = '.html'
    $FinderRadioFormat3.AutoSize = $True
    $FinderRadioFormat3.Checked = $True

    $FinderForm.Controls.Add($FinderRadioFormat1)
    $FinderForm.Controls.Add($FinderRadioFormat2)
    $FinderForm.Controls.Add($FinderRadioFormat3)
    #$FinderForm.Controls.AddRange(@($FinderRadioFormat1,$FinderRadioFormat2))
    $FinderForm.ShowDialog()
