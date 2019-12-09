

    $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
    $Font = New-Object System.Drawing.Font("Tempus Sans ITC",9,[System.Drawing.FontStyle]::Bold)
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

    #Firm Label
    $FontBanksy = New-Object System.Drawing.Font("Tempus Sans ITC",8,[System.Drawing.FontStyle]::Bold)
    $FOKINLAB = New-Object System.Windows.Forms.Label
    $FOKINLAB.Location = ('290,260')
    $FOKINLAB.Text = "Created By Fokin"
    $FOKINLAB.Font = $FontBanksy
    $FOKINLAB.BackColor =  'Transparent'
    $FOKINLAB.AutoSize = $True


    $FinderCheckGroup1 = New-Object System.Windows.Forms.GroupBox
    $FinderCheckGroup1.Location = ('10,10')
    $FinderCheckGroup1.Size = ('220,40')
    $FinderCheckGroup1.BackColor = 'Transparent'
    $FinderCheckGroup1.Text = "Formats: "
    $FinderCheckGroup1.Font = $Font
    
    
    #CheckBox
    $FinderCheckBox1 = New-Object System.Windows.Forms.CheckBox
    $FinderCheckBox1.Location = New-Object System.Drawing.Size('10','15')
    $FinderCheckBox1.Text = "*.xml"
    $FinderCheckBox1.BackColor = 'Transparent'
    $FinderCheckBox1.Name = '*.xml'
    $FinderCheckBox1.AutoSize = $True
    $FinderCheckBox1.Checked = $True
    #CheckBox
    $FinderCheckBox2 = New-Object System.Windows.Forms.CheckBox
    $FinderCheckBox2.Location = New-Object System.Drawing.Size('65','15')
    $FinderCheckBox2.Text ="*.txt"
    $FinderCheckBox2.BackColor = 'Transparent'
    $FinderCheckBox2.Name = '*.txt'
    $FinderCheckBox2.AutoSize = $True
    $FinderCheckBox2.Checked = $True
    #CheckBox
    $FinderCheckBox3 = New-Object System.Windows.Forms.CheckBox
    $FinderCheckBox3.Location = New-Object System.Drawing.Size('115','15')
    $FinderCheckBox3.Text ="*.html"
    $FinderCheckBox3.BackColor = 'Transparent'
    $FinderCheckBox3.Name = '*.html'
    $FinderCheckBox3.AutoSize = $True
    $FinderCheckBox3.Checked = $True
    #CheckBOx
    $FinderCheckBox4 = New-Object System.Windows.Forms.CheckBox
    $FinderCheckBox4.Location = New-Object System.Drawing.Size('175','15')
    $FinderCheckBox4.Text ="*.log"
    $FinderCheckBox4.BackColor = 'Transparent'
    $FinderCheckBox4.Name = '*.log'
    $FinderCheckBox4.AutoSize = $True
    $FinderCheckBox4.Checked = $True

    $FinderSearch  = New-Object System.Windows.Forms.TextBox
    $FinderSearch.Location = ('10,55')
    $FinderSearch.Multiline = $true
    $FinderSearch.Size = ('220,100')
    $FinderSearch.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D

    $FinderButton = New-Object System.Windows.Forms.Button
    $FinderButton.Location = ('240,15')
    $FinderButton.Size = ('90,60')
    $FinderButton.Font = $FontBanksy
    $FinderButton.Text = "START FINDER"

    $FinderButton.add_Click({
       $ARRAY1 = ''
       $ARRAY1 = "{" 

        foreach($G in $FinderCheckGroup1.Controls)
        {
            
            if($G.Checked)
            { $ARRAY1 += $G.Name + ","}
            else
            {
             
            }
        }
       $Array1 = $ARRAY1.split(","[-1])
       $ARRAY1 += "}"
       $FinderSearch.Text = $ARRAY1
    
    })

    $FinderCheckGroup1.Controls.AddRange(@($FinderCheckBox1,$FinderCheckBox2,$FinderCheckBox3,$FinderCheckBox4))
    $FinderForm.Controls.Add($FinderButton)
    $FinderForm.Controls.Add($FinderSearch)
    #$FinderForm.Controls.Add($FinderCheckBox1)
    #$FinderForm.Controls.Add($FinderCheckBox2)
    #$FinderForm.Controls.Add($FinderCheckBox3)
    $FinderForm.Controls.Add($FinderCheckGroup1)
    #$FinderForm.Controls.AddRange(@($FinderCheckBox1,$FinderCheckBox2))
    $FinderForm.ShowDialog()
