# Load Windows Forms & Drawing classes.
function GENERATOR{
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$VRX = ('1','2','3','4','5','6')
$VRQ = ('1','2','3','4','5','6','7')
# Create base form.

$Font = New-Object System.Drawing.Font("Times New Roman",8,[System.Drawing.FontStyle]::Bold)


$objForm = New-Object System.Windows.Forms.Form 
$Image =  [system.drawing.image]::FromFile("C:\Users\ks_fokin\Downloads\NTS.jpg")
$objForm.SizeGripStyle = "Hide"
$objForm.BackgroundImage = $Image
$objForm.BackgroundImageLayout = "None"
$objForm.Text = "Программа для безумного управления сервисами"
#$objForm.Size = New-Object System.Drawing.Size(1000,1000) 
$objForm.StartPosition = "CenterScreen"
$objForm.Height = '340'
$objForm.Width = $Image.Width
$objForm.AutoSizeMode = "GrowAndShrink"

# Configure keyboard intercepts for ESC & ENTER.

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") 
    {
        $x=$objListBox.SelectedItem
        $objForm.Close()
    }
})
$objForm.Add_KeyDown({
    if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()
    }
})


#GROUP BOX SRED
$MyGroupBox = New-Object System.Windows.Forms.GroupBox
$MyGroupBox.Location = '5,180'
$MyGroupBox.size = '120,80'
$MyGroupBox.Font = $Font
$MyGroupBox.text = "ВЫБОР СРЕДЫ:"
$MyGroupBox.Backcolor = 'Transparent'
#$objForm.Controls.Add($MyGroupBox)

#RADIO VRX
$RadioVRX = New-Object System.Windows.Forms.RadioButton
$RadioVRX.Location = New-Object System.Drawing.Size(10,15)
$RadioVRX.Checked = $False
$RadioVRX.Text = "VRX"

#RADIO VRQ
$RadioVRQ = $RadioButton2 = New-Object System.Windows.Forms.RadioButton
$RadioVRQ.Location = New-Object System.Drawing.Size(10,35)
$RadioVRQ.Text = "VRQ"
$RadioVRQ.Checked = $False
$eventSRED = {
            if($RadioVRQ.Checked)
            {
             foreach($i in $VRQ){
             $Combo_Srez.Controls.Add($i)
             }
            }
            elseif ($RadioVRX.Checked)
            {
             foreach($i in $VRX){
             $Combo_Srez.Controls.Add($i)
             }
            }
        }
$RadioVRQ.Add_Click($eventSRED)
$RadioVRX.Add_Click($eventSRED)
#
$objForm.Controls.AddRange(@($MyGroupBox))
$MyGroupBox.Controls.AddRange(@($RadioVRX,$RadioVRQ))


$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,122)
$objLabel.Font = $Font
$objLabel.AutoSize = 'True'
$objLabel.BackColor = 'Transparent'
$objLabel.Text = "!!!!!!!!!!!!!!"
$objLabel.Visible = 'TRUE'
$objForm.Controls.Add($objLabel) 

#TEXTBOX
$TextBox = New-Object System.Windows.Forms.TextBox
$TextBox = New-Object System.Drawing.Size(10,30)
$TextBox.Visible = 'False' 
$TextBox


#GROUP BOX TRIGGERS
$MyGroupBox2 = New-Object System.Windows.Forms.GroupBox
$MyGroupBox2.Location = '130,180'
$MyGroupBox2.size = '115,80'
$MyGroupBox2.Font = $Font
$MyGroupBox2.text = "ТРИГГЕРЫ:"
$MyGroupBox2.Backcolor = 'Transparent'

#КОНТУР
$RadioContur = New-Object System.Windows.Forms.RadioButton
$RadioContur.Location = New-Object System.Drawing.Size(10,10)
$RadioContur.Text = "Контур"
$RadioContur.BackColor = 'Transparent'
$RadioContur.Checked = 'True'
#МАГАЗИНЫ
$RadioMAG = New-Object System.Windows.Forms.RadioButton
$RadioMAG.Location = New-Object System.Drawing.Size(10,30)
$RadioMAG.Text = "Магазины"
$RadioMAG.BackColor = 'Transparent'
$eventMAG = {
             if($RadioMAG.Checked){
             $Combo_Srez.Visible = $False
             $TextBox.Visible = $True
             }
             elseif($RadioContur.Checked){
             $Combo_Srez.Visible = $True
             $TextBox.Text = ''
             $TextBox.Visible = $False
             }
            }
#ИНТЕРФЕЙСЫ
$RadioINT = New-Object System.Windows.Forms.RadioButton
$RadioINT.Location = New-Object System.Drawing.Size(10,50)
$RadioINT.Text = "Интерфейс"
#
$objForm.Controls.Add($MyGroupBox2)
$MyGroupBox2.Controls.AddRange(@($RadioContur,$RadioMAG,$RadioINT)) 



#GROUP BLOCK CHOICE
$MyGroupBox3 = New-Object System.Windows.Forms.GroupBox
$MyGroupBox3.Location = '250,180'
$MyGroupBox3.size = '140,80'
$MyGroupBox3.Font = $Font
$MyGroupBox3.text = "ВЫБОР СТАНЦИИ:"
$MyGroupBox3.Backcolor = 'Transparent'


#COMBO
$Combo_Srez = New-Object System.Windows.Forms.ComboBox
$Combo_Srez.AutoSize = 'True'
$Combo_Srez.Location = New-Object System.Drawing.Size(10,30)
$Combo_Srez.Text = 'Выберите станцию'
if($RadioVRX.Checked -eq 'True')
{
$SREZ = $RadioVrx.Text
foreach($i in $VRX)
{
  $Combo_Srez.Items.Add($i);  
}
}
elseif ($RadioVRQ.Checked)
{
$SREZ = $RadioVRQ.Text
$Combo_Srez.Items.Add($i); 
}
else
{
  $Combo_Srez.Items.Add($i);  
}
$objForm.Controls.Add($MyGroupBox3)
$MyGroupBox3.Controls.Add($Combo_Srez) 


# Create BUTTON FOR START REDEPLOY WILDFLY
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(10,270)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "RESTART WILDFLY"
$OKButton.AutoSize = 'True'
$OKButton.Add_Click(
{
    
    if($RadioVRX.Checked)
    {
      $SRED = 'vrx'  
    }
    elseif ($RadioVRQ.Checked)
    {
      $SRED = 'vrq'
    }

    if ($RadioContur.Checked)
    {
     [System.Windows.Forms.MessageBox]::Show("ajb","Контур",'OK','Info')
     $CONT = "ajb"
    }
    elseif ($RadioMAG.Checked)
    {
     [System.Windows.Forms.MessageBox]::Show("a","МАГАЗИН")
     $CONT = "a"
    }
    elseif ($RadioINT.Checked = $true)
    {
     [System.Windows.Forms.MessageBox]::Show("int","ИНТЕРФЕЙС")
     $CONT = "int"
    }
    $Combo_Srez.SelectItem
    if($Combo_Srez.Text -eq 'Выберите станцию')
    {
      [System.Windows.Forms.MessageBox]::Show($Combo_Srez.Text,"Ошибка выбора",'RetryCancel','ERROR')
    }
    else
    {
     [System.Windows.Forms.MessageBox]::Show($Combo_Srez.Text,"Выбор сделан",'OK','WARNING')
    }
    $SERVER = 'fobo-'+ $SRED + "-" + $CONT + $Combo_Srez.SelectedItem
    [System.Windows.Forms.MessageBox]::Show($Server,"Выбран "+ $SERVER + ". Подтверждаем?",'OKCANCEL','INFO')

})

#$OKButton.Add_Click({$x=$objListBox.SelectedItem;$objForm.Close()})
$objForm.Controls.Add($OKButton)

# Cancel EXIT Button
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(470,270)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Отменa"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton)






#TABLO
$objListBox = New-Object System.Windows.Forms.ListBox 
$objListBox.Location = New-Object System.Drawing.Size(250,320) 
$objListBox.Size = New-Object System.Drawing.Size(100,20)
$objListBox.Height = 90










$objForm.TopMost = $true
$objForm.Add_Shown({$objForm.Activate()})
$objForm.ShowDialog()
}
GENERATOR




<#[void] $objListBox.Items.Add($SREZ + ' 1')
[void] $objListBox.Items.Add($SREZ + ' 2')
[void] $objListBox.Items.Add($SREZ + ' 3')
[void] $objListBox.Items.Add($SREZ + ' 4')
[void] $objListBox.Items.Add($SREZ + ' 5')
[void] $objListBox.Items.Add($SREZ + ' 6')
#$objForm.Controls.Add($objListBox) 

#LABEL SREDA
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,122)
$objLabel.Font = $Font
$objForm
$objLabel.AutoSize = 'True'
$objLabel.BackColor = 'Transparent'
$objLabel.Text = "ВЫБЕРИТЕ СРЕДУ:"
#$objForm.Controls.Add($objLabel) 

#LABEL TRIGGERS
$objLabel1 = New-Object System.Windows.Forms.Label
$objLabel1.Location = New-Object System.Drawing.Size(130,122)
$objLabel1.BackColor = 'Transparent' 
$objLabel1.Font = $Font
$objLabel1.Size = New-Object System.Drawing.Size(100,20) 
$objLabel1.Text = "ТРИГГЕРЫ:"
#$objForm.Controls.Add($objLabel1) 

#LABEL COMBO
$objLabel2 = New-Object System.Windows.Forms.Label
$objLabel2.Location = New-Object System.Drawing.Size(250,122) 
$objLabel2.Font = $Font
$objLabel2.BackColor = 'Transparent'
$objLabel2.Text = "МАШИНА:"
$objForm.Controls.Add($objLabel2)

#>