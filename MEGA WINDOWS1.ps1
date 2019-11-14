# Load Windows Forms & Drawing classes.
$Global:RLS=''

#RESTART WILDFLY DELETE FILES
Function KillWildfly([string]$SRV)
{
    Progress
    #Get-Process -Name java -ComputerName $SRV -ErrorAction SilentlyContinue | Format-List
    (Get-WmiObject -Class Win32_Process -ComputerName $SRV -Filter "name='java.exe'").terminate() | Out-Null
    Write-Output "Process Java was Terminated!"    
    Get-Service -Name Wildfly -ComputerName $SRV -ErrorAction SilentlyContinue | Stop-Service
    Start-Sleep -Seconds 3
    if($RLS -eq '19'){
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "*.backup","*.deployed","*.failed" | Remove-Item
        #Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\tmp\" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\data\" -Recurse -Force -ErrorAction SilentlyContinue
    }
    if($RLS -eq '20'){
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "*.backup","*.deployed","*.readclaim.*","*.failed","*.facade*","*.transfer*" | Remove-Item
        #Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\tmp\" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\data\" -Recurse -Force -ErrorAction SilentlyContinue
    }
    if($RLS -eq '21'){
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "*.backup","*.deployed","*.readclaim.*","*.failed","*.facade*","*.transfer*" | Remove-Item
        #Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\tmp\" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\data\" -Recurse -Force -ErrorAction SilentlyContinue
    }
    if($RLS -eq '22'){
        Get-ChildItem -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "*.backup","*.deployed","*.readclaim.*","*.failed","*.facade*","*.transfer*" | Remove-Item
        #Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\tmp\" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "\\$SRV\C`$\NTSwincash\jboss\wildfly10\standalone\data\" -Recurse -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 5
    Get-Service -Name Wildfly -ComputerName $server | Start-Service
    Write-Host "Работа программы завершена, сейчас будет открыта директория с файлами сервисов!"
    Start-Sleep -Seconds 2
    Invoke-Item "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"
}


Function Progress(){

$ProcessForm = New-Object System.Windows.Forms.Form
$ProcessForm.SizeGripStyle = "Hide"
$ProcessForm.BackgroundImage = $ImageRelease
$ProcessForm.BackgroundImageLayout = "None"
$ProcessForm.Size = New-Object System.Drawing.Size(250,110)
$ProcessForm
$ProcessForm.StartPosition = "CenterScreen"
$ProcessForm.TopMost = $true
$ProcessForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
$ProcessForm.Text = "Процесс выполнения задания"

$pbrTest = New-Object System.Windows.Forms.ProgressBar
$pbrTest.Maximum = 250
$pbrTest.Minimum = 0
$pbrTest.Location = new-object System.Drawing.Size(10,10)
$pbrTest.size = new-object System.Drawing.Size(200,50)
$i = 0


$pbrTest.Add_Click({
   
    While ($i -le 250) {
        $pbrTest.Value = $i
        Start-Sleep -m 100
        "VALLUE EQ"
        $i
        $i += 1
    }
    $ProcessForm.Close()
})

$ProcessForm.Controls.Add($pbrTest)
$ProcessForm.Add_Shown({$ProcessForm.Activate()})
$ProcessForm.Controls.AddRange(@($ReleaseButton0,$ReleaseButton1,$ReleaseButton2,$ReleaseButton3))
$ProcessForm.ShowDialog()

}

Function RELEASE_WINDOW(){

$FontRelease = New-Object System.Drawing.Font("Colibri",10,[System.Drawing.FontStyle]::Bold)
$ImageRelease =  [system.drawing.image]::FromFile("C:\Users\ks_fokin\Downloads\RLS2.png")

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$ReleaseForm = New-Object System.Windows.Forms.Form
$ReleaseForm.SizeGripStyle = "Hide"
$ReleaseForm.BackgroundImage = $ImageRelease
$ReleaseForm.BackgroundImageLayout = "None"
$ReleaseForm.Size = New-Object System.Drawing.Size(150,160)
$ReleaseForm
$ReleaseForm.StartPosition = "CenterScreen"
$ReleaseForm.TopMost = $true
$ReleaseForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
$ReleaseForm.Text = "ВЫБЕРИТЕ РЕЛИЗ"




#Release Button 0
$ReleaseButton0 = New-Object System.Windows.Forms.RadioButton
$ReleaseButton0.Location = New-Object System.Drawing.Size(10,15)
$ReleaseButton0.Text = 'Релиз 19.0.0'
$ReleaseButton0.AutoSize = 'True'
$ReleaseButton0.Font = $FontRelease
$ReleaseButton0.Backcolor = 'Transparent'
$ReleaseButton0.Checked = $True
#Release Button 1
$ReleaseButton1 = New-Object System.Windows.Forms.RadioButton
$ReleaseButton1.Location = New-Object System.Drawing.Size(10,35)
$ReleaseButton1.Text = 'Релиз 20.0.0'
$ReleaseButton1.AutoSize = 'True'
$ReleaseButton1.Font = $FontRelease
$ReleaseButton1.Backcolor = 'Transparent'
$ReleaseButton1.Checked = $False
#Release Button 2
$ReleaseButton2 = New-Object System.Windows.Forms.RadioButton
$ReleaseButton2.Location = New-Object System.Drawing.Size(10,55)
$ReleaseButton2.Text = 'Релиз 21.0.0'
$ReleaseButton2.AutoSize = 'True'
$ReleaseButton2.Font = $FontRelease
$ReleaseButton2.Backcolor = 'Transparent'
$ReleaseButton2.Checked = $False
#Release Button 3
$ReleaseButton3 = New-Object System.Windows.Forms.RadioButton
$ReleaseButton3.Location = New-Object System.Drawing.Size(10,75)
$ReleaseButton3.Text = 'Релиз 22.0.0'
$ReleaseButton3.AutoSize = 'True'
$ReleaseButton3.Font = $FontRelease
$ReleaseButton3.Backcolor = 'Transparent'
$ReleaseButton3.Checked = $False


$ButtonAccept =  New-Object System.Windows.Forms.Button
$ButtonAccept.Location = New-Object System.Drawing.Size(-3,100)
#$ButtonAccept.Size = New-Object System.Drawing.Size(75,23)
$ButtonAccept.Text = "Подтвердить"
$ButtonAccept.Width = $ReleaseForm.Width
$ButtonAccept.Height = '33'

if ($ReleaseButton0.Checked -eq $true) {$Global:RLS = 19} 
if ($ReleaseButton1.Checked -eq $true) {$Global:RLS = 20}
if ($ReleaseButton2.Checked -eq $true) {$Global:RLS = 21}
if ($ReleaseButton3.Checked -eq $true) {$Global:RLS = 22}

function BTN_CLICK()
{
  if ($ReleaseButton0.Checked -eq $true) {$Global:RLS = 19} 
  if ($ReleaseButton1.Checked -eq $true) {$Global:RLS = 20}
  if ($ReleaseButton2.Checked -eq $true) {$Global:RLS = 21}
  if ($ReleaseButton3.Checked -eq $true) {$Global:RLS = 22}
  #
    
}
$ButtonAccept.Add_Click(
{
BTN_CLICK;$ReleaseForm.Close()
})
                        

$ReleaseForm.Controls.Add($ButtonAccept)
$ReleaseForm.Add_Shown({$ReleaseForm.Activate()})
$ReleaseForm.Controls.AddRange(@($ReleaseButton0,$ReleaseButton1,$ReleaseButton2,$ReleaseButton3))
$ReleaseForm.ShowDialog()


}







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
$objForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D

# Configure keyboard intercepts for ESC & ENTER.

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") 
    {
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
$RadioVRX.Checked = $True
$RadioVRX.Text = "VRX"

#RADIO VRQ
$RadioVRQ = $RadioButton2 = New-Object System.Windows.Forms.RadioButton
$RadioVRQ.Location = New-Object System.Drawing.Size(10,35)
$RadioVRQ.Text = "VRQ"
$RadioVRQ.Checked = $False
#
$objForm.Controls.AddRange(@($MyGroupBox))


$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,122)
$objLabel.Font = $Font
$objLabel.AutoSize = 'True'
$objLabel.BackColor = 'Transparent'
$objLabel.Text = "!!!!!!!!!!!!!!"
$objLabel.Visible = 'TRUE'
#$objForm.Controls.Add($objLabel) 

#TEXTBOX
$TextBox = New-Object System.Windows.Forms.TextBox
$TextBox.Location = New-Object System.Drawing.Size(10,30)
$TextBox.Visible = 'False'
$TextBox.Size = '120,80'
$TextBox.MaxLength = '3'


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
#ИНТЕРФЕЙСЫ
$RadioINT = New-Object System.Windows.Forms.RadioButton
$RadioINT.Location = New-Object System.Drawing.Size(10,50)
$RadioINT.Text = "Интерфейс"
#
$objForm.Controls.Add($MyGroupBox2)
$eventMAG = {
             if($RadioMAG.Checked){
             $TextBox.Text = 'Введите магазин'
             $Combo_Srez.Visible = $False
             $TextBox.Visible = $True
             }
             elseif($RadioContur.Checked){
             $Combo_Srez.Visible = $True
             $TextBox.Text = ''
             $TextBox.Visible = $False
             }
             elseif($RadioINT.Checked){
             $Combo_Srez.Visible = $True
             $TextBox.Text = ''
             $TextBox.Visible = $False
             }
            }

$eventBOX = {
             $TextBox.Text = ''
             }
$TextBox.Add_DoubleClick($eventBOX)
$RadioMAG.Add_Click($eventMAG)
$RadioContur.Add_Click($eventMAG)
$RadioINT.Add_Click($eventMAG)





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
            if($RadioVRQ.Checked)
            {
             $Combo_Srez.DataSource = $VRQ}
            elseif ($RadioVRX.Checked){
             $Combo_Srez.DataSource = $VRX}
            else{
             $Combo_Srez.DataSource = $VRX}
$MyGroupBox.Controls.AddRange(@($RadioVRX,$RadioVRQ))
$Combo_Srez.Text
$eventSRED = {
            #$Combo_Srez.Items.Clear()
            if($RadioVRQ.Checked)
            {
             $Combo_Srez.DataSource = $VRQ}
            elseif ($RadioVRX.Checked){
             $Combo_Srez.DataSource = $VRX}
            else{
             $Combo_Srez.DataSource = $VRX}
            }


$objForm.Controls.Add($MyGroupBox3)
$MyGroupBox3.Controls.AddRange(@($Combo_Srez,$TextBox)) 
$RadioVRQ.Add_Click($eventSRED)
$RadioVRX.Add_Click($eventSRED)
$Combo_Srez.add_SelectedIndexChanged($eventSRED)
#$Combo_Srez.Add_Click($eventSRED)
$MyGroupBox2.Controls.AddRange(@($RadioContur,$RadioMAG,$RadioINT)) 


# Create BUTTON FOR START REDEPLOY WILDFLY
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(10,270)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "RESTART WILDFLY"
$OKButton.AutoSize = 'True'



#Обработка ВЫБОРА + RESTART SERVICES
$OKButton.Add_Click(
{
    #Проверка Среды
    if($RadioVRX.Checked)
    {
      $SRED = 'vrx'  
    }
    elseif ($RadioVRQ.Checked)
    {
      $SRED = 'vrq'
    }
    else
    {
     [System.Windows.Forms.MessageBox]::Show("НЕ ВЫБРАН КОНТУР!","ВЫБЕРИТЕ КОНТУР",'OK','ERROR')
     return
    }

    #Проверка контура
    if ($RadioContur.Checked)
    {
     #[System.Windows.Forms.MessageBox]::Show("ajb","Контур",'OK','Info')
     $CONT = "ajb"
     $MACHINE = $Combo_Srez.SelectedItem
    }
    elseif ($RadioMAG.Checked)
    {
     #[System.Windows.Forms.MessageBox]::Show("a","МАГАЗИН")
     $CONT = "a"
     $MACHINE = $TextBox.Text
    }
    elseif ($RadioINT.Checked = $true)
    {
     #[System.Windows.Forms.MessageBox]::Show("int","ИНТЕРФЕЙС")
     $CONT = "int"
     $MACHINE = $Combo_Srez.SelectedItem
    }



    #Проверка ввода.
    if($MACHINE -eq 'Введите магазин' -or $Machine -eq '')
    {
      [System.Windows.Forms.MessageBox]::Show('Обнаружена ошибка выбора станции. Повторите ввод!',"Ошибка выбора",'RetryCancel','ERROR')
      return
    }
    else
    {
     $SERVER = 'fobo-'+ $SRED + "-" + $CONT + $MACHINE
     $Answer = [System.Windows.Forms.MessageBox]::Show("Выбрана машина: " + $SERVER + ". ВЫБОР ВЕРЕН?","Выбор сделан",'YesNo','WARNING')
     switch($Answer){
        "YES"{
               RELEASE_WINDOW
               [System.Windows.Forms.MessageBox]::Show("В Переменной сейчас -  " + $SERVER ,"Выбор сделан",'YesNo','WARNING')
               [System.Windows.Forms.MessageBox]::Show($RLS)
               KillWildfly($SERVER)
             }
        "NO"{ return }
        }
    }
    
    

})

#$OKButton.Add_Click({$x=$objListBox.SelectedItem;$objForm.Close()})
$objForm.Controls.Add($OKButton)

# Cancel EXIT Button
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(450,280)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Закрыть программу"
$CancelButton.AutoSize = 'True'
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