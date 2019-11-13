# Load Windows Forms & Drawing classes.

#GLOBAL VARIABLES
$Global:RLS=''

#


##########################################################################################################################################################################################################
#CHECK SERVICES FUNCTION START

Function CheckServices([string]$Server)
{
    $FontCheck = New-Object System.Drawing.Font("Colibri",9,[System.Drawing.FontStyle]::Bold)
    $ImageCheck =  [system.drawing.image]::FromFile("C:\Users\ks_fokin\Downloads\Services2.jpg")
    $FontLabelCheck = New-Object System.Drawing.Font("Colibri",11,[System.Drawing.FontStyle]::Bold)

    #CHECK SERVICES MAIN FORM 
    $CheckForm = New-Object System.Windows.Forms.Form
    $CheckForm.SizeGripStyle = "Hide"
    $CheckForm.BackgroundImage = $ImageCheck
    $CheckForm.BackgroundImageLayout = "None"
    #$CheckForm.Size = New-Object System.Drawing.Size(250,110)
    $CheckForm.Width = $ImageCheck.Width
    $CheckForm.Height = $ImageCheck.Height
    $CheckForm.StartPosition = "CenterScreen"
    $CheckForm.TopMost = $true
    $CheckForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $CheckForm.Text = "Монитор контроля сервисов $Server"
    
  
    $ServiceWildState = $Wildfly = Get-Service -Name Wildfly -ComputerName $Server -ErrorAction SilentlyContinue
    $ServiceNTSState = $NTSwincash = Get-Service -Name "NTSwincash distributor" -ComputerName $Server -ErrorAction SilentlyContinue

    

      

    #TEST LABEL
    $CheckLabelWild = New-Object System.Windows.Forms.Label
    $CheckLabelWild.Location = New-Object System.Drawing.Size(0,20)
    $CheckLabelWild.Width  = '170'
    $CheckLabelWild.Height = '25'
    $CheckLabelWild.Font = $FontLabelCheck
    #$CheckLabelWild.AutoSize = 'True'
    #$CheckLabelWild.BackColor = 'Transparent'
    $CheckLabelWild.Text = "Сервис : Wildfly"
    $CheckLabelWild.Visible = 'TRUE'
    $CheckLabelWild.ClientRectangle.Width = 3000
    $CheckLabelWild.ClientRectangle.Height = 20

    $CheckLabelWild.add_paint(
    {if($ServiceWildState.Status -eq 'Running'){
    $brush  = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point 0,0),(new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),"green","white")
    $_.graphics.fillrectangle($brush,$this.clientrectangle)}
    else{
    $brush  = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point 0,0),(new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),"red","white")
    $_.graphics.fillrectangle($brush,$this.clientrectangle)
    }
    $brush2 = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point 0,0),(new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),"black","black")
    $_.graphics.drawstring("Сервис : Wildfly",(new-object System.Drawing.Font("times new roman",14,[System.Drawing.FontStyle]::Bold)),$brush2,(new-object system.drawing.pointf(5,0)))
    })
     

    $CheckLabelNTS= New-Object System.Windows.Forms.Label
    $CheckLabelNTS.Location = New-Object System.Drawing.Size(0,80)
    $CheckLabelNTS.Font = $FontLabelCheck
    $CheckLabelNTS.Width  = '210'
    $CheckLabelNTS.Height = '25'
    #$CheckLabelNTS.AutoSize = 'True'
    #$CheckLabelNTS.BackColor = 'Transparent'
    $CheckLabelNTS.Text = "Сервис : NTSWincash"
    $CheckLabelNTS.Visible = 'TRUE'
    $CheckLabelNTS.add_paint(
    {if($ServiceNTSState.Status -eq 'Running'){
    $brush = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point 0,0),(new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),"orange","white")
    $_.graphics.fillrectangle($brush,$this.clientrectangle)}
    else{
    $brush = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point 0,0),(new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),"red","white")
    $_.graphics.fillrectangle($brush,$this.clientrectangle)}

    $brush2 = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point 0,0),(new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),"black","black")
    $_.graphics.drawstring("Сервис : NTSWincash",(new-object System.Drawing.Font("times new roman",14,[System.Drawing.FontStyle]::Bold)),$brush2,(new-object system.drawing.pointf(5,0)))
    })


    #Окно текущего статуса WILDFLY
    $StatusWild= New-Object System.Windows.Forms.Label
    $StatusWild.Location = New-Object System.Drawing.Size(220,20)
    $StatusWild.Font = $FontLabelCheck
    $StatusWild.Width  = '90'
    $StatusWild.Height = '25'
    $StatusWild.BackColor = 'Transparent'
    $StatusWild.Visible = 'TRUE'

    $StatusWild.add_paint(
    {if($ServiceWildState.Status -eq 'Running'){
    $brush = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),(new-object system.drawing.point 0,0),"green","white")
    $_.graphics.fillrectangle($brush,$this.clientrectangle)}
    else{
    $brush = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),(new-object system.drawing.point 0,0),"red","white")
    $_.graphics.fillrectangle($brush,$this.clientrectangle)}

    $brush2 = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point 0,0),(new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),"black","black")
    $_.graphics.drawstring($ServiceWildState.status,(new-object System.Drawing.Font("times new roman",14,[System.Drawing.FontStyle]::Bold)),$brush2,(new-object system.drawing.pointf(5,0)))
    })

    #Окно текущего статуса NTS WINCASH
    $StatusNts =
    $StatusNts = New-Object System.Windows.Forms.Label
    $StatusNts.Location = New-Object System.Drawing.Size(220,80)
    $StatusNts.Font = $FontLabelCheck
    $StatusNts.Width  = '90'
    $StatusNts.Height = '25'
    $StatusNts.BackColor = 'Transparent'
    $StatusNts.Visible = 'TRUE'

    #Create ToolTip
    $ToolTipService = New-Object System.Windows.Forms.ToolTip
    $ToolTipService.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
    #$ToolTipService.IsBalloon = $true
    $ToolTipService.SetToolTip($StatusWild,"НАЖМИ ДЛЯ ОБНОВЛЕНИЯ СТАТУСА")
    $ToolTipService.SetToolTip($StatusNts,"НАЖМИ ДЛЯ ОБНОВЛЕНИЯ СТАТУСА")

    $StatusWild.add_Click({
        $ServiceWildState = $Wildfly = Get-Service -Name Wildfly -ComputerName $Server -ErrorAction SilentlyContinue
        $StatusWild.Text = $ServiceWildState
    })


    $StatusNts.add_Click({
        $ServiceNTSState = $NTSwincash = Get-Service -Name "NTSwincash distributor" -ComputerName $Server -ErrorAction SilentlyContinue
        $StatusNts.Text = $ServiceNTSState
    })

    $StatusNts.add_paint(
    {if($ServiceNTSState.Status -eq 'Running'){
    $brush = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),(new-object system.drawing.point 0,0),"orange","white")
    $_.graphics.fillrectangle($brush,$this.clientrectangle)}
    
    else{
    $brush = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),(new-object system.drawing.point 0,0),"red","white")
    $_.graphics.fillrectangle($brush,$this.clientrectangle)}

    $brush2 = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point 0,0),(new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),"black","black")
    $_.graphics.drawstring($ServiceNTSState.Status,(new-object System.Drawing.Font("times new roman",14,[System.Drawing.FontStyle]::Bold)),$brush2,(new-object system.drawing.pointf(5,0)))
    })


    #Start Service Wildfly
    $StartWildBtn = New-Object System.Windows.Forms.Button
    $StartWildBtn.Location = New-Object System.Drawing.Size(5,46)
    $StartWildBtn.Size = New-Object System.Drawing.Size(35,20)
    $StartWildBtn.Text = "START"
    $StartWildBtn.Font = $FontCheck
    $StartWildBtn.AutoSize = 'True'
    $StartWildBtn.ForeColor = 'green'
    #Restart Service Wildfly
    $RestartWildBtn = New-Object System.Windows.Forms.Button
    $RestartWildBtn.Location = New-Object System.Drawing.Size(60,46)
    $RestartWildBtn.Size = New-Object System.Drawing.Size(35,20)
    $RestartWildBtn.Text = "RESTART"
    $RestartWildBtn.Font = $FontCheck
    $RestartWildBtn.AutoSize = 'True'
    $RestartWildBtn.ForeColor = 'Blue'
    #Stop Service Wildfly
    $StopWildBtn = New-Object System.Windows.Forms.Button
    $StopWildBtn.Location = New-Object System.Drawing.Size(133,46)
    $StopWildBtn.Size = New-Object System.Drawing.Size(75,23)
    $StopWildBtn.Text = "STOP"
    $StopWildBtn.Font = $FontCheck
    $StopWildBtn.AutoSize = 'True'
    $StopWildBtn.ForeColor = 'Red'


    #EVENT WILDFLY BTN
    $StartWildBtn.add_Click({
        Get-Service -Name Wildfly -ComputerName $Server | Start-Service
    })
    $RestartWildBtn.add_Click({
        Get-Service -Name Wildfly -ComputerName $Server | Restart-Service
    })
    $StopWildBtn.add_Click({
        Get-Service -Name Wildfly -ComputerName $Server | Stop-Service
    })
    
    
    

    #Start Service NTS
    $StartNTSBtn = New-Object System.Windows.Forms.Button
    $StartNTSBtn.Location = New-Object System.Drawing.Size(5,106)
    $StartNTSBtn.Size = New-Object System.Drawing.Size(35,20)
    $StartNTSBtn.Text = "START"
    $StartNTSBtn.Font = $FontCheck
    $StartNTSBtn.AutoSize = 'True'
    $StartNTSBtn.ForeColor = 'Green'
    #Restart Service NTS
    $RestartNTSBtn = New-Object System.Windows.Forms.Button
    $RestartNTSBtn.Location = New-Object System.Drawing.Size(60,106)
    $RestartNTSBtn.Size = New-Object System.Drawing.Size(35,20)
    $RestartNTSBtn.Text = "RESTART"
    $RestartNTSBtn.Font = $FontCheck
    $RestartNTSBtn.AutoSize = 'True'
    $RestartNTSBtn.ForeColor = 'Blue'
    #Stop Service NTS
    $StopNTSBtn = New-Object System.Windows.Forms.Button
    $StopNTSBtn.Location = New-Object System.Drawing.Size(133,106)
    $StopNTSBtn.Size = New-Object System.Drawing.Size(75,23)
    $StopNTSBtn.Text = "STOP"
    $StopNTSBtn.Font = $FontCheck
    $StopNTSBtn.AutoSize = 'True'
    $StopNTSBtn.ForeColor = 'Red'

    #EVENT NTS BTN
    $StartNTSBtn.add_Click({
        Get-Service -Name "NTSwincash distributor" -ComputerName $Server | Start-Service
    })
    $RestartNTSBtn.add_Click({
        Get-Service -Name "NTSwincash distributor" -ComputerName $Server | Restart-Service
    })
    $StopNTSBtn.add_Click({
        Get-Service -Name "NTSwincash distributor" -ComputerName $Server | Stop-Service
    })

    $CheckForm.Controls.AddRange(@($CheckLabelNTS,$CheckLabelWild,$StatusWild,$StatusNts))
    $CheckForm.Controls.AddRange(@($StartWildBtn,$RestartWildBtn,$StopWildBtn))
    $CheckForm.Controls.AddRange(@($StartNTSbtn,$RestartNTSBtn,$StopNTSBtn))
    $CheckForm.Controls.Add($pbrTest)
    $CheckForm.Add_Shown({$CheckForm.Activate()})
    #$CheckForm.Controls.AddRange(@($ReleaseButton0,$ReleaseButton1,$ReleaseButton2,$ReleaseButton3))
    #$CheckForm.ShowDialog()
    $CheckForm.Add_Shown({$CheckForm.Activate()})
    $CheckForm.ShowDialog()
    $ServiceNTSState = ''
    $ServiceWildState = ''
}

#CHECK SERVICES FUNCTION END
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
Function CHECK_SETTINGS(){
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
     $Global:SERVER = 'fobo-'+ $SRED + "-" + $CONT + $MACHINE
     $Answer = [System.Windows.Forms.MessageBox]::Show("Выбрана машина: " + $SERVER + ". ВЫБОР ВЕРЕН?","Выбор сделан",'YesNo','WARNING')     
    }

 return $Answer
}
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#ReDeploy WARNIKA
Function REDEPLOY([string]$Server){

  Add-Type -AssemblyName System.Windows.Forms
  $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
  Filter = 'Deploy (*.war)|*.war|Все файлы |*.*'
  Title = 'Выберите файл сервиса для деплоя'}
  $FileBrowser.ShowDialog()

  
  $DestinationPoint = "\\" + $Server + "\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"
  $PathTest = Test-Path $DestinationPoint
  #$DestinationPoint = 'C:\1\deployments\'
  $DestinationFileName = $DestinationPoint + $FileBrowser.SafeFileName
  $TST = $FileBrowser.SafeFileName
  if($TST -eq ''){ [System.Windows.Forms.MessageBox]::Show("Не выбран файл!");return}
  $FileTestName = $FileBrowser.SafeFileName -split ".war"
  $ProvFile = [System.Windows.Forms.MessageBox]::Show("Будет выполнен деплой файла " + $TST+ ". На сервер :" + $Server,"Путь деплоя","OKCancel",'Info')
  switch($ProvFile)
  {
    "Cancel"{[System.Windows.Forms.MessageBox]::Show("Отмена Деплоя!");return}
  }
  #Invoke-Item "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"
  #Compare-Object -ReferenceObject $Hash1 -DifferenceObject $Hash2

  $FileTest = Test-Path $DestinationFileName
  

  if($PathTest -eq $False)
  {
   [System.Windows.MessageBox]::Show("Путь не существует, или недоступен,проверьте выбор сервера")
   return
  }
  elseif($FileTest -eq $False)
  {
    [System.Windows.MessageBox]::Show("Сервиса ранее не было на сервере, проверьте выбор файла")
    return
  }
  elseif($PathTest -eq $True -and $FileTest -eq $True)
  {

    Get-ChildItem -Path  "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "$TST*.backup","$TST*.deployed","$TST*.failed" | Remove-Item
    Rename-Item $DestinationFileName -NewName "$DestinationFileName.backup"
    Copy-Item -Path $FileBrowser.FileName -Destination $DestinationFileName
    Start-Sleep 20
    Get-ChildItem -Path "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "$TST*.undeployed" | Remove-Item         
    $Hash1 = Get-FileHash $FileBrowser.FileName
    $Hash2 = Get-FileHash $DestinationFileName
    #[System.Windows.MessageBox]::Show($Hash1.Hash)
    #[System.Windows.MessageBox]::Show($Hash2.Hash)
        if ($Hash1.Hash -eq $Hash2.Hash -and $Hash1.Hash -ne $NULL -and $Hash2.Hash -ne $NULL)
        {[System.Windows.MessageBox]::Show("Файл успешно перенесен $DestinationFileName","Перенос файла успешен")}
        else
        {
        $CHECKHASH = [System.Windows.MessageBox]::Show("Файл перенесен в $DestinationFileName с ошибками
        Будет выполнено восстановление файла!","Перенос файла провалился","OK",'ERROR')
        Switch($CHECKHASH){
        "OK"{
        Rename-Item $DestinationFileName -NewName "$DestinationFileName.backup.FAILED"
        Rename-Item "$DestinationFileName.backup" -NewName "$DestinationFileName"}
        }
        
  }
  }

}
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#RESTART WILDFLY DELETE FILES
Function KillWildfly([string]$SRV)
{
    
    #Get-Process -Name java -ComputerName $SRV -ErrorAction SilentlyContinue | Format-List
    (Get-WmiObject -Class Win32_Process -ComputerName $SRV -Filter "name='java.exe'").terminate() | Out-Null    
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
    Start-Sleep -Seconds 3
    Progress
    Get-Service -Name Wildfly -ComputerName $server | Start-Service
    Start-Sleep -Seconds 2
    Invoke-Item "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"
}
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#PROGRESS BAR FUNCTION START
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
$pbrTest.Name = 'Выполнение перезапуска службы'

Function StartProgressBar{
   $i = 0
        While ($i -le 250) {
        $pbrTest.Value = $i
        Start-Sleep -m 30
        "VALLUE EQ"
        $i
        $i += 1
        $ProcessForm.Refresh()
    }
    $ProcessForm.Close()     
}
$pbrTest.Add_MouseEnter({StartProgressBar})

$ProcessForm.Controls.Add($pbrTest)
$ProcessForm.Add_Shown({$ProcessForm.Activate()})
$ProcessForm.Controls.AddRange(@($ReleaseButton0,$ReleaseButton1,$ReleaseButton2,$ReleaseButton3))
$ProcessForm.ShowDialog()
$ProcessForm.Focused
$ProcessForm.Refresh()

}
#RELEASE CHOICE FUNCTION END
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#RELEASE CHOICE FUNCTION START 
Function RELEASE_WINDOW(){

$FontRelease = New-Object System.Drawing.Font("Colibri",10,[System.Drawing.FontStyle]::Bold)
$ImageRelease =  [system.drawing.image]::FromFile("C:\Users\ks_fokin\Downloads\RLS2.png")

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

#Initialize Release Choice FORM
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
#Burron Accept
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
#RELEASE CHOICE FUNCTION END
##########################################################################################################################################################################################################


##########################################################################################################################################################################################################
##########################################################################################################################################################################################################
#MAIN FUNCTION FOR ALL PROGRAMM
#CONTAINS MAIN FORM AND BUTTONS FOR START ALL UPPER FUNCTIONS
#
function GENERATOR{
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$VRX = ('1','2','3','4','5','6')
$VRQ = ('1','2','3','4','5','6','7')
# Create base form.

$Image =  [system.drawing.image]::FromFile("C:\Users\ks_fokin\Downloads\NTS.jpg")
$Font = New-Object System.Drawing.Font("Times New Roman",8,[System.Drawing.FontStyle]::Bold)

# Initialize Main Form #
$objForm = New-Object System.Windows.Forms.Form 
$objForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$objForm.SizeGripStyle = "Hide"
$objForm.BackgroundImage = $Image
$objForm.BackgroundImageLayout = "None"
$objForm.Text = "Программа для безумного управления сервисами"
$objForm.StartPosition = "CenterScreen"
$objForm.Height = '380'
$objForm.Width = $Image.Width


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
# ADD GROUP BOX ON FORM
$objForm.Controls.AddRange(@($MyGroupBox))

#TEST LABEL
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
#ADD SECOND GROUP BOX
$objForm.Controls.Add($MyGroupBox2)
#EVENT FOR CHECK GROUP BOX 2 RADIO BUTTONS
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
#EVENT FOR TEXT BOX MAGAZINES
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
$RestartButton = New-Object System.Windows.Forms.Button
$RestartButton.Location = New-Object System.Drawing.Size(5,270)
$RestartButton.Size = New-Object System.Drawing.Size(75,23)
$RestartButton.Text = "RESTART WILDFLY"
$RestartButton.Font = $Font
$RestartButton.AutoSize = 'True'



#Обработка ВЫБОРА + RESTART SERVICES
$RestartButton.Add_Click(
{
    $Answer = CHECK_SETTINGS
    switch($Answer){
        "YES"{
               RELEASE_WINDOW
               KillWildfly($SERVER)
               $Server = ''
               $RLS = ''
             }
        "NO"{ return }
        }    
})


$DeployWAR = New-Object System.Windows.Forms.Button
$DeployWAR.Location = New-Object System.Drawing.Size(130,270)
$DeployWAR.Size = New-Object System.Drawing.Size(75,23)
$DeployWAR.Text = "DEPLOY "".WAR"""
$DeployWAR.Font = $Font
$DeployWAR.AutoSize = 'True'

$DeployWAR.Add_Click({
    $Answer = CHECK_SETTINGS
    switch($Answer){
        "YES"{
               #RELEASE_WINDOW
               REDEPLOY($SERVER)
               $Server = ''
             }
        "NO"{ return }
        }    

})

$CheckServicesBTN = New-Object System.Windows.Forms.Button
$CheckServicesBTN.Location = New-Object System.Drawing.Size(5,300)
$CheckServicesBTN.Size = New-Object System.Drawing.Size(75,23)
$CheckServicesBTN.Text = "ПРОВЕРКА СЕРВИСОВ"
$CheckServicesBTN.Font = $Font
$CheckServicesBTN.AutoSize = 'True'

$CheckServicesBTN.add_Click({
    $Answer = CHECK_SETTINGS
    switch($Answer){
        "YES"{
               CheckServices($SERVER)
               $Server = ''
             }
        "NO"{ return }
        }    
            
})


#$RestartButton.Add_Click({$x=$objListBox.SelectedItem;$objForm.Close()})
$objForm.Controls.Add($RestartButton)

# Cancel EXIT Button
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(450,300)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Закрыть программу"
$CancelButton.Font = $Font
$CancelButton.AutoSize = 'True'
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton)






#TABLO
$objListBox = New-Object System.Windows.Forms.ListBox 
$objListBox.Location = New-Object System.Drawing.Size(250,320) 
$objListBox.Size = New-Object System.Drawing.Size(100,20)
$objListBox.Height = 90


$objForm.Controls.Add($CheckServicesBTN)
$objForm.Controls.Add($DeployWAR)
$objForm.TopMost = $true
$objForm.Add_Shown({$objForm.Activate()})
$objForm.ShowDialog()
}
##########################################################################################################################################################################################################
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#START MAIN PROGRAMM
#
#EXECUTION MAIN PROGRAMM
GENERATOR
##########################################################################################################################################################################################################



##########################################################################################################################################################################################################
#SOME TEST SCURB
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



$Server = "C:\1\"
  #
  $DestinationPoint = $Server 
  $DestinationPoint +=$FileBrowser.SafeFileName
  [System.Windows.Forms.MessageBox]::Show($DEST)
  Rename-Item $DestinationPoint -NewName "$Dest.backup" 
  Copy-Item -Path $FileBrowser.FileName -Destination $DestinationPoint
  $Hash1 = Get-FileHash $FileBrowser.FileName
  $Hash2 = Get-FileHash $DestinationPoint
  #
#>