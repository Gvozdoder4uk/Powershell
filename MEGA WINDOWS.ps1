##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjofG5iZk2UD9fW4kZcyVhZKo04+w8OvoqBnwarVaQFd49g==
##Kd3HFJGZHWLWoLaVvnQnhQ==
##LM/RF4eFHHGZ7/K1
##K8rLFtDXTiW5
##OsHQCZGeTiiZ4NI=
##OcrLFtDXTiW5
##LM/BD5WYTiiZ4tI=
##McvWDJ+OTiiZ4tI=
##OMvOC56PFnzN8u+Vs1Q=
##M9jHFoeYB2Hc8u+Vs1Q=
##PdrWFpmIG2HcofKIo2QX
##OMfRFJyLFzWE8uK1
##KsfMAp/KUzWJ0g==
##OsfOAYaPHGbQvbyVvnQX
##LNzNAIWJGmPcoKHc7Do3uAuO
##LNzNAIWJGnvYv7eVvnQX
##M9zLA5mED3nfu77Q7TV64AuzAgg=
##NcDWAYKED3nfu77Q7TV64AuzAgg=
##OMvRB4KDHmHQvbyVvnQX
##P8HPFJGEFzWE8tI=
##KNzDAJWHD2fS8u+Vgw==
##P8HSHYKDCX3N8u+Vgw==
##LNzLEpGeC3fMu77Ro2k3hQ==
##L97HB5mLAnfMu77Ro2k3hQ==
##P8HPCZWEGmaZ7/K1
##L8/UAdDXTlaDjofG5iZk2UD9fW4kZcyVhZKi14qo8PrQnQr7ZtogZntbpWf5HE7d
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba

#GLOBAL VARIABLE
$Global:RLS=''
# .Net methods for hiding/showing the console
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

function Show-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 4)
}

function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}

##########################################################################################################################################################################################################
#AUTOINSTALL FOBO START
#
Function DEPLOYFOBO([string]$Machine,[string]$Server){
     
}

Function FOBO_INSTALL([string]$Server)
{
    $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
    $FontJob = New-Object System.Drawing.Font("Colibri",9,[System.Drawing.FontStyle]::Bold)
    $Image =  [system.drawing.image]::FromFile("\\dubovenko\D\SOFT\wallapers\Fobo_INTERFACE.jpg")

    #Create FOBO FORM
    $Fobo_Form = New-Object System.Windows.Forms.Form
    $Fobo_Form.SizeGripStyle = "Hide"
    $Fobo_Form.BackgroundImage = $Image
    $Fobo_Form.BackgroundImageLayout = "None"
    $Fobo_Form.Width = $Image.Width
    $Fobo_Form.Height = $Image.Height
    $Fobo_Form.StartPosition = "CenterScreen"
    $Fobo_Form.Top = $true
    $Fobo_Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $Fobo_Form.Text = "Окно установки FOBO"
    $Fobo_Form.TopMost = 'True'
    $Fobo_Form.Icon = $Icon

    

    $FoboRadio = New-Object System.Windows.Forms.RadioButton
    $FoboRadio.Location = New-Object System.Drawing.Point('10','10')
    $FoboRadio.Text = "Одна станция"
    $FoboRadio.BackColor = 'Transparent'
    $FoboRadio.AutoSize = 'True'
    $FoboRadio.Checked = 'True'

    $FoboRadio1 = New-Object System.Windows.Forms.RadioButton
    $FoboRadio1.Location = New-Object System.Drawing.Point('120','10')
    $FoboRadio1.Text = "Пакетная установка"
    $FoboRadio1.BackColor = 'Transparent'
    $FoboRadio1.AutoSize = 'True'

    $FoboMachine = New-Object System.Windows.Forms.TextBox
    $FoboMachine.Location = New-Object System.Drawing.Point('10','40')
    $FoboMachine.Size = '250,25'

    $FoboMachines = New-Object System.Windows.Forms.ListBox
    $FoboMachines.Location = New-Object System.Drawing.Point('10','40')
    $FoboMachines.Size = '200,150'
    $FoboMachines.Visible = 'False'

#Кнопка запуска процесса одиночного выполнения.
    $FoboProcess = New-Object System.Windows.Forms.Button
    $FoboProcess.Location = New-Object System.Drawing.Point('10','70')
    $FoboProcess.Size = '250,25'
    $FoboProcess.Text = 'ПГНАЛИ'

#Кнопка запуска процесса пакетного выполнения.
    $FoboProcess1 = New-Object System.Windows.Forms.Button
    $FoboProcess1.Location = New-Object System.Drawing.Point('215','70')
    $FoboProcess1.Text = 'ПГНАЛИ'


    $FoboFile = New-Object System.Windows.Forms.Button
    $FoboFile.Location = New-Object System.Drawing.Point('215','40')
    $FoboFile.Text = 'FILE'

    $EventMachine = {
        if($FoboRadio1.Checked)
        {
            $Fobo_Form.Height = $Image.Height
            $FoboProcess1.Visible = $True
            $FoboProcess.Visible = $False
            $FoboFile.Visible = $True
            $FoboMachine.Visible = $False
            $FoboMachine.Text = ''
            $FoboMachines.Visible = $True
        }
        elseif ($FoboRadio.Checked)
        {
            $Fobo_Form.Height = '150'
            $FoboProcess.Visible = $True
            $FoboProcess1.Visible = $False
            $FoboFile.Visible = $False
            $FoboMachines.Visible = $False
            $FoboMachine.SelectedText = ''
            $FoboMachine.Text = ''
            $FoboMachine.Visible = $True
        }
                        
    }


    $EventFile = {
                $FoboMachines.Items.Clear()
                Add-Type -AssemblyName System.Windows.Forms
                $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
                InitialDirectory = 'Desktop'
                Filter = 'Text (*.txt)|*.txt|Все файлы |*.*'
                Title = 'Выберите список машин'}
                $FileBrowser.ShowDialog()
                Get-Content $FileBrowser.FileName | ForEach-Object {$FoboMachines.Items.Add($_)}

                }

#ФУНКЦИЯ УСТАНОВКИ И ПЕРЕУСТАНОВКИ FOBO


#Заготовка под функцию
$eventStartDeploy={
    
    $Machine = $FoboMachine.Text
    $TPath = Test-Path \\$Machine\C$\NTSwincash\config
    $TPath2 = Test-Path \\$Machine\C$\NTSwincash\jbin
    
    if($FoboRadio.Checked)
    {
        if($Machine -eq '' -or $Machine -eq $null)
        {
            $O = [System.Windows.Forms.MessageBox]::Show("Машина для установки FOBO не выбрана!","Ошибка",'OK','ERROR')
            switch($O)
            {
                'OK' {return}
            }
        }
        elseif($Machine -ne '' -or $Machine -ne $null)
        {
            
            if ($TPath -eq $False)
            {
                if($TPath2 -eq $True)
                {
                }
                else
                {
                    New-Item -ItemType Directory -Path \\$Machine\C$\NTSwincash\jbin  
                }
                    New-Item -ItemType Directory  -Path \\$Machine\C$\NTSwincash\config
                    $ACL = ''
                    $FolderAcl = "\\$Machine\C`$\NTSwincash\"
                    $ACL = Get-Acl $FolderAcl
                    $AccessRule =  New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Пользователи","modify","Containerinherit, ObjectInherit","None","Allow")
                    $ACL.SetAccessRule($AccessRule)
                    $ACL | Set-Acl $FolderAcl

                    #Копирование файлов на удаленную машину DBLINK и JBIN
                    Copy-Item '\\$Server\C$\NTSwincash\config\*' -Filter 'dblink_*' -Destination '\\$Machine\C$\NTSwincash\config\'
                    start xcopy "\\dubovenko\D\SOFT\Fobo\jbin" "\\$Machine\C$\NTSWincash\jbin" /S /E /d
                    Start-Sleep -Seconds 10       
                    #Invoke-Command -ComputerName $Machine {cmd.exe "/c start C:\NTSWincash\jbin\InstallDistributor-NT.bat"}
                    $PS = Test-Path C:\Windows\System32\PsExec.exe
                        if($PS -eq $True)
                        {
                            psexec -d \\$machine cmd /c "C:\NTSwincash\jbin\NTSWincash Service Installer.exe" DistributorService /install
                        }
                        else
                        {
                            $PSEXEC = [System.Windows.Forms.MessageBox]::Show("Не установлен PSEXEC!!!","PSexec","YesNoCancel")
                            switch($PSEXEC)
                            {
                            "YES"{[System.Windows.Forms.MessageBox]::Show("Поиск решения корректной установки")}
                                  #xcopy "\\dubovenko\D\SOFT\PSEXEC\" "C:\Windows\System32"}
                            "NO" {return}
                            "CANCEL" {return}
                            }
                            return  
                        }
            
                            if(Get-Service -Name "NTSwincash distributor" -eq $True)
                            {
                                [System.Windows.Forms.MessageBox]::Show("Служба NTSWincash успешно установлена","Успех",'OK','INFO')
                                #psexec -d \\$machine cmd /c 'C:\NTSwincash\jbin\Configurator.exe'
                                #start 'C:\NTSwincash\jbin\Configurator.exe'
                            }
                            else
                            {
                                [System.Windows.Forms.MessageBox]::Show("Служба NTSWincash не была установлена!","Ошибка",'OK','ERROR')
                            }
                    }
                    elseif ($TPath -eq $True)
                    {
                    $Answer = [System.Windows.Forms.MessageBox]::Show("FOBO уже установлен на ПК
Выполнить переустановку?","Ошибка",'YesNoCancel','Warning')
                        switch($Answer)
                        {
                        "YES"{
                                psexec -d \\$machine cmd /c "C:\NTSwincash\jbin\NTSWincash Service Installer.exe" DistributorService /stop
                                Start-Sleep -Seconds 3  
                                Remove-Item -Recurse C:\NTSwincash
                                Start-Sleep -Seconds 3
                                New-Item -ItemType Directory -Path \\$Machine\C$\NTSwincash\jbin
                                New-Item -ItemType Directory  -Path \\$Machine\C$\NTSwincash\config
                                $ACL = ''
                                $FolderAcl = "\\$Machine\C`$\NTSwincash\"
                                $ACL = Get-Acl $FolderAcl
                                $AccessRule =  New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Пользователи","modify","Containerinherit, ObjectInherit","None","Allow")
                                $ACL.SetAccessRule($AccessRule)
                                $ACL | Set-Acl $FolderAcl

                                #Копирование файлов на удаленную машину DBLINK и JBIN
                                Copy-Item '\\$Server\C$\NTSwincash\config\*' -Filter 'dblink_*' -Destination '\\$Machine\C$\NTSwincash\config\'
                                start xcopy "\\dubovenko\D\SOFT\Fobo\jbin" "\\$Machine\C$\NTSWincash\jbin" /S /E /d
                                Start-Sleep -Seconds 10       
                                #Invoke-Command -ComputerName $Machine {cmd.exe "/c start C:\NTSWincash\jbin\InstallDistributor-NT.bat"}
                                $PS = Test-Path C:\Windows\System32\PsExec.exe
                                    if($PS -eq $True)
                                    {
                                        psexec -d \\$machine cmd /c "C:\NTSwincash\jbin\NTSWincash Service Installer.exe" DistributorService /install
                                    }
                                    else
                                    {
                                        $PSEXEC = [System.Windows.Forms.MessageBox]::Show("Не установлен PSEXEC!!!","PSexec","YesNoCancel")
                                        switch($PSEXEC)
                                    {
                                        "YES"{[System.Windows.Forms.MessageBox]::Show("Поиск решения корректной установки")}
                                        #xcopy "\\dubovenko\D\SOFT\PSEXEC\" "C:\Windows\System32"}
                                        "NO" {return}
                                        "CANCEL" {return}
                                    }
                                    return  
                                    }
            
                                    if(Get-Service -Name "NTSwincash distributor" -eq $True)
                                    {
                                        [System.Windows.Forms.MessageBox]::Show("Служба NTSWincash успешно установлена","Успех",'OK','INFO')
                                        #psexec -d \\$machine cmd /c 'C:\NTSwincash\jbin\Configurator.exe'
                                        #start 'C:\NTSwincash\jbin\Configurator.exe'
                                    }
                                    else
                                    {
                                        [System.Windows.Forms.MessageBox]::Show("Служба NTSWincash не была установлена!","Ошибка",'OK','ERROR')
                                    }
                
                               }
                    "NO"{return}
                    "CANCEL"{return}
              }

            } 
        }
        }
    elseif($FoboRadio1.Checked)
    {
        [System.Windows.Forms.MessageBox]::Show("Над Пакетным решением ведется работа!","Ошибка",'OK','WARNING')
         return
    }                    
}


        
               
    $FoboProcess.add_Click($eventStartDeploy)
    $FoboProcess1.Add_Click($eventStartDeploy)
    $FoboFile.add_Click($EventFile)
    $FoboRadio.add_Click($EventMachine)
    $FoboRadio1.add_Click($EventMachine)
    $Fobo_Form.Controls.AddRange(@($FoboRadio,$FoboRadio1,$FoboMachine,$FoboMachines,$FoboFile,$FoboProcess,$FoboProcess1))
    $Fobo_Form.ShowDialog()
}
##AUTOINSTALL FOBO END
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#JOB MANIPULATOR START
Function JOB_WORKER([string]$SERVER){
    $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
    $FontJob = New-Object System.Drawing.Font("Colibri",9,[System.Drawing.FontStyle]::Bold)
    $Imagejob =  [system.drawing.image]::FromFile("\\dubovenko\D\SOFT\wallapers\Worker2.jpg")
    #Create JOB Form
    $JobForm = New-Object System.Windows.Forms.Form
    $JobForm.SizeGripStyle = "Hide"
    $JobForm.BackgroundImage = $Imagejob
    $JobForm.BackgroundImageLayout = "None"
    $JobForm.Width = $Imagejob.Width
    $JobForm.Height = $Imagejob.Height
    $JobForm.StartPosition = "CenterScreen"
    #$JobForm.Top = $true
    $JobForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $JobForm.Text = "Список заданий сервера - $Server"
    $JobForm.TopMost = 'True'
    $JobForm.Icon = $Icon

    $JobForm.KeyPreview = $True
    $JobForm.Add_KeyDown({
    if ($_.KeyCode -eq "Escape") 
    {$JobForm.Close()
    }
    })

    
    
    if($Server -like '*vrq*')
    {[System.Windows.Forms.MessageBox]::Show("Для VRQ среды пока беда!","VRQ")
     return
        }
    elseif($Server -like '*ajb*')
    {[System.Windows.Forms.MessageBox]::Show("Для серверов центра работа сервиса не предусмотрена!","Контуры")
     return
        }
    elseif($Server -like '*int*')
    {#[System.Windows.Forms.MessageBox]::Show("Выбран интерфейсный сервер","Интерфейс")
    $JOBS_OF_SERVER = Get-ScheduledTask -CimSession $Server -TaskName "FOBO*"
        }
    elseif($Server -like '*a*')
    {#[System.Windows.Forms.MessageBox]::Show("Сервер магазина!","Магазин")
     $JOBS_OF_SERVER = Get-ScheduledTask -CimSession $Server -TaskName "JOB*"
        }
    
#Create Listbox
    $JobList = New-Object System.Windows.Forms.ListBox
    $JobList.Location = New-Object System.Drawing.Size(5,10)
    $JobList.Size = '200,260'
    $JobList.ScrollAlwaysVisible = 'False'
    $JobList.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $JobList.DataSource = $JOBS_OF_SERVER.TaskName

#Label Status
    $JobLabel = New-Object System.Windows.Forms.Label
    $JobLabel.Location = New-Object System.Drawing.Size(210,10)
    $JobLabel.Width = '200'
    $JobLabel.Height = '30'
    $JobLabel.ForeColor = 'Green'
    $JobLabel.BackColor = 'Transparent'
    $JobLabel.Text = '    Состояние задачи:'
    

    $JobStatus = New-Object System.Windows.Forms.Textbox
    $JobStatus.Location = New-Object System.Drawing.Size(210,40)
    $JobStatus.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $JobStatus.ReadOnly = 'True'
    $JobStatus.Size = '150,23'
    $JobStatus.Font = $FontJob
    $JobStatus.Text = '  Состояние задачи'

    $JL_SELECT ={
       $JB = $JobList.SelectedItem 
       $JBSTAT = Get-ScheduledTask -CimSession $Server -TaskName $JB
       $jobStatus.text = "            "+$JBSTAT.State
       $JobLabel.Text = '      Cостояние задачи:     ' + $JB     
    }
    $JobList.add_SelectedIndexChanged($JL_SELECT)


#Label Actions:
    $JobLabel1 = New-Object System.Windows.Forms.Label
    $JobLabel1.Location = New-Object System.Drawing.Size(210,70)
    $JobLabel1.Width = '200'
    $JobLabel1.Height = '30'
    $JobLabel1.ForeColor = 'Black'
    $JobLabel1.BackColor = 'Transparent'
    $JobLabel1.Text = '  Действия:'


#Create START Button
    $JobStart = New-Object System.Windows.Forms.RadioButton
    $JobStart.Location =  New-Object System.Drawing.Point(210,90)
    $JobStart.AutoSize = 'True'
    $JobStart.BackColor = 'Transparent'
    $JobStart.Text = 'Start'
    $JobStart.Font = $FontJob
#Create STOP Button
    $JobStop = New-Object System.Windows.Forms.RadioButton
    $JobStop.Location =  New-Object System.Drawing.Point(210,110)
    $JobStop.AutoSize = 'True'
    $JobStop.BackColor = 'Transparent'
    $JobStop.Text = 'End'
    $JobStop.Font = $FontJob
#Create Enable Button
    $JobEnable = New-Object System.Windows.Forms.RadioButton
    $JobEnable.Location =  New-Object System.Drawing.Point(210,130)
    $JobEnable.AutoSize = 'True'
    $JobEnable.BackColor = 'Transparent'
    $JobEnable.Text = 'Enable'
    $JobEnable.Font = $FontJob
#Create Disable Button
    $JobDisable
    $JobDisable= New-Object System.Windows.Forms.RadioButton
    $JobDisable.Location =  New-Object System.Drawing.Point(210,150)
    $JobDisable.AutoSize = 'True'
    $JobDisable.BackColor = 'Transparent'
    $JobDisable.Text = 'Disable'
    $JobDisable.Font = $FontJob
#Create Processing Button
    $JobProcessButton = New-Object  System.Windows.Forms.Button
    $JobProcessButton.Location = New-Object System.Drawing.Size(217,170)
    $JobProcessButton.Text = "Process"

    $JobProcessButton.add_Click({
              
              if($JobStart.Checked){
                $JB = $JobList.SelectedItem
                Start-ScheduledTask -CimSession $Server -TaskName $JB
                Start-Sleep -Seconds 5
                $JBSTAT = Get-ScheduledTask -CimSession $Server -TaskName $JB
                $jobStatus.text = "            "+$JBSTAT.State
                $JobLabel.Text = '      Cостояние задачи:     ' + $JB  
              }
              elseif($JobStop.Checked){
                $JB = $JobList.SelectedItem
                Stop-ScheduledTask -CimSession $Server -TaskName $JB
                Start-Sleep -Seconds 5
                $JBSTAT = Get-ScheduledTask -CimSession $Server -TaskName $JB
                $jobStatus.text = "            "+$JBSTAT.State
                $JobLabel.Text = '      Cостояние задачи:     ' + $JB
              }
              elseif($JobEnable.Checked){
                $JB = $JobList.SelectedItem
                Get-ScheduledTask -CimSession $Server -TaskName $JB | Enable-ScheduledTask
                #Enable-ScheduledTask -CimSession $Server -TaskName $JB
                Start-Sleep -Seconds 5
                $JBSTAT = Get-ScheduledTask -CimSession $Server -TaskName $JB
                $jobStatus.text = "            "+$JBSTAT.State
                $JobLabel.Text = '      Cостояние задачи:     ' + $JB
              }
              elseif($JobDisable.Checked){
                $JB = $JobList.SelectedItem
                Get-ScheduledTask -CimSession $Server -TaskName $JB | Disable-ScheduledTask
                #Disable-ScheduledTask -CimSession $Server -TaskName $JB
                Start-Sleep -Seconds 5
                $JBSTAT = Get-ScheduledTask -CimSession $Server -TaskName $JB
                $jobStatus.text = "            "+$JBSTAT.State
                $JobLabel.Text = '      Cостояние задачи:     ' + $JB
              }
    
    })
    
    $JobForm.Controls.AddRange(@($JobStart,$JobStop,$JobEnable,$JobDisable))
    $JobForm.Controls.AddRange(@($JobList,$JobStatus,$JobLabel,$JobLabel1,$JobProcessButton))
    $JobForm.Add_Shown({$JobForm.Activate()})
    $JobForm.ShowDialog()

}
#JOB MANIPULATOR END
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#CHECK SERVICES FUNCTION START
Function CheckServices([string]$Server)
{
    $FontCheck = New-Object System.Drawing.Font("Colibri",7,[System.Drawing.FontStyle]::Bold)
    $ImageCheck =  [system.drawing.image]::FromFile("\\dubovenko\D\SOFT\wallapers\Services.jpg")
    $FontLabelCheck = New-Object System.Drawing.Font("Colibri",11,[System.Drawing.FontStyle]::Bold)
    $FontStatus = New-Object System.Drawing.Font("Colibri",9,[System.Drawing.FontStyle]::Bold)
    $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")

    #CHECK SERVICES MAIN FORM 
    $CheckForm = New-Object System.Windows.Forms.Form
    $CheckForm.SizeGripStyle = "Hide"
    $CheckForm.BackgroundImage = $ImageCheck
    $CheckForm.BackgroundImageLayout = "None"
    #$CheckForm.Size = New-Object System.Drawing.Size(250,110)
    $CheckForm.Width = $ImageCheck.Width
    $CheckForm.Height = $ImageCheck.Height
    $CheckForm.StartPosition = "CenterScreen"
    $CheckForm.Top = $true
    $CheckForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $CheckForm.Text = "Монитор контроля сервисов $Server"
    $CheckForm.Icon = $Icon

    $CheckForm.KeyPreview = $True
    $CheckForm.Add_KeyDown({
    if ($_.KeyCode -eq "Escape") 
    {$CheckForm.Close()
    }
    })


    Function Checker_Wild([object]$ServiceWildState){
        if($ServiceWildState.Status -eq 'Running')
        {
            $StatusWild.Text = $ServiceWildState.status
            $StatusWild.BackColor = '#90ee90'
        }
        else
        {
            $StatusWild.Text = $ServiceWildState.status
            $StatusWild.BackColor = 'Red'
        }
    }

    Function Checker_NTS([object]$ServiceNTSState){
        if($ServiceNTSState.Status -eq 'Running')
        {
            $StatusNTS.Text = $ServiceNTSState.status
            $StatusNTS.BackColor = '#90ee90'
        }
        else
        {
            $StatusNTS.Text = $ServiceNTSState.status
            $StatusNTS.BackColor = 'Red'
        }
    }



    $ServiceWildState = $Wildfly = Get-Service -Name Wildfly -ComputerName $Server 
    $ServiceNTSState = $NTSwincash = Get-Service -Name "NTSwincash distributor" -ComputerName $Server 

#TEXTBOX WILDFLY
    $CheckLabelWild = New-Object System.Windows.Forms.TextBox
    $CheckLabelWild.Location = New-Object System.Drawing.Size(5,20)
    $CheckLabelWild.Width  = '170'
    $CheckLabelWild.Height = '25'
    $CheckLabelWild.Font = $FontLabelCheck
    $CheckLabelWild.AutoSize = 'True'
    $CheckLabelWild.Text = "  Сервис : Wildfly"
    $CheckLabelWild.ReadOnly = 'True'
    $CheckLabelWild.BackColor = '#90ee90'
    $CheckLabelWild.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    
#TEXTBOX NTS
    $CheckLabelNTS= New-Object System.Windows.Forms.TextBox
    $CheckLabelNTS.Location = New-Object System.Drawing.Size(5,80)
    $CheckLabelNTS.Font = $FontLabelCheck
    $CheckLabelNTS.Width  = '200'
    $CheckLabelNTS.Height = '25'
    $CheckLabelNTS.AutoSize = 'True'
    $CheckLabelNTS.Text = " Сервис : NTSWincash"
    $CheckLabelNTS.ReadOnly = 'True'
    $CheckLabelNTS.BackColor = 'Orange'
    $CheckLabelNTS.SelectionStart = '0'
    $CheckLabelNTS.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D

#Подпись Статуса
    $STLabelW = New-Object System.Windows.Forms.Label
    $STLabelW.Location = New-Object System.Drawing.Point(210,5)
    $STLabelW.AutoSize = 'True'
    $STLabelW.Font = $FontCheck
    $STLabelW.Text = 'Состояние задания:'
    $STLabelW.BackColor = 'Transparent'

#Подпись Статуса 2
    $STLabelN = New-Object System.Windows.Forms.Label
    $STLabelN.Location = New-Object System.Drawing.Point(210,67)
    $STLabelN.AutoSize = 'True'
    $STLabelN.Font = $FontCheck
    $STLabelN.Text = 'Состояние задания:'
    $STLabelN.BackColor = 'Transparent'
            
#Окно текущего статуса WILDFLY
    $StatusWild= New-Object System.Windows.Forms.TextBox
    $StatusWild.Location = New-Object System.Drawing.Size(215,20)
    $StatusWild.Font = $FontStatus
    $StatusWild.Width  = '90'
    $StatusWild.Height = '30'
    $StatusWild.ReadOnly  = 'True'
    $StatusWild.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    Checker_Wild($ServiceWildState)
    
#Окно текущего статуса NTS WINCASH
    $StatusNts = New-Object System.Windows.Forms.TextBox
    $StatusNts.Location = New-Object System.Drawing.Size(215,80)
    $StatusNts.Font = $FontStatus
    $StatusNts.Width  = '90'
    $StatusNts.Height = '30'
    $StatusNts.ReadOnly  = 'True'
    $StatusNts.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    Checker_NTS($ServiceNTSState)

#Create ToolTip
    $ToolTipService = New-Object System.Windows.Forms.ToolTip
    $ToolTipService.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
    $ToolTipService.SetToolTip($StatusWild,"Click For Update Status")
    $ToolTipService.SetToolTip($StatusNts,"Click For Update Status")
    
#Start Service Wildfly
    $StartWildBtn = New-Object System.Windows.Forms.Button
    $StartWildBtn.Location = New-Object System.Drawing.Size(5,46)
    $StartWildBtn.Size = New-Object System.Drawing.Size(50,20)
    $StartWildBtn.Text = "START"
    $StartWildBtn.Font = $FontCheck
    $StartWildBtn.ForeColor = 'green'

#Restart Service Wildfly
    $RestartWildBtn = New-Object System.Windows.Forms.Button
    $RestartWildBtn.Location = New-Object System.Drawing.Size(57,46)
    $RestartWildBtn.Size = New-Object System.Drawing.Size(65,20)
    $RestartWildBtn.Text = "RESTART"
    $RestartWildBtn.Font = $FontCheck
    #$RestartWildBtn.AutoSize = 'True'
    $RestartWildBtn.ForeColor = 'Blue'

#Stop Service Wildfly
    $StopWildBtn = New-Object System.Windows.Forms.Button
    $StopWildBtn.Location = New-Object System.Drawing.Size(125,46)
    $StopWildBtn.Size = New-Object System.Drawing.Size(50,20)
    $StopWildBtn.Text = "STOP"
    $StopWildBtn.Font = $FontCheck
    $StopWildBtn.ForeColor = 'Red'


#EVENT WILDFLY BTN
    $StartWildBtn.add_Click({
        $ServiceWildState =  Get-Service -Name Wildfly -ComputerName $Server | Start-Service
    })
    $RestartWildBtn.add_Click({
        (Get-WmiObject -Class Win32_Process -ComputerName $Server -Filter "name='java.exe'").terminate() | Out-Null
        $ServiceWildState = Get-Service -Name Wildfly -ComputerName $Server | Restart-Service
    })
    $StopWildBtn.add_Click({
        (Get-WmiObject -Class Win32_Process -ComputerName $Server -Filter "name='java.exe'").terminate() | Out-Null
        $ServiceWildState = Get-Service -Name Wildfly -ComputerName $Server | Stop-Service
    })
            
#Start Service NTS
    $StartNTSBtn = New-Object System.Windows.Forms.Button
    $StartNTSBtn.Location = New-Object System.Drawing.Size(5,106)
    $StartNTSBtn.Size = New-Object System.Drawing.Size(50,20)
    $StartNTSBtn.Text = "START"
    $StartNTSBtn.Font = $FontCheck
    $StartNTSBtn.ForeColor = 'Green'

#Restart Service NTS
    $RestartNTSBtn = New-Object System.Windows.Forms.Button
    $RestartNTSBtn.Location = New-Object System.Drawing.Size(58,106)
    $RestartNTSBtn.Size = New-Object System.Drawing.Size(65,20)
    $RestartNTSBtn.Text = "RESTART"
    $RestartNTSBtn.Font = $FontCheck
    $RestartNTSBtn.ForeColor = 'Blue'

#Stop Service NTS
    $StopNTSBtn = New-Object System.Windows.Forms.Button
    $StopNTSBtn.Location = New-Object System.Drawing.Size(125,106)
    $StopNTSBtn.Size = New-Object System.Drawing.Size(50,20)
    $StopNTSBtn.Text = "STOP"
    $StopNTSBtn.Font = $FontCheck
    $StopNTSBtn.ForeColor = 'Red'
    
#EVENT NTS BTN
    $StartNTSBtn.add_Click({
        $ServiceNTSState =  Get-Service -Name "NTSwincash distributor" -ComputerName $Server | Start-Service
    })
    $RestartNTSBtn.add_Click({
        (Get-WmiObject -Class Win32_Process -ComputerName $Server -Filter "name='javaw.exe'").terminate() | Out-Null
        $ServiceNTSState = Get-Service -Name "NTSwincash distributor" -ComputerName $Server | Restart-Service
    })
    $StopNTSBtn.add_Click({
        (Get-WmiObject -Class Win32_Process -ComputerName $Server -Filter "name='javaw.exe'").terminate() | Out-Null
        $ServiceNTSState = Get-Service -Name "NTSwincash distributor" -ComputerName $Server | Stop-Service

    })

#EVENTS FOR UPDATE STATUS
        $StatusWild.add_Click({
        $ServiceWildState = $Wildfly = Get-Service -Name Wildfly -ComputerName $Server 
        Checker_Wild($ServiceWildState)
    })

    $StatusNts.add_Click({
        $ServiceNTSState = $NTSwincash = Get-Service -Name "NTSwincash distributor" -ComputerName $Server 
        Checker_NTS($ServiceNTSState)
    })

    $CheckForm.Controls.AddRange(@($CheckLabelNTS,$CheckLabelWild,$StatusWild,$StatusNts,$STLabelW,$STLabelN))
    $CheckForm.Controls.AddRange(@($StartWildBtn,$RestartWildBtn,$StopWildBtn))
    $CheckForm.Controls.AddRange(@($StartNTSbtn,$RestartNTSBtn,$StopNTSBtn))
    #$CheckForm.Controls.Add($pbrTest)
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
     $Answer = [System.Windows.Forms.MessageBox]::Show("Выбрана машина: " + $SERVER + ".
Подтверждаем выбор?","Выбор сделан",'YesNo','WARNING')     
    }

 return $Answer
}
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#ReDeploy WARNIKA
Function REDEPLOY([string]$Server){
  
  $ChoiceD = [System.Windows.Forms.MessageBox]::Show("YES: Выполнить передеплой существующего сервиса.
NO: Выбрать файл и выполнить передеплой.
Cancel: Выход","Выбор действия!","YesNoCancel")
  switch($ChoiceD)
  {
    "YES" {
            if($Server -like '*int*')
            {
                $DestinationPoint = "\\" + $Server + "\C`$\wildfly\wildfly10\standalone\deployments\"

            #Вызов диалога выбора файла с заданными параметрами
                Add-Type -AssemblyName System.Windows.Forms
                $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
                InitialDirectory = $DestinationPoint
                Filter = 'Deploy (*.war)|*.war;*.failed|Все файлы |*.*'
                Title = 'Выберите файл сервиса для деплоя'}
                $FileBrowser.ShowDialog()

            #Формирование имени файла.
                $DestinationFileName = $DestinationPoint + $FileBrowser.SafeFileName
                $TST = $FileBrowser.SafeFileName

            #Отработка тестов выбора файла и его наличия.
                if($TST -eq ''){ [System.Windows.Forms.MessageBox]::Show("Не выбран файл!");return}

            #Открытие директории и начало выполнения передеплоя файла.
                Invoke-Item $DestinationPoint
                Get-ChildItem -Path  "\\$Server\C`$\wildfly\wildfly10\standalone\deployments\*" -Include "$TST*.backup","$TST*.failed" | Remove-Item
                $DeployedServiceName = "$DestinationFileName.deployed"
                $DeployedServiceName = $DeployedServiceName -replace '\s',''
                Rename-Item "$DestinationFileName.deployed" -NewName "$DestinationFileName.undeploy"
                Start-Sleep 13
                Get-ChildItem -Path "\\$Server\C`$\wildfly\wildfly10\standalone\deployments\*" -Include "$TST.undeploy" | Remove-Item
                Rename-Item "$DestinationFileName.undeployed" -NewName "$DestinationFileName.dodeploy"
                Start-Sleep 5
                if(Get-ChildItem -Path "\\$Server\C`$\wildfly\wildfly10\standalone\deployments\*" -Include "$TST.isdeploying")
                {
                    [System.Windows.Forms.MessageBox]::Show("Переустановка сервиса выполнена успешно")
                }
                else
                {
                    [System.Windows.Forms.MessageBox]::Show("Переустановка сервиса не выполнена!")
                }
            }

            else
            {
            #Формирование пути к серверу и файлу.
                $DestinationPoint = "\\" + $Server + "\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"
            #Вызов диалога выбора файла с заданными параметрами
                Add-Type -AssemblyName System.Windows.Forms
                $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
                InitialDirectory = $DestinationPoint
                Filter = 'Deploy (*.war)|*.war;*.failed|Все файлы |*.*'
                Title = 'Выберите файл сервиса для деплоя'}
                $FileBrowser.ShowDialog()

            #Формирование имени файла.
                $DestinationFileName = $DestinationPoint + $FileBrowser.SafeFileName
                $TST = $FileBrowser.SafeFileName
                if($TST -eq ''){ [System.Windows.Forms.MessageBox]::Show("Не выбран файл!");return}
            #Выполнение передеплоя  
                Invoke-Item "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"
                Get-ChildItem -Path  "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "$TST*.backup","$TST*.deployed","$TST*.failed" | Remove-Item
                Start-Sleep 13
                Get-ChildItem -Path "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "$TST*.undeployed" | Remove-Item
                Start-Sleep 3
                if(Get-ChildItem -Path "\\$Server\C`$\wildfly\wildfly10\standalone\deployments\*" -Include "$TST.isdeploying")
                {
                    [System.Windows.Forms.MessageBox]::Show("Переустановка сервиса выполнена успешно")
                }
                else
                {
                    [System.Windows.Forms.MessageBox]::Show("Переустановка сервиса не выполнена!")
                }

            }
            
    
    }
    "NO"{
            if($Server -like '*int*')
            { [System.Windows.Forms.MessageBox]::Show("Данная функция для Интерфейсных серверов в Разработке!","В РАЗРАБОТКЕ",'OK','ERROR');return}
            Add-Type -AssemblyName System.Windows.Forms
            $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            Filter = 'Deploy (*.war)|*.war|Все файлы |*.*'
            Title = 'Выберите файл сервиса для деплоя'}
            $FileBrowser.ShowDialog()

            $DestinationPoint = "\\" + $Server + "\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"

            $PathTest = Test-Path $DestinationPoint

            
            $DestinationFileName = $DestinationPoint + $FileBrowser.SafeFileName
            $TST = $FileBrowser.SafeFileName
            if($TST -eq ''){ [System.Windows.Forms.MessageBox]::Show("Не выбран файл!");return}
            $FileTestName = $FileBrowser.SafeFileName -split ".war"
            $ProvFile = [System.Windows.Forms.MessageBox]::Show("Будет выполнен деплой файла " + $TST+ ". На сервер :" + $Server,"Путь деплоя","OKCancel",'Info')

            switch($ProvFile)
            {
             "Cancel"{[System.Windows.Forms.MessageBox]::Show("Отмена Деплоя!");return}
            }

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

                Invoke-Item "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"
                Get-ChildItem -Path  "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "$TST*.backup","$TST*.deployed","$TST*.failed" | Remove-Item
                Rename-Item $DestinationFileName -NewName "$DestinationFileName.backup"
                Copy-Item -Path $FileBrowser.FileName -Destination $DestinationFileName
                Start-Sleep 13
                Get-ChildItem -Path "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\*" -Include "$TST*.undeployed" | Remove-Item         
                $Hash1 = Get-FileHash $FileBrowser.FileName
                $Hash2 = Get-FileHash $DestinationFileName

                if ($Hash1.Hash -eq $Hash2.Hash -and $Hash1.Hash -ne $NULL -and $Hash2.Hash -ne $NULL)
                    {
                    [System.Windows.MessageBox]::Show("Файл успешно перенесен $DestinationFileName","Перенос файла успешен")
                    }
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
         #Invoke-Item "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"
    }

    "CANCEL"{
    return}


  }

}
##########################################################################################################################################################################################################

##########################################################################################################################################################################################################
#RESTART WILDFLY DELETE FILES
Function KillWildfly([string]$SRV)
{
    if($Server -like '*int*')
    {
     #[System.Windows.Forms.MessageBox]::Show("ИНТЕРФЕЙС")
     (Get-WmiObject -Class Win32_Process -ComputerName $SRV -Filter "name='java.exe'").terminate() | Out-Null    
     Get-Service -Name Wildfly -ComputerName $SRV -ErrorAction SilentlyContinue | Stop-Service
     Start-Sleep -Seconds 3
     Get-Service -Name Wildfly -ComputerName $server | Start-Service
     Start-Sleep -Seconds 2
     Invoke-Item "\\$Server\C`$\wildfly\wildfly10\standalone\deployments" 
    }
    else{
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
    #Progress
    Get-Service -Name Wildfly -ComputerName $server | Start-Service
    Start-Sleep -Seconds 2
    Invoke-Item "\\$Server\C`$\NTSwincash\jboss\wildfly10\standalone\deployments\"
    }
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
$ImageRelease =  [system.drawing.image]::FromFile("\\dubovenko\D\SOFT\wallapers\RLS.png")
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

#Initialize Release Choice FORM
$ReleaseForm = New-Object System.Windows.Forms.Form
$ReleaseForm.SizeGripStyle = "Hide"
$ReleaseForm.BackgroundImage = $ImageRelease
$ReleaseForm.BackgroundImageLayout = "None"
$ReleaseForm.Size = New-Object System.Drawing.Size(150,165)
$ReleaseForm
$ReleaseForm.StartPosition = "CenterScreen"
$ReleaseForm.Top = $true
$ReleaseForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
$ReleaseForm.Text = "ВЫБЕРИТЕ РЕЛИЗ"
$ReleaseForm.Icon  = $Icon



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

$Image =  [system.drawing.image]::FromFile("\\dubovenko\D\SOFT\wallapers\NTS.jpg")
$Font = New-Object System.Drawing.Font("Times New Roman",8,[System.Drawing.FontStyle]::Bold)
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")


# Initialize Main Form #
$objForm = New-Object System.Windows.Forms.Form 
$objForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$objForm.SizeGripStyle = "Hide"
$objForm.BackgroundImage = $Image
$objForm.BackgroundImageLayout = "None"
$objForm.Text = "Программа для безумного управления сервисами V1.2"
$objForm.StartPosition = "CenterScreen"
$objForm.Height = '370'
$objForm.Width = $Image.Width
$objForm.Icon = $Icon

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
    {
        $objForm.Close()
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
$Combo_Srez.DropDownStyle = 'DropDownList'
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
        "YES"{ if($Server -like '*int*')
               { KillWildfly($SERVER) }
               else{
               RELEASE_WINDOW
               KillWildfly($SERVER)
               }
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


$JobButton = New-Object System.Windows.Forms.Button
$JobButton.Location = New-Object System.Drawing.Size(150,300)
$JobButton.Size = New-Object System.Drawing.Size(75,24)
$JobButton.Font = $Font
$JobButton.Text = "JOB'S"

$JobButton.add_Click({
            
        $Answer = CHECK_SETTINGS
        switch($Answer){
        "YES"{ 
               JOB_WORKER($SERVER)
               $Server = ''
             }
        "NO"{ return }
        }    

})


#FOBO INSTALL
$FoboButton = New-Object System.Windows.Forms.Button
$FoboButton.Location = New-Object System.Drawing.Size(230,300)
$FoboButton.Size = New-Object System.Drawing.Size(120,23)
$FoboButton.Font = $Font
$FoboButton.Text = "Fobo_Install (Beta)"

$FoboButton.Add_Click({

    $Answer = CHECK_SETTINGS
        switch($Answer){
        "YES"{ 
               if($Server -like '*int*' -or $Server -like '*ajb*')
               { [System.Windows.Forms.MessageBox]::Show("Для интерфейсных и центральных серверов процесс установки недоступен!")
                 return
               }
               else{
               FOBO_INSTALL($Server)
               $Server = ''
               }
             }
        "NO"{ return }
        }    
        

})

# Cancel EXIT Button
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(440,300)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Закрыть программу"
$CancelButton.Font = $Font
$CancelButton.AutoSize = 'True'
$CancelButton.Add_Click({$objForm.Close()})



#ADD TO FORM
$objForm.Controls.Add($FoboButton)
$objForm.Controls.Add($CancelButton)
$objForm.Controls.Add($JobButton)
$objForm.Controls.Add($RestartButton)
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
Hide-Console
GENERATOR
$SERVER = ''
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

<#
    $StatusNts.add_paint(
    {if($ServiceNTSState.Status -eq 'Running'){
    $brush = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),(new-object system.drawing.point 0,0),"orange","white")
    $_.graphics.fillrectangle($brush,$this.clientrectangle)}
    
    else{
    $brush = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),(new-object system.drawing.point 0,0),"red","white")
    $_.graphics.fillrectangle($brush,$this.clientrectangle)}

    $brush2 = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point 0,0),(new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),"black","black")
    $_.graphics.drawstring($ServiceNTSState.Status,(new-object System.Drawing.Font("times new roman",11,[System.Drawing.FontStyle]::Bold)),$brush2,(new-object system.drawing.pointf(20,3)))
    })
    #>