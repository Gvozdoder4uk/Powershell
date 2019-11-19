<#$TPath = Test-Path C:\NTSwincash\config
if ($TPath -eq $False)
{
    if(Test-Path C:\NTSwincash\jbin){}
    else
    {
      New-Item -ItemType Directory -Path C:\NTSwincash\jbin  
    }
    New-Item -ItemType Directory  -Path C:\NTSwincash\config
    New-Item -ItemType Directory -Path C:\NTSwincash\jbin
    $ACL = Get-Acl 'C:\NTSwincash\'
    $perm = "BUILTIN\Пользователи","modify,ReadandExecute","Containerinherit, ObjectInherit","None","Allow"
    $AccessRule =  New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Пользователи","modify","Containerinherit, ObjectInherit","None","Allow")
    $ACL.SetAccessRule($AccessRule)
    $ACL | Set-Acl 'C:\NTSwincash'
    xcopy "\\dubovenko\D\SOFT\Fobo\jbin" "C:\NTSWincash\jbin" /S /E /d
    start 'C:\NTSWincash\jbin\InstallDistributor-NT.bat'
    $SRV = Get-Service -Name "NTSwincash distributor"
}
else
{
    if(Test-Path C:\NTSWincash\jbin){}
    else
    {
      New-Item -ItemType Directory -Path C:\NTSwincash\jbin  
    }
    $ACL = Get-Acl 'C:\NTSwincash\'
    $perm = "BUILTIN\Пользователи","modify,ReadandExecute","Containerinherit, ObjectInherit","None","Allow"
    $AccessRule =  New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Пользователи","modify","Containerinherit, ObjectInherit","None","Allow")
    $ACL.SetAccessRule($AccessRule)
    $ACL | Set-Acl 'C:\NTSwincash\'
    xcopy "\\dubovenko\D\SOFT\Fobo\jbin" "C:\NTSWincash\jbin" /S /E /d
    start cmd.exe "/c C:\NTSwincash\jbin\InstallDistributor-NT.bat"
    Start-Sleep -Seconds 2
    #$SRV  = Get-Service -Name "NTSwincash distributor" | Out-Null
    if(Get-Service -Name "NTSwincash distributor")
    {
       [System.Windows.Forms.MessageBox]::Show("Служба NTSWincash успешно установлена","Успех",'OK','INFO')
       start 'C:\NTSwincash\jbin\Configurator.exe'
    }
    else
    {
        [System.Windows.Forms.MessageBox]::Show("Служба NTSWincash не была установлена!","Ошибка",'OK','ERROR')
    }

}

#$ACL.Access




#>


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



    $EventProcess = {
      
      $Machine = $FoboMachine.Text
      $Machine = 'cassa1-000'
      if($Machine -eq '' -or $Machine -eq $null)
      {
        $O = [System.Windows.Forms.MessageBox]::Show("Машина для установки FOBO не выбрана!","Ошибка",'OK','ERROR')
        switch($O){
        'OK' {return}
        }
      }
      else
      {
            $TPath = Test-Path \\$Machine\C$\NTSwincash\config
            $TPath2 = Test-Path \\$Machine\C$\NTSwincash\jbin
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
            xcopy "\\dubovenko\D\SOFT\Fobo\jbin" "\\$Machine\C$\NTSWincash\jbin" /S /E /d         
          
            #Invoke-Command -ComputerName $Machine {cmd.exe "/c start C:\NTSWincash\jbin\InstallDistributor-NT.bat"}
            $PS = Test-Path C:\Windows\System32\PsExec.exe
            if($PS -eq $True)
            {
                psexec -d \\$machine cmd /c "C:\NTSwincash\jbin\NTSWincash Service Installer.exe" DistributorService /install
            }
            else
            {
               [System.Windows.Forms.MessageBox]::Show("Не установлен PSEXEC!!!")
               return  
            }
            
            if(Get-Service -Name "NTSwincash distributor")
            {
                [System.Windows.Forms.MessageBox]::Show("Служба NTSWincash успешно установлена","Успех",'OK','INFO')
                start 'C:\NTSwincash\jbin\Configurator.exe'
            }
            else
            {
                [System.Windows.Forms.MessageBox]::Show("Служба NTSWincash не была установлена!","Ошибка",'OK','ERROR')
            }
            }
            elseif ($TPath -eq $True)
            {
            

            }      
      }
   }
    $FoboProcess.add_Click($EventProcess)
    $FoboFile.add_Click($EventFile)
    $FoboRadio.add_Click($EventMachine)
    $FoboRadio1.add_Click($EventMachine)
    $Fobo_Form.Controls.AddRange(@($FoboRadio,$FoboRadio1,$FoboMachine,$FoboMachines,$FoboFile,$FoboProcess,$FoboProcess1))
    $Fobo_Form.ShowDialog()
}
$Server = '1'
FOBO_INSTALL($Server)