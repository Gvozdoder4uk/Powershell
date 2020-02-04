#Get-ADUser -Filter * -SearchBase "ou=MSK,dc=rusagrotrans,dc=ru" | Where-Object -Like "*fokin*" |Format-Table Name,SamAccountname
#

#
############################################################################
# Programm For IT Infrastructure Config And Control
#
# Made By Fokin L@b 
############################################################################

$Global:Version = "V0.1"





$Global:ServerKeys = @{
"Новосибирск (mskts1)"="mskts1";
"Челябинск (mskts2)"="mskts2";
"Санкт-Петербург (mskts3)"="mskts3";
"Воронеж (mskts4)"="mskts4";
"Саратов (mskts5)"="mskts5";
"Ростов (mskts6)"="mskts6"
}

$Global:ServerList = 
"Новосибирск (mskts1)",
"Челябинск (mskts2)",
"Санкт-Петербург (mskts3)",
"Воронеж (mskts4)",
"Саратов (mskts5)",
"Ростов (mskts6)"

Function KillEtran([string]$Server)
{
    (Get-WmiObject -Class Win32_Process -ComputerName $Server -Filter "name='iexplore.exe'").terminate() | Out-Null
    (Get-WmiObject -Class Win32_Process -ComputerName $Server -Filter "name='EtranShell.exe'").terminate() | Out-Null
    Get-Process -ComputerName $Server -Name "iexplore*" | Stop-Process -ErrorAction SilentlyContinue
    Get-Process -ComputerName $Server -Name "Etran*" | Stop-Process -ErrorAction SilentlyContinue
}

#Проверка Выбора сервера
Function CheckServer()
{
    $Machine = $ComboTerminal.SelectedItem
    foreach ($Check in $Global:ServerKeys.Keys)
    {
        #[System.Windows.Forms.MessageBox]::Show($Check)
        
        if($Machine -eq $Check)
        {
            $Global:SRV = $Global:ServerKeys.$Check
            $Answer = [System.Windows.Forms.MessageBox]::Show("Вы выбрали Сервер: "+$SRV,"Подтвердите выбор!","YesNoCancel")
            switch($Answer){
                "Yes"{
                }
                "No"{return}
                "Cancel"{return}
            }
            
        }
    }
    return $SRV
    
}

#Soft For Reload Etran On Terminals
Function ReloadEtran()
{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    
function ENTERCOLOR($ELEMENT){
		$ELEMENT.BackColor = 'LightGreen'
	}
function LEAVECOLOR($ELEMENT){
        $ELEMENT.BackColor = 'Control'
    }


$Image =  [system.drawing.image]::FromFile("\\mskdc7\DFS\IT\Install\Distributes\Wallapers\EtranLoader.jpg")
$Font = New-Object System.Drawing.Font("Comic Sans MS",8,[System.Drawing.FontStyle]::Bold)
$FontTerminal = New-Object System.Drawing.Font("Cambria",9,[System.Drawing.FontStyle]::Regular)
$FontTerminalBold = New-Object System.Drawing.Font("Cambria",9,[System.Drawing.FontStyle]::Bold)
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$FontBanksy = New-Object System.Drawing.Font("Tempus Sans ITC",8,[System.Drawing.FontStyle]::Regular)

# Initialize Main Form #
$MainForm = New-Object System.Windows.Forms.Form 
$MainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$MainForm.SizeGripStyle = "Hide"
$MainForm.BackgroundImage = $Image
$MainForm.BackgroundImageLayout = "None"
$MainForm.Text = "Программа восстановления работы Etran: $Global:Version"
$MainForm.StartPosition = "CenterScreen"
$MainForm.Height = '160'
$MainForm.TopMost = $True
    if($Image -eq $null){
        $MainForm.Width = '350'}
    else{
        $MainForm.Width = $Image.Width 
        }
$MainForm.Icon = $Icon

$ProgLabel = New-Object System.Windows.Forms.Label
$ProgLabel.Location = ('40,0')
$ProgLabel.AutoSize = $true
$ProgLabel.Text = "ПО Для перезагрузки ETRAN"
$ProgLabel.Font = $FontTerminalBold
$ProgLabel.BackColor = "Transparent"
#Combo Box Servers

$ComboLabel = New-Object System.Windows.Forms.Label
$ComboLabel.Text = "Список Терминальных Серверов:"
$ComboLabel.Location = ('20,30')
$ComboLabel.AutoSize = $true
$ComboLabel.BackColor = "Transparent"
$ComboLabel.Font = $FontTerminalBold


$ComboTerminal = New-Object System.Windows.Forms.ComboBox
$ComboTerminal.Location = ('40,50')
$ComboTerminal.Size = ('170,30')
$ComboTerminal.DataSource = $Global:ServerList
$ComboTerminal.DropDownStyle = "DropDownList"
$ComboTerminal.Font = $FontTerminal


$TerminalButton = New-Object System.Windows.Forms.Button
$TerminalButton.Location = ('20,80')
$TerminalButton.AutoSize = $true
$TerminalButton.Text = "Reload ETRAN"
$TerminalButton.Font = $FontTerminal

$TerminalButton.Add_Click({
    $Checked = CheckServer
    if($Checked -eq $null)
    {[System.Windows.Forms.MessageBox]::Show("Операция будет отменена");return}
    else{
    [System.Windows.Forms.MessageBox]::Show("Операция перезагрузки будет выполнена для: "+$Checked)
    KillEtran($Checked)
    }
})

$TerminalButton2 = New-Object System.Windows.Forms.Button
$TerminalButton2.Location = ('120,80')
$TerminalButton2.AutoSize = $true
$TerminalButton2.Text = "Reload ALL ETRAN"
$TerminalButton2.Font = $FontTerminal

$TerminalButton.add_MouseHover({ENTERCOLOR($TerminalButton)})
$TerminalButton.add_MouseLeave({LEAVECOLOR($TerminalButton)})

$TerminalButton2.add_MouseHover({ENTERCOLOR($TerminalButton2)})
$TerminalButton2.add_MouseLeave({LEAVECOLOR($TerminalButton2)})


$MainForm.Controls.AddRange(@($ComboTerminal,$ComboLabel,$ProgLabel,$TerminalButton,$TerminalButton2))
$MainForm.ShowDialog()

}











Function WeakPassWords()
{
 # Form For Check and Out Weak Users Passwords:
    $Image =  [system.drawing.image]::FromFile("C:\Wallapers\background.jpg")
    $Font = New-Object System.Drawing.Font("Comic Sans MS",8,[System.Drawing.FontStyle]::Bold)
    $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
    $FontBanksy = New-Object System.Drawing.Font("Tempus Sans ITC",8,[System.Drawing.FontStyle]::Regular)

    # Initialize Main Form #
    $PasswordForm = New-Object System.Windows.Forms.Form
    $PasswordForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $PasswordForm.SizeGripStyle = "Hide"
    $PasswordForm.BackgroundImage = $Image
    $PasswordForm.BackgroundImageLayout = "None"
    $PasswordForm.Text = "Поиск слабых паролей $Global:Version"
    $PasswordForm.StartPosition = "CenterScreen"
    $PasswordForm.Height = '270'
    $PasswordForm.Width = '200'


    $PasswordForm.ShowDialog()
    
}













# MAIN PROGRAM BODY
# Start Function For MainWindow
Function MainWindow()
{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    
function ENTERCOLOR($ELEMENT){
		$ELEMENT.BackColor = 'LightGreen'
	}
function LEAVECOLOR($ELEMENT){
        $ELEMENT.BackColor = 'Control'
    }



$Image =  [system.drawing.image]::FromFile("C:\Wallapers\background.jpg")
$Font = New-Object System.Drawing.Font("Comic Sans MS",8,[System.Drawing.FontStyle]::Bold)
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$FontBanksy = New-Object System.Drawing.Font("Tempus Sans ITC",10,[System.Drawing.FontStyle]::Regular)

# Initialize Main Form #
$MainForm = New-Object System.Windows.Forms.Form 
$MainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$MainForm.SizeGripStyle = "Hide"
$MainForm.BackgroundImage = $Image
$MainForm.BackgroundImageLayout = "None"
$MainForm.Text = "Программа администрирования инфраструктуры: $Global:Version"
$MainForm.StartPosition = "CenterScreen"
$MainForm.Height = '370'
    if($Image -eq $null){
        $MainForm.Width = '580'}
    else{
        $MainForm.Width = $Image.Width 
        }
$MainForm.Icon = $Icon

$FOKINLAB = New-Object System.Windows.Forms.Label
$FOKINLAB.Location = ('450,310')
$FOKINLAB.Text = "Created By Fokin"
$FOKINLAB.Font = $FontBanksy
#$FOKINLAB.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$FOKINLAB.BackColor =  'Transparent'
$FOKINLAB.AutoSize = $True


#Button For open Form For Check Weak Passwords
$PasswordsButton = New-Object System.Windows.Forms.Button
$PasswordsButton.Location = ('10,290')
$PasswordsButton.Size = ('120,25')
$PasswordsButton.Text = "Weak Passwords"
#
# Password Button Click Event
$PasswordsButton.Add_Click({
    WeakPassWords
})

$ReloadEtranButton = New-Object System.Windows.Forms.Button
$ReloadEtranButton.Location = ('130,290')
$ReloadEtranButton.Size = ('100,25')
$ReloadEtranButton.Text = "Reload Etran"

$ReloadEtranButton.Add_Click({
    ReloadEtran
})


$PasswordsButton.add_MouseHover({ENTERCOLOR($PasswordsButton)})
$PasswordsButton.add_MouseLeave({LEAVECOLOR($PasswordsButton)})

$ReloadEtranButton.add_MouseHover({ENTERCOLOR($ReloadEtranButton)})
$ReloadEtranButton.add_MouseLeave({LEAVECOLOR($ReloadEtranButton)})



$MainForm.Controls.AddRange(@($PasswordsButton,$ReloadEtranButton,$FOKINLAB))
$MainForm.ShowDialog()

}

MainWindow
