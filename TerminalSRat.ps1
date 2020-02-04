##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjofG5iZk2U3hSWElUcqQhZKo04+w8OvoqBnpQcpHR01v2xrpFE6vFN8TR/wa+eEQRRg4Yt8K8LvfVe6qSsI=
##Kd3HFJGZHWLWoLaVvnQnhQ==
##LM/RF4eFHHGZ7/K1
##K8rLFtDXTiS5
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
##L8/UAdDXTlaDjofG5iZk2U3hSWElUcqQhZKi14qo8PrQnjHLSJRUe1F7mSj4Sk6lXJI=
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
$Global:Version = "V1.1"

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

function Show-Console
{
    # 4 SHOW
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 4)
}

function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}

$Global:ServerKeys = @{
"Новосибирск (mskts1)"="mskts1";
"Челябинск (mskts2)"="mskts2";
"Санкт-Петербург (mskts3)"="mskts3";
"Воронеж (mskts4)"="mskts4";
"Саратов (mskts5)"="mskts5";
"Ростов (mskts6)"="mskts6";
"ЛПТранс (mskts7)"="mskts7";
"СИАМ (msktssiam)"="msktssiam"

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
    if($RadioTerminal1.Checked)
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
                "Yes"{ $Global:SRV = $Global:ServerKeys.$Check
                }
                "No"{$Global:SRV = "";return}
                "Cancel"{$Global:SRV = "";return}
            }
            
        }
    }
    
    }
    elseif($RadioWork2.Checked)
    {
    $Machine = $WorkText.Text
    if($Machine -eq $Null -or $Machine -eq "" -or $Machine -eq "Введите имя рабочей станции")
    {
       $Global:SRV = ""
       [System.Windows.Forms.MessageBox]::Show("Ошибка выбора!")
       return 
    }
    else
    {
    $Answer = [System.Windows.Forms.MessageBox]::Show("Вы выбрали Сервер: "+$Machine,"Подтвердите выбор!","YesNoCancel")
    switch($Answer){
                "Yes"{ $Global:SRV = $Machine
                }
                "No"{$Global:SRV = "";return}
                "Cancel"{$Global:SRV = "";return}
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
$MainForm.Height = '200'
$MainForm.TopMost = $True
    if($Image -eq $null){
        $MainForm.Width = '450'}
    else{
        $MainForm.Width = $Image.Width 
        }
$MainForm.Icon = $Icon

$WorkToolTip = New-Object System.Windows.Forms.ToolTip

$WorkToolTipEvent={
    #display popup help
    #each value is the name of a control on the form. 
     Switch ($this.name) {
        "ComboBox"  {$tip = "Enter the name of a computer"}
        "WorkText" {$tip = "Ввод выполняйте указывая полное имя машины, Например W00-0027 или VOLKOVPC"}
        "WorkLabel" {$tip = "Смехуечки"}
        "ComboLabel" {$tip = "Query Win32_BIOS"}
      }
     $WorkToolTip.SetToolTip($this,$tip)
} #end ShowHelp


$ProgLabel = New-Object System.Windows.Forms.Label
$ProgLabel.Location = ('25,5')
$ProgLabel.AutoSize = $true
$ProgLabel.Text = "ПО Для перезагрузки ETRAN/CTM"
$ProgLabel.Font = $Font
$ProgLabel.BackColor = "Transparent"
#Combo Box Servers

$ComboLabel = New-Object System.Windows.Forms.Label
$ComboLabel.Text = "Список Терминальных Серверов"
$ComboLabel.Location = ('35,50')
$ComboLabel.AutoSize = $true
$ComboLabel.BackColor = "Transparent"
$ComboLabel.Font = $Font
$ComboLabel.Visible = $False

$RadioTerminal2 = New-Object System.Windows.Forms.RadioButton
$RadioTerminal2.Location = ('35,25')
$RadioTerminal2.Text = "Terminals"
$RadioTerminal2.Checked = $False
$RadioTerminal2.BackColor = "Transparent"
$RadioTerminal2.AutoSize = $True


$WorkLabel = New-Object System.Windows.Forms.Label
$WorkLabel.Text = "Выбор рабочей станции"
$WorkLabel.Location = ('35,50')
$WorkLabel.AutoSize = $true
$WorkLabel.BackColor = "Transparent"
$WorkLabel.Font = $Font
$WorkLabel.Name = "WorkLabel"
$WorkLabel.Visible = $True
$WorkLabel.Add_MouseHover($WorkToolTipEvent)




$WorkText = New-Object System.Windows.Forms.TextBox
$WorkText.Location = ('40,70')
$WorkText.Size = ('170,30')
$WorkText.Text = "Введите имя рабочей станции"
$WorkText.BackColor = "Transparent"
$WorkText.Name = "WorkText"
$WorkText.Visible = $True

$WorkText.Add_MouseHover($WorkToolTipEvent)

$WorkText.Add_DoubleClick({
    $WorkText.Text = ""
    })


$RadioWork2 = New-Object System.Windows.Forms.RadioButton
$RadioWork2.Location = ('120,25')
$RadioWork2.Text = "WorkStations"
$RadioWork2.Checked = $True
$RadioWork2.BackColor = "Transparent"
$RadioWork2.AutoSize = $True


$ComboTerminal = New-Object System.Windows.Forms.ComboBox
$ComboTerminal.Location = ('40,70')
$ComboTerminal.Size = ('170,30')
$ComboTerminal.DataSource = $Global:ServerList
$ComboTerminal.DropDownStyle = "DropDownList"
$ComboTerminal.Font = $FontTerminal
$ComboTerminal.SelectionStart = 0
$ComboTerminal.SelectionLength = 0
$ComboTerminal.Visible = $False

$ComboTerminal.Add_MouseHover($WorkToolTipEvent)

$ComboTerminal.Add_SelectedIndexChanged({

    $TerminalButton.Focus()
})

$TerminalButton = New-Object System.Windows.Forms.Button
$TerminalButton.Location = ('20,130')
$TerminalButton.AutoSize = $true
$TerminalButton.Text = "Reload ETRAN"
$TerminalButton.Font = $FontTerminal
$TerminalButton.Focus()


#########################################
####
### Button Events Click
$TerminalButton.Add_Click({
    $Checked = CheckServer
    [System.Windows.Forms.MessageBox]::Show("$Checked")
    if($Checked -eq $null -or $Checked -eq "Введите имя рабочей станции")
    {[System.Windows.Forms.MessageBox]::Show("Операция будет отменена");return}
    else{
    [System.Windows.Forms.MessageBox]::Show("Операция перезагрузки будет выполнена для: "+$Checked)
    KillEtran($Checked)
    }
})


$CTM_Button.Add_Click({
    $Checked = CheckServer
    if($Checked -eq $null -or $Checked -eq "Введите имя рабочей станции")
    {[System.Windows.Forms.MessageBox]::Show("Операция будет отменена");return}
    else{
    [System.Windows.Forms.MessageBox]::Show("Операция перезагрузки будет выполнена для: "+$Checked)
    KillCTM($Checked)
    }
})


#EVENTS

$EventTerminal = {
        if($RadioWork2.Checked)
        {
            $ComboTerminal.Visible = $False
            $ComboTerminal.SelectedItem = ""
            $ComboTerminal.Text = ""
            $ComboLabel.Visible = $False

#Enable Workstation WorkPage
            $WorkLabel.Visible = $True
            $WorkText.Visible = $True
            $WorkText.Text = "Введите имя рабочей станции"

        }
        elseif($RadioTerminal2.Checked)
        {

            $ComboTerminal.Visible = $True
            $ComboLabel.Visible = $True
#Disable Workstation Page
            $WorkLabel.Visible = $False
            $WorkText.Visible = $False
            $WorkText.Text = ""
        }
        }

$RadioTerminal2.Add_Click($EventTerminal)
$RadioWork2.Add_Click($EventTerminal)



$TerminalButton2 = New-Object System.Windows.Forms.Button
$TerminalButton2.Location = ('120,130')
$TerminalButton2.AutoSize = $true
$TerminalButton2.Text = "Reload ALL ETRAN"
$TerminalButton2.Font = $FontTerminal
$TerminalButton.Image = $Image

$CTM_Button2 = New-Object System.Windows.Forms.Button
$CTM_Button2.Location = ('120,105')
$CTM_Button2.AutoSize = $true
$CTM_Button2.Text = "Reload All CTM"
$CTM_Button2.Font = $FontTerminal


$CTM_Button = New-Object System.Windows.Forms.Button
$CTM_Button.Location = ('20,105')
$CTM_Button.AutoSize = $true
$CTM_Button.Text = "Reload CTM"
$CTM_Button.Font = $FontTerminal


$CTM_Button.add_MouseHover({ENTERCOLOR($CTM_Button)})
$CTM_Button.add_MouseLeave({LEAVECOLOR($CTM_Button)})

$CTM_Button2.add_MouseHover({ENTERCOLOR($CTM_Button2)})
$CTM_Button2.add_MouseLeave({LEAVECOLOR($CTM_Button2)})

$TerminalButton.add_MouseHover({ENTERCOLOR($TerminalButton)})
$TerminalButton.add_MouseLeave({LEAVECOLOR($TerminalButton)})

$TerminalButton2.add_MouseHover({ENTERCOLOR($TerminalButton2)})
$TerminalButton2.add_MouseLeave({LEAVECOLOR($TerminalButton2)})

$TerminalButton.Focus()
$MainForm.Controls.AddRange(@($ComboTerminal,$ComboLabel,$ProgLabel,$TerminalButton,$TerminalButton2,$RadioTerminal2,$RadioWork2,$WorkText,$WorkLabel,$CTM_Button2,$CTM_Button))
$MainForm.ShowDialog()

}
Hide-Console
ReloadEtran