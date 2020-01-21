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
$Global:Version = "V1.0"

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
$ComboTerminal.Location = ('20,50')
$ComboTerminal.Size = ('170,30')
$ComboTerminal.DataSource = $Global:ServerList
$ComboTerminal.DropDownStyle = "DropDownList"
$ComboTerminal.Font = $FontTerminal


$TerminalButton = New-Object System.Windows.Forms.Button
$TerminalButton.Location = ('20,80')
$TerminalButton.AutoSize = $true
$TerminalButton.Text = "Reload ETRAN"
$TerminalButton.Font = $Font

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
$TerminalButton2.Font = $Font




$MainForm.Controls.AddRange(@($ComboTerminal,$ComboLabel,$ProgLabel,$TerminalButton,$TerminalButton2))
$MainForm.ShowDialog()

}
Hide-Console
ReloadEtran