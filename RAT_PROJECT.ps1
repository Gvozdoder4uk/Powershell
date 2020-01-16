#Get-ADUser -Filter * -SearchBase "ou=MSK,dc=rusagrotrans,dc=ru" | Where-Object -Like "*fokin*" |Format-Table Name,SamAccountname
#

#
############################################################################
# Programm For IT Infrastructure Config And Control
#
# Made By Fokin L@b 
############################################################################

$Global:Version = "V0.1"

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
$FontBanksy = New-Object System.Drawing.Font("Tempus Sans ITC",8,[System.Drawing.FontStyle]::Regular)

# Initialize Main Form #
$PasswordForm = New-Object System.Windows.Forms.Form 
$PasswordForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$PasswordForm.SizeGripStyle = "Hide"
$PasswordForm.BackgroundImage = $Image
$PasswordForm.BackgroundImageLayout = "None"
$PasswordForm.Text = "Программа администрирования инфраструктуры: $Global:Version"
$PasswordForm.StartPosition = "CenterScreen"
$PasswordForm.Height = '370'
    if($Image -eq $null){
        $PasswordForm.Width = '580'}
    else{
        $PasswordForm.Width = $Image.Width 
        }
$PasswordForm.Icon = $Icon




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



$PasswordForm.Controls.AddRange(@($PasswordsButton))
$PasswordForm.ShowDialog()

}

MainWindow
