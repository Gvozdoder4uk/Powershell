##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjofG5iZk2U3hSWElUcqQhZKo04+w8OvoqBnDR5sfBGZWomf1B0Td
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
##L8/UAdDXTlaDjofG5iZk2U3hSWElUcqQhZKi14qo8PrQnj3aTJZUXVs3pgfbSk6lXJI=
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
# Soft for convert Excel to PNG or JPG
$image = $Null



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
Function General(){
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

# Create base form.

function ENTERCOLOR($ELEMENT){
		$ELEMENT.BackColor = 'LightGreen'
	}
function LEAVECOLOR($ELEMENT){
        $ELEMENT.BackColor = 'Control'
    }

$Formats = ("PNG",
            "JPEG")
$Numbers = (1,
            2,
            3,
            4,
            5,
            6)
$Image = [system.drawing.image]::FromFile("C:\Test\Format.png")
$Font = New-Object System.Drawing.Font("Comic Sans MS",8,[System.Drawing.FontStyle]::Bold)
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$FontBanksy = New-Object System.Drawing.Font("Tempus Sans ITC",8,[System.Drawing.FontStyle]::Regular)

# Initialize Main Form #
$objForm = New-Object System.Windows.Forms.Form 
$objForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$objForm.SizeGripStyle = "Hide"
$objForm.BackgroundImage = $Image
$objForm.BackgroundImageLayout = "None"
$objForm.Text = "Excel в Image"
$objForm.StartPosition = "CenterScreen"
$objForm.Height = '160'
$objForm.Width = '270'


$TextSheetLabel = New-Object System.Windows.Forms.Label
$TextSheetLabel.Location = ('10,5')
$TextSheetLabel.AutoSize = $true
$TextSheetLabel.Text = "Выберите номер листа"
$TextSheetLabel.BackColor = "Transparent"
$TextSheet = New-Object System.Windows.Forms.ComboBox
$TextSheet.DataSource = $Numbers
$TextSheet.DropDownStyle = "DropDownList"
$TextSheet.Location = ('10,20')
$TextSheet.Autosize = $true

$ComboLabel = New-Object System.Windows.Forms.Label
$ComboLabel.Location = ('10,45')
$ComboLabel.AutoSize = $true
$ComboLabel.Text = "Выберите расширение"
$ComboLabel.BackColor = "Transparent"

$ComboSheet = New-Object System.Windows.Forms.ComboBox
$ComboSheet.Location = ('10,60')
$ComboSheet.DataSource = $Formats
$ComboSheet.DropDownStyle = "DropDownList"

$ExcelButton = New-Object System.Windows.Forms.Button
$ExcelButton.Location = ('10,85')
$ExcelButton.AutoSize = $true 
$ExcelButton.Text = "Выберите Файл"
$ExcelButton.Add_Click({
 SelectFile
 ConvertExcelToImage -FilePath $Global:FilePath -SheetNumber $TextSheet.SelectedItem -Format $ComboSheet.SelectedItem
})
$objForm.Controls.AddRange(@($ComboSheet,$ComboLabel,$ExcelButton,$TextSheet,$TextSheetLabel))
$objForm.ShowDialog()

}

function ConvertExcelToImage()
{
param(
$FilePath,
$SheetNumber,
$Format
)
    

    $Excel = New-Object -ComObject Excel.Application
    #$Excel.Visible = $true
    $Workbook = $Excel.Workbooks.Open($FilePath)
    $InventoryFile = $WorkBook.Worksheets.Item($SheetNumber)
    $Range = $InventoryFile.UsedRange
    $Range.Copy()
    $Exl_Img  = Get-Clipboard -Format Image
    $Exl_Img.Save("C:\Image\Report2.$Format", $Format)   
}


#$FilePath = "C:\Test\Image.xlsx"
#$SheetNumber = 1
#$Format = "PNG"
#ConvertExcelToImage -FilePath $FilePath -SheetNumber $SheetNumber -Format $Format






Function SelectFile {
$DestinationPoint = [Environment]::GetFolderPath("Desktop")
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            InitialDirectory = $DestinationPoint
            Filter = 'Файл Excel (*.xlsx)|*.xlsx|Все файлы |*.*'
            Title = 'Выберите файл EXCEL для конвертации'}
            $FileBrowser.ShowDialog()

$TST = $FileBrowser.SafeFileName
$Path_To_File = $FileBrowser.FileName
if($TST -eq '')
{ 
    [System.Windows.Forms.MessageBox]::Show("Не выбран файл!");
    return
}
$Global:FilePath = $FileBrowser.FileName
}

Hide-Console
General





function Resize-Image
{
   
    Param([Parameter(Mandatory=$true)][string]$InputFile, [string]$OutputFile, [int32]$Width, [int32]$Height, [int32]$Scale, [Switch]$Display)

    # Add System.Drawing assembly
    Add-Type -AssemblyName System.Drawing

    # Open image file
    $img = [System.Drawing.Image]::FromFile((Get-Item $InputFile))

    # Define new resolution
    if($Width -gt 0) { [int32]$new_width = $Width }
    elseif($Scale -gt 0) { [int32]$new_width = $img.Width * ($Scale / 100) }
    else { [int32]$new_width = $img.Width / 2 }
    if($Height -gt 0) { [int32]$new_height = $Height }
    elseif($Scale -gt 0) { [int32]$new_height = $img.Height * ($Scale / 100) }
    else { [int32]$new_height = $img.Height / 2 }

    # Create empty canvas for the new image
    $img2 = New-Object System.Drawing.Bitmap($new_width, $new_height)

    # Draw new image on the empty canvas
    $graph = [System.Drawing.Graphics]::FromImage($img2)
    $graph.DrawImage($img, 0, 0, $new_width, $new_height)

    # Create window to display the new image
    if($Display)
    {
        Add-Type -AssemblyName System.Windows.Forms
        $win = New-Object Windows.Forms.Form
        $box = New-Object Windows.Forms.PictureBox
        $box.Width = $new_width
        $box.Height = $new_height
        $box.Image = $img2
        $win.Controls.Add($box)
        $win.AutoSize = $true
        $win.ShowDialog()
    }

    # Save the image
    if($OutputFile -ne "")
    {
        $img2.Save($OutputFile);
    }
}


#Resize-Image -InputFile "C:\Image\Report2.PNG" -Width 1920 -Height 1080 -Display 