##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjofG5iZk2UD9fW4kZcyVhZKo04+w8OvoqBnQSpUaT115kizuO220VfcBavMauNUURyI5KvMZ4brvCPOmV6MNl9wuM73Y/+FxW1Pb7PM=
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
##L8/UAdDXTlaDjofG5iZk2UD9fW4kZcyVhZKi14qo8PrQmA38arFUZGdH2CzkASs=
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
#curl URL -d "" -X POST GET
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

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




Function MSP(){

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$VRQ = 'http://uat3.sp.mvideo.ru/stocks/rest/search'
$VRX = 'http://uatx.sp.mvideo.ru/stocks/rest/search'

#STANDART CALL IMAGE + FONT
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$Font = New-Object System.Drawing.Font("Comic Sans MS",9,[System.Drawing.FontStyle]::Bold)
$FontRES = New-Object System.Drawing.Font("ARIAL",9,[System.Drawing.FontStyle]::Regular)
$Image =  [system.drawing.image]::FromFile("\\dubovenko\D\SOFT\wallapers\RESTFORM.jpg")
$FontBanksy = New-Object System.Drawing.Font("Tempus Sans ITC",8,[System.Drawing.FontStyle]::Bold)

if (-not ([System.Management.Automation.PSTypeName]"TrustAllCertsPolicy").Type)
{
    Add-Type -TypeDefinition  @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem)
    {
        return true;
    }
}
"@
}

if ([System.Net.ServicePointManager]::CertificatePolicy.ToString() -ne "TrustAllCertsPolicy")
    {
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    }


#FORM RTD INITIALIZE
    $MSP_FORM = New-Object System.Windows.Forms.Form
    $MSP_FORM.SizeGripStyle = "Hide"
    $MSP_FORM.BackgroundImage = $Image
    $MSP_FORM.BackgroundImageLayout = "None"
    $MSP_FORM.Width = '400'
    $MSP_Form.Height = '320'
    $MSP_FORM.StartPosition = "CenterScreen"
    $MSP_FORM.Font = $Font
    $MSP_FORM.TopMost = $True
    $MSP_FORM.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $MSP_FORM.Text = "Проверка остатков MSP"
    $MSP_FORM.TopMost = $True
    $MSP_FORM.Icon = $Icon

#FORm FOR INPUT
    $MSP_VALUE = New-Object System.Windows.Forms.TextBox
    $MSP_VALUE.Location = New-Object System.Drawing.Point('10','20')
    $MSP_VALUE.Size = ('140,25')
    $MSP_VALUE.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $MSP_VALUE.Text = ''
    $MSP_VALUE.Font = $Font
    $MSP_VALUE.SelectionStart = '0'

    $MSP_VALUE.Add_DoubleClick({
        $MSP_VALUE.Text = ''
    })
#LABEL FOR INPUT
    $MSP_GROUP = New-Object System.Windows.Forms.GroupBox
    $MSP_GROUP.Location = '10,40'
    $MSP_GROUP.size = '160,55'
    $MSP_GROUP.Font = $Font
    $MSP_GROUP.text = "КОД ТОВАРА:"
    $MSP_GROUP.Backcolor = 'Transparent'
    $MSP_GROUP.Controls.Add($MSP_VALUE)

    $FOKINLAB = New-Object System.Windows.Forms.Label
    $FOKINLAB.Location = ('290,260')
    $FOKINLAB.Text = "Created By Fokin"
    $FOKINLAB.Font = $FontBanksy
    $FOKINLAB.BackColor =  'Transparent'
    $FOKINLAB.AutoSize = $True


    $MSP_RESULT = New-Object System.Windows.Forms.TextBox
    $MSP_RESULT.Location = ('10,110')
    $MSP_RESULT.Size = ('360,150')
    $MSP_RESULT.Name = 'RESULT'
    $MSP_RESULT.Multiline = $True
    $MSP_RESULT.Font = $FontRes
    $MSP_RESULT.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D


    $MSP_Buton = New-Object System.Windows.Forms.Button
    $MSP_Buton.Location = ('300,45')
    $MSP_Buton.AutoSize = $False
    $MSP_Buton.Size = '80,60'
    $MSP_Buton.Text = 'Проверить остатки'
    $MSP_Buton.Font = $Font


    $MSP_SHOP = New-Object System.Windows.Forms.TextBox
    $MSP_SHOP.Location = ('10,20')
    $MSP_SHOP.size = ('100,25')
    $MSP_SHOP.Name = 'S'
    $MSP_SHOP.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D

    $MSP_GROUP2 = New-Object System.Windows.Forms.GroupBox
    $MSP_GROUP2.Location = '175,40'
    $MSP_GROUP2.size = '120,55'
    $MSP_GROUP2.Font = $Font
    $MSP_GROUP2.text = "МАГАЗИН:"
    $MSP_GROUP2.Backcolor = 'Transparent'
    $MSP_GROUP2.Controls.Add($MSP_SHOP)


$BodyStock = @{

   "RequestBody" = @{
        "materials" =@(
            "20021413",
            "20021414"
        );
        "stockObjects"=@(
            "C0111",
            "C025"
        );
        "objectGroups"=@(
            "S543"
        );
        "key" = "stockCondition"
        "value" = "inStorage"

        }

} | ConvertTo-Json

$Headers = @{
   "Content-Type"="application/json" 
}

                                #"20021413",
                                #"20021414",
                                #10000301


$PWD = ConvertTo-SecureString "FOBO" -AsPlainText -Force
$CRED = New-Object Management.Automation.PSCredential ('FOBO', $PWD)

    $MSP_Buton.Add_Click({
        $MSP_RESULT.Text = ''
        [string]$MAG = "S"+$MSP_SHOP.Text
        $Result = $null
        $MATERIAL = @($MSP_VALUE.Text) 
        $R = $null
                $BodyStock = @{
                        "RequestBody" = @{
                        "materials" = @(
                                $MSP_VALUE.Text
                                 );   
                        "stockObjects"=@(
                                 "C0111",
                                 "C025"
                                    );
                        "objectGroups"=@(
                                 $MAG
                                    );
                        "key" = "stockCondition"
                        "value" = "inStorage"
                                }
                                    } | ConvertTo-Json


        if($MSP_VALUE.Text -ne $null)
        {
            if($MSP_RADIO1.Checked -eq $True)
            { 
                $URL = $VRQ 
            }
            elseif($MSP_RADIO2.Checked -eq $True) 
            { 
                $URL = $VRX
            }
            #[System.Windows.Forms.MessageBox]::Show("Запуск проверки остатков НА $URL","MSP",'OK','WARNING')    
            $Result = Invoke-RestMethod -Method 'Post' -Uri $URL -Body $BodyStock -Credential $CRED -Headers $Headers
            foreach($R in $Result.ResponseBody.stocks)
            {
                if($R.ObjectId -like "C*")
                { $MSP_RESULT.AppendText("S:"+$R.ObjectId + " - " )

                  $MSP_RESULT.AppendText(" Z: "+ ($R.stockparams.Value -replace ("inStorage",""))  + " - ") 
                }
                else
                {
                 $MSP_RESULT.AppendText("M: "+$R.ObjectId + " - ")

                 $MSP_RESULT.AppendText(" Z: "+($R.stockparams.Value -replace ("inStorage",""))  + " - ") 
                }  
                if($R.stocklevels.Material -eq $null){
                  $MSP_RESULT.AppendText("`r`n")
                }
                else
                { $MSP_RESULT.AppendText("MAT: " + $R.stocklevels.Material +" OST: " + $R.stocklevels.stock +"`r`n") } 
  
            }
        }
       $Result = $null   
    })

    $MSP_RADIO1 = New-Object System.Windows.Forms.RadioButton
    $MSP_RADIO1.Location = ('10,15')
    $MSP_RADIO1.Text = "VRQ"
    $MSP_RADIO1.Checked = $True
    $MSP_RADIO1.Size = ('50,20')
    $MSP_RADIO1.BackColor = 'LightBlue'
    $MSP_RADIO1.Font = $Font
    

    $MSP_RADIO2 = New-Object System.Windows.Forms.RadioButton
    $MSP_RADIO2.Location = ('60,15')
    $MSP_RADIO2.Text = "VRX"
    $MSP_RADIO2.Checked = $False
    $MSP_RADIO2.Size = ('50,20')
    $MSP_RADIO2.BackColor = 'LightBlue'
    $MSP_RADIO2.Font = $Font
    
    $MSP_GROUP3 = New-Object System.Windows.Forms.GroupBox
    $MSP_GROUP3.Location = '10,0'
    $MSP_GROUP3.size = '120,40'
    $MSP_GROUP3.Font = $Font
    $MSP_GROUP3.text = "СРЕДА:"
    $MSP_GROUP3.Backcolor = 'Transparent'
    $MSP_GROUP3.Controls.AddRange(@($MSP_RADIO1,$MSP_RADIO2))

$MSP_FORM.Controls.Add($FOKINLAB)
$MSP_FORM.Controls.Add($MSP_GROUP)   
$MSP_FORM.Controls.Add($MSP_GROUP2)
$MSP_FORM.Controls.Add($MSP_GROUP3)
$MSP_FORM.Controls.AddRange(@($MSP_Buton,$MSP_RESULT))

$MSP_FORM.TopMost = $true
$MSP_FORM.Add_Shown({$MSP_FORM.Activate()})
$MSP_FORM.ShowDialog()

}
Hide-Console
MSP

