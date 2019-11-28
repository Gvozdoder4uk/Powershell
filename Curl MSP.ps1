#curl URL -d "" -X POST GET
$VRQ = 'http://uat3.sp.mvideo.ru/stocks/rest/search'
$VRX = 'http://uatx.sp.mvideo.ru/stocks/rest/search'
#STANDART CALL IMAGE + FONT
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Bold)
$Image =  [system.drawing.image]::FromFile("\\dubovenko\D\SOFT\wallapers\M.VIDEO1.jpg")

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
    $MSP_FORM.Width = '300'
    $MSP_FORM.Height = $Image.Height
    $MSP_FORM.StartPosition = "CenterScreen"
    $MSP_FORM.TopMost = $True
    $MSP_FORM.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $MSP_FORM.Text = "Проверка остатков MSP"
    $MSP_FORM.TopMost = $True
    $MSP_FORM.Icon = $Icon

#FORm FOR INPUT
    $MSP_VALUE = New-Object System.Windows.Forms.TextBox
    $MSP_VALUE.Location = New-Object System.Drawing.Point('10','40')
    $MSP_VALUE.Size = ('220,25')
    $MSP_VALUE.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $MSP_VALUE.Text = '              ВВЕДИТЕ КОД ТОВАРА'
    $MSP_VALUE.Font = $Font
    $MSP_VALUE.SelectionStart = '0'

    $MSP_VALUE.Add_DoubleClick({
        $MSP_VALUE.Text = ''
    })


    $MSP_RESULT = New-Object System.Windows.Forms.TextBox
    $MSP_RESULT.Location = ('10,100')
    $MSP_RESULT.Size = ('280,100')
    $MSP_RESULT.Name = 'RESULT'
    $MSP_RESULT.Multiline = $True


    $MSP_Buton = New-Object System.Windows.Forms.Button
    $MSP_Buton.Location = ('55,65')
    $MSP_Buton.AutoSize = $True
    $MSP_Buton.Text = 'Проверить остатки'
    $MSP_Buton.Font = $Font

    $body = @(
    
    )



$Body = @{

  "RequestBody" = '';
  "operationType" = "MRC";
  "srcStock" = "Main";
  "dstStock"= "TradeFloor";
  "workflowId"= 131041113;
  "workflowStepId"= 1;
  "refDocNum"= 62361;
  "refDocType"= "CMS";
  "employee"= 102549;
  "labelEmployee"= 102549;
  "isMRC"= "Y";
  "positionList" = @{
      "positionId"= 1;
      "material"= 10000180;
      "custGoodsId"= "00140000000040";
      "quantity"= 1;
      "windowSale"= "Y";
      "reservationCode"= "e6619569-4387-431d-a084-ca76f57066d5";
      "srcStockLocation"= 13;
      "dstStockLocation"= 13;
      "refDocPosition"=1;
      "webOrder"= 1000051620;
      "webOrderDetail"= 2;
      }
}




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

$PWD = ConvertTo-SecureString "FOBO" -AsPlainText -Force
$CRED = New-Object Management.Automation.PSCredential ('FOBO', $PWD)

    $MSP_Buton.Add_Click({
        $MSP_RESULT.Text = ''
        $Result = $null
        $R = $null
                $BodyStock = @{
                        "RequestBody" = @{
                        "materials" =@(
                                #"20021413",
                                #"20021414",
                                #10000301
                                $MSP_VALUE.Text
                                    );
                        "stockObjects"=@(
                                 "C0111",
                                 "C025"
                                    );
                        "objectGroups"=@(
                                 "S465"
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

                  $MSP_RESULT.AppendText(" L:"+$R.stockparams.Value  + "  ") 
                  
                if($R.stocklevels.Material -eq $null){
                  $MSP_RESULT.AppendText("`r`n")
                }
                else
                { $MSP_RESULT.AppendText("MAT:" + $R.stocklevels.Material +" OST:" + $R.stocklevels.stock +"`r`n") } }
                else
                {<#$MSP_RESULT.AppendText("M:"+$R.ObjectId + " - " )#>}
                
                
                
            }
        }
       $Result = $null   
    })

    $MSP_RADIO1 = New-Object System.Windows.Forms.RadioButton
    $MSP_RADIO1.Location = ('60,10')
    $MSP_RADIO1.Text = "VRQ"
    $MSP_RADIO1.Checked = $True
    $MSP_RADIO1.Size = ('50,20')
    $MSP_RADIO1.BackColor = 'Transparent'
    $MSP_RADIO1.Font = $Font
    

    $MSP_RADIO2 = New-Object System.Windows.Forms.RadioButton
    $MSP_RADIO2.Location = ('110,10')
    $MSP_RADIO2.Text = "VRX"
    $MSP_RADIO2.Checked = $False
    $MSP_RADIO2.Size = ('50,20')
    $MSP_RADIO2.BackColor = 'Transparent'
    $MSP_RADIO2.Font = $Font
    

$MSP_FORM.Controls.Add($MSP_RADIO1)
$MSP_FORM.Controls.Add($MSP_RADIO2)
$MSP_FORM.Controls.AddRange(@($MSP_VALUE,$MSP_Buton,$MSP_RESULT))
$MSP_FORM.ShowDialog()
$Result = $null 