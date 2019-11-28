$BodyStock1 = @{

              
   "RequestBody" = @{
        "materials" = @(
            "20021413"
            "20021414"
        );
        "stockObjects"= @(
            "C017",
            "C095"
        );
        "objectGroups"= @(
            "S002"
        );
        "stockParams"= @{
                "key" = "stockCondition";
                "value"= "inStorage";
                        }
        }
} | ConvertTo-Json

$BodyStock = @{

   "RequestBody" = @{
        "materials" =@(
            "10000301"
            
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

$Cred = Get-Credential
$Result = Invoke-RestMethod -Method 'Post' -Uri "http://uat3.sp.mvideo.ru/stocks/rest/search" -Body $BodyStock -Headers $Headers -Credential $Cred


$Result.ResponseBody.stocks
$Result.ResponseBody.stocks.Id
$Result.ResponseBody.stocks.ObjectID
$Result.ResponseBody.stocks.stockparams.Value
$Result.ResponseBody.stocks.stockLevels
